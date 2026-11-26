"""
pSEO 키워드 검색량 + 전월/이번달 트렌드 일괄 조회

1) 네이버 검색광고 API (keywordstool) : 최근 30일 월간 검색수 (PC / 모바일)
   - 1회 호출에 키워드 5개 → 1,090개면 218회 자동 반복
2) 네이버 데이터랩 검색어트렌드 API : 전월 vs 이번달 일평균 검색 비율
   - 1회 호출에 키워드 5개 → 218회 (일 한도 1,000회 이내)
   - ratio는 같은 호출 안에서만 비교 가능한 상대값이라
     '같은 키워드의 전월 대비 이번달 증감률'로만 사용

필요 환경변수 (GitHub Secrets):
  NAVER_AD_API_KEY, NAVER_AD_SECRET, NAVER_AD_CUSTOMER_ID
  NAVER_CLIENT_ID, NAVER_CLIENT_SECRET
"""
import base64
import calendar
import hashlib
import hmac
import os
import sys
import time
from datetime import date, timedelta

import pandas as pd
import requests

KEYWORD_FILE = sys.argv[1] if len(sys.argv) > 1 else "keyword_research/keywords.txt"
OUT_XLSX = "output/pseo_keyword_volume_trend.xlsx"
BATCH = 5
TREND_THRESHOLD = 10.0  # ±10% 이상이면 상승/하락


def norm(s: str) -> str:
    return s.replace(" ", "").upper()


def chunks(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i + n]


def to_int(v):
    # 검색광고 API는 10 미만을 "< 10" 문자열로 줌
    if isinstance(v, (int, float)):
        return int(v)
    return 5 if "<" in str(v) else int(v or 0)


def request_with_retry(fn, tries=5):
    for t in range(tries):
        r = fn()
        if r.status_code == 200:
            return r
        if r.status_code in (429, 500, 502, 503, 504):
            time.sleep(2 ** t)
            continue
        raise RuntimeError(f"{r.status_code}: {r.text[:300]}")
    raise RuntimeError(f"재시도 초과: {r.status_code} {r.text[:300]}")


# ---------- 1) 검색광고 API: 월간 검색수 ----------
def ad_headers(method, uri):
    ts = str(int(time.time() * 1000))
    msg = f"{ts}.{method}.{uri}"
    sig = base64.b64encode(
        hmac.new(os.environ["NAVER_AD_SECRET"].encode(), msg.encode(), hashlib.sha256).digest()
    ).decode()
    return {
        "X-Timestamp": ts,
        "X-API-KEY": os.environ["NAVER_AD_API_KEY"],
        "X-Customer": os.environ["NAVER_AD_CUSTOMER_ID"],
        "X-Signature": sig,
    }


def fetch_volume(keywords):
    uri = "/keywordstool"
    result = {}
    total = (len(keywords) + BATCH - 1) // BATCH
    for i, batch in enumerate(chunks(keywords, BATCH), 1):
        hint = ",".join(norm(k) for k in batch)  # 공백 있으면 API가 거부함
        r = request_with_retry(lambda: requests.get(
            "https://api.searchad.naver.com" + uri,
            params={"hintKeywords": hint, "showDetail": "1"},
            headers=ad_headers("GET", uri), timeout=30))
        rows = {norm(x["relKeyword"]): x for x in r.json().get("keywordList", [])}
        for k in batch:
            x = rows.get(norm(k))
            if x:
                pc, mo = to_int(x["monthlyPcQcCnt"]), to_int(x["monthlyMobileQcCnt"])
                result[k] = {"PC검색수": pc, "모바일검색수": mo, "월간검색수": pc + mo,
                             "경쟁정도": x.get("compIdx")}
            else:
                result[k] = {"PC검색수": 0, "모바일검색수": 0, "월간검색수": 0, "경쟁정도": None}
        if i % 20 == 0 or i == total:
            print(f"[검색량] {i}/{total}", flush=True)
        time.sleep(0.25)
    return result


# ---------- 2) 데이터랩: 전월 vs 이번달 ----------
def fetch_trend(keywords):
    today = date.today()
    end = today - timedelta(days=1)  # 어제까지
    this_start = end.replace(day=1)
    prev_end = this_start - timedelta(days=1)
    prev_start = prev_end.replace(day=1)
    headers = {
        "X-Naver-Client-Id": os.environ["NAVER_CLIENT_ID"],
        "X-Naver-Client-Secret": os.environ["NAVER_CLIENT_SECRET"],
        "Content-Type": "application/json",
    }
    result = {}
    total = (len(keywords) + BATCH - 1) // BATCH
    for i, batch in enumerate(chunks(keywords, BATCH), 1):
        body = {
            "startDate": prev_start.isoformat(),
            "endDate": end.isoformat(),
            "timeUnit": "date",
            "keywordGroups": [{"groupName": k, "keywords": [k]} for k in batch],
        }
        r = request_with_retry(lambda: requests.post(
            "https://openapi.naver.com/v1/datalab/search", json=body, headers=headers, timeout=30))
        by_name = {g["title"]: g["data"] for g in r.json().get("results", [])}
        for k in batch:
            data = by_name.get(k, [])
            prev = [d["ratio"] for d in data if d["period"] <= prev_end.isoformat()]
            cur = [d["ratio"] for d in data if d["period"] >= this_start.isoformat()]
            # 데이터 없는 날은 응답에서 빠지므로 달력 일수로 나눠 일평균 계산
            prev_days = calendar.monthrange(prev_start.year, prev_start.month)[1]
            cur_days = (end - this_start).days + 1
            p = sum(prev) / prev_days
            c = sum(cur) / cur_days
            if p == 0 and c == 0:
                chg, label = None, "데이터없음"
            elif p == 0:
                chg, label = None, "신규발생"
            else:
                chg = round((c / p - 1) * 100, 1)
                label = "상승" if chg >= TREND_THRESHOLD else "하락" if chg <= -TREND_THRESHOLD else "보합"
            result[k] = {"전월_일평균지수": round(p, 3), "이번달_일평균지수": round(c, 3),
                         "증감률(%)": chg, "트렌드": label}
        if i % 20 == 0 or i == total:
            print(f"[트렌드] {i}/{total}", flush=True)
        time.sleep(0.15)
    period = f"전월 {prev_start}~{prev_end} / 이번달 {this_start}~{end}"
    return result, period


def main():
    with open(KEYWORD_FILE, encoding="utf-8") as f:
        keywords = list(dict.fromkeys(l.strip() for l in f if l.strip()))
    print(f"키워드 {len(keywords)}개", flush=True)

    vol = fetch_volume(keywords)
    trend, period = fetch_trend(keywords)

    df = pd.DataFrame([{"키워드": k, **vol[k], **trend[k]} for k in keywords])
    os.makedirs("output", exist_ok=True)
    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="결과", index=False)
        pd.DataFrame({
            "항목": ["조회일", "비교기간", "월간검색수", "트렌드 지수", "트렌드 판정"],
            "설명": [date.today().isoformat(), period,
                   "검색광고 API 최근 30일 PC+모바일 (10 미만은 5로 표기)",
                   "데이터랩 상대지수 — 같은 키워드 안에서 전월/이번달 비교용 (키워드끼리 비교 X)",
                   f"이번달 일평균 / 전월 일평균 −1, ±{TREND_THRESHOLD:.0f}% 기준 상승/보합/하락"],
        }).to_excel(w, sheet_name="기준", index=False)
    df.to_csv(OUT_XLSX.replace(".xlsx", ".csv"), index=False, encoding="utf-8-sig")
    print(df["트렌드"].value_counts().to_string())
    print("완료:", OUT_XLSX)


if __name__ == "__main__":
    main()
