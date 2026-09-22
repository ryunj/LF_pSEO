#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pSEO 대시보드 데이터 빌드

입력
  1) 실적 RAW  : PV·UV 크로스탭 (UTF-16 탭 구분)
                 행 = 지표(PV/UV) x AF코드, 열 = 유입 연월일(YYYYMMDD)
  2) 콘텐츠 대장: pSEO_전체대장_*.xlsx  ('전체대장' 시트)
                 차수 · 키워드 · AF코드 · slug · 현재 운영중 · 업데이트 날짜 ·
                 최신 반영 · 미반영 분류 · 미반영 사유
                 → AF코드 ↔ 키워드 매핑의 기준(구글시트 PSEO 탭보다 우선)
  3) (선택) mapping_seed.tsv : 대장에 없는 코드용 보조 매핑

출력
  data.js  : window.PSEO_DATA = {...}

사용법
  python3 build_data.py <실적CSV> [대장XLSX] [출력경로]
"""
import csv, glob, json, os, re, sys
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))

# 브랜드는 긴 이름부터 매칭 (예: '질스튜어트 뉴욕'이 '질스튜어트'보다 먼저)
BRANDS = [
    "닥스 런던 골프", "위크엔드 막스마라", "질 바이 질스튜어트", "질스튜어트 뉴욕",
    "막스마라 선글라스", "아일랜드 슬리퍼", "오피신 제네랄", "포르테 포르테",
    "티톤 브로스", "이에르 로르", "더블플래그", "바네사브루노", "이자벨마랑",
    "마에스트로", "프리미아타", "바이이에르", "티엔지티", "일꼬르소", "알레그리",
    "레오나드", "이누이키", "막스마라", "앳코너", "우포스", "던스트", "헤지스",
    "닥스", "리복", "라움", "바버", "바쉬", "빈스", "빠투", "아떼", "킨",
]
GENDERS = ["남성", "여성", "공용", "아동"]


def split_name(name):
    """키워드 -> (유형, 브랜드, 성별, 카테고리)"""
    s = re.sub(r"\s*추천\s*$", "", (name or "").strip())
    brand = ""
    for b in BRANDS:
        if s == b or s.startswith(b + " "):
            brand = b
            s = s[len(b):].strip()
            break
    gender = ""
    for g in GENDERS:
        if s.startswith(g + " "):
            gender, s = g, s[len(g):].strip()
            break
        if s.startswith(g) and len(s) > len(g):   # '여성지갑' 처럼 붙여 쓴 경우
            gender, s = g, s[len(g):].strip()
            break
    return ("br" if brand else "cat"), brand, gender, (s.strip() or "기타")


def ds(v):
    if v is None:
        return ""
    if hasattr(v, "strftime"):
        return v.strftime("%Y-%m-%d")
    return str(v).strip()[:10]


# ---------------------------------------------------------------- 대장 읽기
def read_ledger(path):
    import openpyxl
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True)["전체대장"]
    rows = list(ws.iter_rows(values_only=True))
    head = [str(h or "").strip() for h in rows[0]]
    I = {h: i for i, h in enumerate(head)}
    need = ["차수", "키워드", "AF코드", "slug", "현재 운영중", "업데이트 날짜",
            "최신 반영", "미반영 분류", "미반영 사유"]
    for n in need:
        if n not in I:
            raise SystemExit("대장에서 '%s' 열을 찾지 못했습니다: %s" % (n, head))
    out = {}
    for r in rows[1:]:
        code = str(r[I["AF코드"]] or "").strip()
        if not code.startswith("PS"):
            continue
        out[code] = {
            "ch":   str(r[I["차수"]] or "").strip(),
            "n":    re.sub(r"\s*추천\s*$", "", str(r[I["키워드"]] or "").strip()),
            "slug": str(r[I["slug"]] or "").strip(),
            "live": str(r[I["현재 운영중"]] or "").strip(),
            "upd":  ds(r[I["업데이트 날짜"]]),
            "new":  str(r[I["최신 반영"]] or "").strip(),
            "miss": str(r[I["미반영 분류"]] or "").strip(),
            "why":  str(r[I["미반영 사유"]] or "").strip(),
        }
    return out


def read_seed():
    p = os.path.join(HERE, "mapping_seed.tsv")
    out = {}
    if os.path.exists(p):
        for line in open(p, encoding="utf-8"):
            if line.strip():
                n, name = line.rstrip("\n").split("\t")
                out["PSBRD" + n] = re.sub(r"\s*추천\s*$", "", name)
    return out


# ---------------------------------------------------------------- 실적 읽기
def read_perf(path):
    raw = open(path, "rb").read()
    text = None
    for enc in ("utf-16", "utf-8-sig", "cp949"):
        try:
            t = raw.decode(enc)
            if "\t" in t or "," in t:
                text = t
                break
        except UnicodeDecodeError:
            continue
    if text is None:
        raise SystemExit("실적 파일 인코딩을 읽지 못했습니다: " + path)

    lines = [l for l in text.splitlines() if l.strip("\t, ")]
    rows = [l.split("\t") if "\t" in l else next(csv.reader([l])) for l in lines]

    hi, best = 0, -1
    for i, r in enumerate(rows[:8]):
        n = sum(1 for c in r if re.fullmatch(r"\d{8}", (c or "").strip()))
        if n > best:
            hi, best = i, n
    cols = [(j, c.strip()) for j, c in enumerate(rows[hi]) if re.fullmatch(r"\d{8}", (c or "").strip())]
    if not cols:
        raise SystemExit("날짜(YYYYMMDD) 열을 찾지 못했습니다.")

    data, totals = {}, {}
    for r in rows[hi + 1:]:
        met = (r[0] or "").strip().lower()
        code = (r[1] or "").strip() if len(r) > 1 else ""
        if met not in ("pv", "uv"):
            continue
        for j, d in cols:
            v = (r[j] or "").strip().replace(",", "") if j < len(r) else ""
            if not v:
                continue
            try:
                n = int(round(float(v)))
            except ValueError:
                continue
            if not n:
                continue
            if code == "총계":
                totals.setdefault(d, {"pv": 0, "uv": 0})[met] += n
            elif code.startswith("PS"):
                data.setdefault((d, code), {"pv": 0, "uv": 0})[met] += n
    return data, totals


# ---------------------------------------------------------------- 빌드
def build(perf_path, ledger_path, out_path):
    led = read_ledger(ledger_path) if ledger_path else {}
    seed = read_seed()
    perf, totals = read_perf(perf_path)

    seen = {c for (_, c) in perf}
    all_codes = sorted(set(led) | seen, key=lambda c: int(re.sub(r"\D", "", c)))

    idx, codes, orphan = {}, [], []
    for c in all_codes:
        L = led.get(c)
        if L:
            name, src = L["n"], "ledger"
        else:
            name, src = seed.get(c, c), ("sheet" if c in seed else "none")
            orphan.append(c)
        t, b, g, k = split_name(name)
        if c.startswith("PSCAT"):
            t = "cat"
        rec = {"c": c, "n": name, "t": t, "b": b, "g": g, "k": k, "src": src,
               "ch": (L or {}).get("ch", ""), "slug": (L or {}).get("slug", ""),
               "live": (L or {}).get("live", ""), "upd": (L or {}).get("upd", ""),
               "new": (L or {}).get("new", ""), "miss": (L or {}).get("miss", ""),
               "why": (L or {}).get("why", "")}
        idx[c] = len(codes)
        codes.append(rec)

    rows = [[f"{d[:4]}-{d[4:6]}-{d[6:]}", idx[c], v["pv"], v["uv"]]
            for (d, c), v in sorted(perf.items())]
    dates = sorted({r[0] for r in rows})

    data = {
        "asof": dates[-1] if dates else "",
        "from": dates[0] if dates else "",
        "built": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "source": {"perf": os.path.basename(perf_path),
                   "map": os.path.basename(ledger_path) if ledger_path else "mapping_seed.tsv"},
        "metrics": ["pv", "uv"],
        "codes": codes,
        "rows": rows,
        "total": {f"{d[:4]}-{d[4:6]}-{d[6:]}": [v["pv"], v["uv"]] for d, v in sorted(totals.items())},
    }
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("window.PSEO_DATA = " + json.dumps(data, ensure_ascii=False, separators=(",", ":")) + ";\n")

    live = sum(1 for c in codes if c["live"] == "Y")
    print(f"대장          : {os.path.basename(ledger_path) if ledger_path else '-'} ({len(led):,}건)")
    print(f"코드 합계     : {len(codes):,}건 (대장 밖 {len(orphan)}건: {', '.join(orphan[:8])})")
    print(f"현재 운영중   : {live:,}건")
    print(f"실적 코드     : {len(seen):,}건 / 기간 {data['from']} ~ {data['asof']} ({len(dates)}일)")
    print(f"출력          : {out_path} ({os.path.getsize(out_path):,} bytes)")
    return data


if __name__ == "__main__":
    if len(sys.argv) < 2:
        raise SystemExit("사용법: python3 build_data.py <실적CSV> [대장XLSX] [data.js]")
    perf = sys.argv[1]
    ledger = sys.argv[2] if len(sys.argv) > 2 else (sorted(glob.glob(os.path.join(HERE, "*전체대장*.xlsx"))) or [None])[-1]
    out = sys.argv[3] if len(sys.argv) > 3 else os.path.join(HERE, "data.js")
    build(perf, ledger, out)
