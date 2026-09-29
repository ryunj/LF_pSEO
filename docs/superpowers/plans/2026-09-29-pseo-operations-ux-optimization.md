# pSEO Operations UX Optimization Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** pSEO 대시보드의 기존 지표 산식을 유지하면서 파일 갱신, 상태 확인, 백업·복원, Excel 내보내기, 반응형 표 사용을 비개발자도 안전하게 수행할 수 있게 만든다.

**Architecture:** `dashboard.html`은 화면 렌더링과 기존 지표 계산을 계속 담당한다. 파일 판별, 문자열 정규화, 백업 봉투 검증, 보기 상태 직렬화처럼 DOM에 의존하지 않는 로직은 새 `dashboard_logic.js`의 UMD 모듈로 분리해 브라우저와 Node 테스트에서 함께 사용하며, `app.py`는 이 로컬 스크립트와 SheetJS를 Streamlit 컴포넌트 안에 인라인한다.

**Tech Stack:** 정적 HTML/CSS/JavaScript, IndexedDB, localStorage, SheetJS(저장소 내 vendored 파일), Python `unittest`, Node.js `node:test`, Streamlit.

## Global Constraints

- 기존 PV·UV 산식은 `MET`, `agg()`, `val()`, `prevId()`를 기준으로 유지한다.
- 주차는 ISO 월요일~일요일이며 화면 라벨은 `몇월 몇주차`, 비교는 직전 기간만 사용한다.
- 일평균 분모는 적재된 일자 수이며 필터 때문에 유입이 0인 날도 분모에서 빼지 않는다.
- UV는 코드 합산값이며 필터가 없을 때 태블로 중복제거 총계를 함께 표시한다.
- 콘텐츠 대장의 `AF코드 ↔ 키워드`가 정규 기준이고 `매체코드명(NBOS등록)`은 동일 코드의 별칭이다.
- CSV, TSV, TXT, XLSX, XLSM, XLS 업로드를 지원하고 날짜 열이 없는 대장·AF 매핑 파일도 갱신할 수 있어야 한다.
- 새 업로드가 유효하지 않거나 적용 중 실패하면 현재 정상 데이터와 마지막 정상 백업을 덮어쓰지 않는다.
- 외부 CDN, Gemini, 메모·협업, 로그인, 서버 데이터베이스는 추가하지 않는다.
- 실적 원본, 대장, `data.js`, CSV, Excel 파일은 저장소에 커밋하지 않는다.
- 360px 폭에서도 핵심 버튼과 상태 문구가 잘리지 않고, 넓은 표는 고정 헤더·첫 열과 가로 스크롤을 제공한다.
- 데스크톱 흐름 차트는 세 칸을 기준으로 두 개만 렌더링하며 1024px 이하 두 칸, 700px 이하 한 칸으로 전환한다.

## File Structure

- Create: `dashboard_logic.js` — 업로드 판별, 정규화, 데이터 요약, 백업·보기 상태 스키마와 검증을 담당하는 DOM 비의존 UMD 모듈.
- Modify: `dashboard.html` — 상태 영역, 갱신 모달, 드롭 오버레이, 토스트, 백업 메뉴, 실제 XLSX 내보내기 및 UI 이벤트를 담당.
- Modify: `app.py` — `dashboard_logic.js`와 vendored SheetJS를 Streamlit HTML 안에 인라인.
- Modify: `tests/test_dashboard_contract.py` — HTML 구조, 접근성, 인라인 계약, 기존 지표 불변 계약 검증.
- Create: `tests/dashboard_logic.test.mjs` — 순수 JavaScript 로직의 실제 입력·출력 검증.
- Modify: `README.md` — 새 갱신·백업·내보내기 흐름과 제한 사항 문서화.
- Modify: `HANDOFF.md` — 운영자가 실행할 회귀 검증 체크리스트 갱신.

---

### Task 1: 테스트 가능한 데이터·저장 스키마 분리

**Files:**
- Create: `dashboard_logic.js`
- Create: `tests/dashboard_logic.test.mjs`
- Modify: `dashboard.html:1-8, 847-895`
- Modify: `app.py:24-35`
- Modify: `tests/test_dashboard_contract.py`

**Interfaces:**
- Consumes: 브라우저의 2차원 업로드 행 배열과 기존 `window.PSEO_DATA` 구조.
- Produces: `window.PSEOLogic` / CommonJS export에 `BACKUP_SCHEMA`, `VIEW_STATE_KEY`, `normalizeToken(value)`, `parseCSV(text)`, `findHeader(rows, names)`, `detectUploadKind(rows)`, `sheetPriority(kind, name)`, `summarizeData(data)`, `makeBackupEnvelope(data, viewState, savedAt)`, `readBackupEnvelope(value)`, `sanitizeViewState(value)`를 제공한다.

- [ ] **Step 1: 순수 로직의 실패 테스트 작성**

```javascript
// tests/dashboard_logic.test.mjs
import test from 'node:test';
import assert from 'node:assert/strict';
import {createRequire} from 'node:module';

const require = createRequire(import.meta.url);
const logic = require('../dashboard_logic.js');

test('대장과 AF 매핑은 날짜 열 없이도 파일 종류를 판별한다', () => {
  assert.equal(logic.detectUploadKind([['AF코드', '키워드'], ['PSBRD1', '닥스']]), 'led');
  assert.equal(logic.detectUploadKind([['매체코드', '매체코드명(NBOS등록)'], ['PSBRD1', '1_brand_닥스']]), 'map');
});

test('PSEO 및 전체대장 시트를 우선한다', () => {
  assert.equal(logic.sheetPriority('map', 'PSEO'), 100);
  assert.equal(logic.sheetPriority('led', '2026 전체대장'), 100);
  assert.equal(logic.sheetPriority('map', '설명'), 10);
});

test('데이터 요약은 기간과 AF 매칭 수를 계산한다', () => {
  const summary = logic.summarizeData({
    asof: '2026-08-29', from: '2026-07-10', built: '2026-09-29 11:00',
    codes: [{c:'PS1', src:'ledger'}, {c:'PS2', src:'none'}],
    rows: [['2026-07-10', 0, 3, 2], ['2026-08-29', 1, 1, 1]], total: {}
  });
  assert.deepEqual(summary, {
    asof:'2026-08-29', from:'2026-07-10', rowCount:2,
    matchedCodes:1, unmatchedCodes:1, built:'2026-09-29 11:00'
  });
});

test('호환되지 않는 백업은 거부하고 정상 백업은 복원한다', () => {
  assert.throws(() => logic.readBackupEnvelope({schema:1, data:{codes:[], rows:[]}}), /호환되지 않는/);
  const envelope = logic.makeBackupEnvelope({codes:[], rows:[], total:{}}, {grain:'week'}, 1234);
  assert.equal(envelope.schema, logic.BACKUP_SCHEMA);
  assert.equal(logic.readBackupEnvelope(envelope).savedAt, 1234);
});

test('보기 상태는 허용된 값만 복원한다', () => {
  assert.deepEqual(
    logic.sanitizeViewState({tab:'bad', grain:'week', agg:'sum', range:8, layout:'cmp', topn:'20',
      types:['br','bad'],gender:['여성'],brands:['닥스'],cats:['가디건'],lf:{c:'PS'},lq:'닥스',lsort:'c',ldir:-1}),
    {tab:'perf',grain:'week',agg:'sum',layout:'cmp',dim:'b',sort:'uv',topn:'20',range:8,
      at:null,types:['br'],gender:['여성'],brands:['닥스'],cats:['가디건'],lf:{c:'PS'},lq:'닥스',lsort:'c',ldir:-1}
  );
});
```

- [ ] **Step 2: 테스트가 모듈 부재로 실패하는지 확인**

Run: `node --test tests/dashboard_logic.test.mjs`

Expected: FAIL with `Cannot find module '../dashboard_logic.js'`.

- [ ] **Step 3: DOM 비의존 로직 모듈 작성**

```javascript
// dashboard_logic.js
(function(root, factory){
  const api = factory();
  if(typeof module === 'object' && module.exports) module.exports = api;
  else root.PSEOLogic = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function(){
  'use strict';
  const BACKUP_SCHEMA = 2;
  const VIEW_STATE_KEY = 'pseo-dashboard-view-v1';
  const ALLOWED = {
    tab:['perf','list'], grain:['month','week','day'], agg:['sum','avg'],
    layout:['cmp','plain'], dim:['b','k','g','n'], sort:['pv','uv'], topn:['10','20','50','9999']
  };
  const DEFAULT_VIEW={tab:'perf',grain:'week',agg:'sum',layout:'cmp',dim:'b',sort:'uv',topn:'10'};
  const LIST_SORTS=['n','c','ch','slug','live','upd','new','miss','why','_flow','_uv','_pv'];

  function normalizeToken(value){
    return String(value == null ? '' : value).normalize('NFKC').trim().toLowerCase().replace(/\s+/g, ' ');
  }
  function parseCSV(text){
    const rows=[];let row=[],field='',quoted=false;
    const sep=String(text).split('\n')[0].includes('\t')?'\t':',';
    for(let i=0;i<text.length;i++){
      const char=text[i];
      if(quoted){
        if(char==='"'){if(text[i+1]==='"'){field+='"';i++;}else quoted=false;}
        else field+=char;
      }else if(char==='"') quoted=true;
      else if(char===sep){row.push(field);field='';}
      else if(char==='\n'){row.push(field);rows.push(row);row=[];field='';}
      else if(char!=='\r') field+=char;
    }
    if(field||row.length){row.push(field);rows.push(row);}
    return rows;
  }
  function findHeader(rows,names){
    for(let hi=0;hi<Math.min(rows.length,12);hi++){
      const H=rows[hi].map(x=>String(x||'').trim());
      if(names.every(name=>H.includes(name))) return {hi,H};
    }
    return null;
  }
  function detectUploadKind(rows){
    if(findHeader(rows,['AF코드','키워드'])) return 'led';
    if(findHeader(rows,['매체코드','매체코드명(NBOS등록)'])||findHeader(rows,['매체코드'])) return 'map';
    return rows.slice(0,12).some(row=>row.some(cell=>/^\d{8}$/.test(String(cell||'').trim())))?'perf':'';
  }
  function sheetPriority(kind,name){
    name=String(name||'').trim();
    if(kind==='map'&&name==='PSEO') return 100;
    if(kind==='led'&&name.includes('전체대장')) return 100;
    return name?10:0;
  }
  function summarizeData(data){
    const codes=Array.isArray(data&&data.codes)?data.codes:[];
    const rows=Array.isArray(data&&data.rows)?data.rows:[];
    return {asof:data&&data.asof||'',from:data&&data.from||'',rowCount:rows.length,
      matchedCodes:codes.filter(code=>code.src==='ledger'||code.src==='sheet').length,
      unmatchedCodes:codes.filter(code=>code.src!=='ledger'&&code.src!=='sheet').length,
      built:data&&data.built||''};
  }
  function sanitizeViewState(value){
    value=value&&typeof value==='object'?value:{};
    const out={};
    for(const [key,values] of Object.entries(ALLOWED)){
      const candidate=String(value[key]??'');
      out[key]=values.includes(candidate)?candidate:DEFAULT_VIEW[key];
    }
    out.range=[3,6,8,12,13,14,26,30,60,90].includes(Number(value.range))?Number(value.range):26;
    out.at=typeof value.at==='string'?value.at:null;
    out.types=Array.isArray(value.types)?value.types.filter(item=>['br','cat','cep'].includes(item)):[];
    for(const key of ['gender','brands','cats']) out[key]=Array.isArray(value[key])?value[key].filter(item=>typeof item==='string'):[];
    out.lf=value.lf&&typeof value.lf==='object'?Object.fromEntries(Object.entries(value.lf).filter(([,item])=>typeof item==='string')):{};
    out.lq=typeof value.lq==='string'?value.lq:'';
    out.lsort=LIST_SORTS.includes(value.lsort)?value.lsort:'n';
    out.ldir=value.ldir===-1?-1:1;
    return out;
  }
  function makeBackupEnvelope(data,viewState,savedAt){
    return {schema:BACKUP_SCHEMA,savedAt:savedAt??Date.now(),data,view:sanitizeViewState(viewState),summary:summarizeData(data)};
  }
  function readBackupEnvelope(value){
    if(!value||value.schema!==BACKUP_SCHEMA) throw new Error('호환되지 않는 백업 형식입니다.');
    if(!value.data||!Array.isArray(value.data.codes)||!Array.isArray(value.data.rows)||typeof value.data.total!=='object')
      throw new Error('백업 데이터가 손상되었습니다.');
    return value;
  }
  return {BACKUP_SCHEMA,VIEW_STATE_KEY,normalizeToken,parseCSV,findHeader,detectUploadKind,
    sheetPriority,summarizeData,makeBackupEnvelope,readBackupEnvelope,sanitizeViewState};
});
```

- [ ] **Step 4: HTML과 Streamlit에서 모듈을 로드하도록 연결**

```html
<!-- dashboard.html: SheetJS 다음 -->
<script src="dashboard_logic.js"></script>
```

```python
# app.py: xlsx 로드 옆
logic = read("dashboard_logic.js")
if logic:
    html = html.replace(
        '<script src="dashboard_logic.js"></script>',
        "<script>%s</script>" % logic,
    )
```

`dashboard.html`의 기존 `parseCSV`, `findHeader`, `detectUploadKind`, `sheetPriority` 본문은 제거하고 아래 별칭을 사용한다.

```javascript
const {parseCSV,findHeader,detectUploadKind,sheetPriority} = window.PSEOLogic;
```

- [ ] **Step 5: 로직·계약 테스트 실행**

Run: `node --test tests/dashboard_logic.test.mjs`

Expected: 5 tests PASS.

Run: `uv run --with openpyxl python -m unittest discover -s tests -v`

Expected: existing 15 tests PASS after contract expectations are updated to accept the new local script and `PSEOLogic` aliases.

- [ ] **Step 6: 커밋**

```bash
git add dashboard_logic.js dashboard.html app.py tests/dashboard_logic.test.mjs tests/test_dashboard_contract.py
git commit -m "refactor: isolate dashboard data logic"
```

---

### Task 2: 적용 전 검증이 있는 데이터 갱신 모달

**Files:**
- Modify: `dashboard.html:37-158, 167-170, 224-235, 896-1055`
- Modify: `tests/test_dashboard_contract.py`
- Modify: `tests/dashboard_logic.test.mjs`

**Interfaces:**
- Consumes: Task 1의 `PSEOLogic.detectUploadKind`, `findHeader`, `sheetPriority`, `normalizeToken`.
- Produces: `UPLOAD_SESSION:{files:Array, selections:Object, draft:Object|null, summary:Object|null}`, `escapeHtml(value):string`, `stageFiles(files):Promise<void>`, `buildDraft(selections):Object`, `applyDraft():Promise<void>`, `renderUploadLog():void`.

- [ ] **Step 1: 모달 구조와 원자적 적용 계약의 실패 테스트 작성**

```python
# tests/test_dashboard_contract.py에 추가
def test_refresh_uses_review_modal_before_atomic_apply(self):
    for token in ('id="refreshDialog"', 'id="refreshDrop"', 'id="refreshFiles"',
                  'id="refreshSummary"', 'id="btnApplyRefresh"'):
        self.assertIn(token, DASHBOARD)
    self.assertRegex(DASHBOARD, r"async function stageFiles\(files\)")
    self.assertRegex(DASHBOARD, r"function buildDraft\(selections\)")
    self.assertRegex(DASHBOARD, r"async function applyDraft\(\)")
    self.assertLess(DASHBOARD.index("const previous=D"), DASHBOARD.index("D=UPLOAD_SESSION.draft.data"))
    self.assertNotIn("alert(", DASHBOARD)

def test_refresh_reports_every_selected_file(self):
    self.assertIn("fileName:f.name", DASHBOARD)
    self.assertIn("validRows", DASHBOARD)
    self.assertIn("skippedRows", DASHBOARD)
    self.assertIn("unmatchedCodes", DASHBOARD)
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `uv run --with openpyxl python -m unittest tests.test_dashboard_contract.DashboardContractTests.test_refresh_uses_review_modal_before_atomic_apply -v`

Expected: FAIL because `refreshDialog` is absent.

- [ ] **Step 3: 접근 가능한 갱신 모달 마크업 추가**

```html
<dialog id="refreshDialog" aria-labelledby="refreshTitle">
  <form method="dialog" class="refresh-modal">
    <div class="modal-head">
      <div><h2 id="refreshTitle">데이터 갱신</h2><p>CSV 또는 Excel 파일을 확인한 뒤 적용합니다.</p></div>
      <button class="icon-btn" value="cancel" aria-label="닫기">×</button>
    </div>
    <button type="button" id="refreshDrop" class="drop-zone">
      파일을 여기에 놓거나 눌러서 선택하세요
      <small>CSV · TSV · TXT · XLSX · XLSM · XLS</small>
    </button>
    <div id="refreshFiles" class="file-log" aria-live="polite"></div>
    <section id="refreshSummary" class="apply-summary" hidden></section>
    <div class="modal-actions">
      <button class="btn" value="cancel">취소</button>
      <button class="btn primary" type="button" id="btnApplyRefresh" disabled>적용</button>
    </div>
  </form>
</dialog>
```

- [ ] **Step 4: 파일별 분석과 적용 전 초안 생성 구현**

`loadFiles(files)`의 즉시 반영 코드를 `stageFiles`, `buildDraft`, `applyDraft`로 나눈다. `buildDraft`는 기존 대장 우선·NBOS 별칭·실적 합산 코드를 그대로 사용하되 전역 `D`를 쓰지 않고 아래 형태만 반환한다.

```javascript
const UPLOAD_SESSION={files:[],selections:{perf:[],map:null,led:null},draft:null,summary:null,lastFocus:null};

async function stageFiles(files){
  UPLOAD_SESSION.files=[];UPLOAD_SESSION.selections={perf:[],map:null,led:null};
  for(const f of files){
    const log={fileName:f.name,format:(f.name.split('.').pop()||'').toUpperCase(),sheet:'',kind:'',readRows:0,validRows:0,skippedRows:0,error:''};
    try{
      const sheets=await readUploadSheets(f),selected={};
      for(const item of sheets){
        const kind=detectUploadKind(item.rows);
        if(!kind) continue;
        const priority=sheetPriority(kind,item.sheet);
        if(!selected[kind]||priority>selected[kind].priority) selected[kind]={...item,priority};
      }
      const choices=Object.entries(selected);
      if(!choices.length) throw new Error('지원하는 열을 찾지 못했습니다.');
      for(const [kind,item] of choices){
        const sourceName=f.name+(item.sheet?' ['+item.sheet+']':'');
        const selection={kind,rows:item.rows,sourceName,fileName:f.name};
        const counts=countCandidateRows(selection);
        if(kind==='perf') UPLOAD_SESSION.selections.perf.push(selection);
        else UPLOAD_SESSION.selections[kind]=selection;
        log.kind=kind;log.sheet=item.sheet;log.readRows=item.rows.length;
        log.validRows+=counts.validRows;log.skippedRows+=counts.skippedRows;
      }
    }catch(error){log.error=error.message||'파일을 읽지 못했습니다.';}
    UPLOAD_SESSION.files.push(log);
  }
  try{
    UPLOAD_SESSION.draft=buildDraft(UPLOAD_SESSION.selections);
    UPLOAD_SESSION.summary=UPLOAD_SESSION.draft.summary;
  }catch(error){
    UPLOAD_SESSION.draft=null;UPLOAD_SESSION.summary={error:error.message};
  }
  renderUploadLog();
}

function countCandidateRows(selection){
  const {kind,rows}=selection;
  const required=kind==='led'?['AF코드','키워드']:kind==='map'?['매체코드']:[];
  if(required.length){
    const found=findHeader(rows,required);if(!found)return {validRows:0,skippedRows:rows.length};
    const codeIndex=found.H.indexOf(kind==='led'?'AF코드':'매체코드');
    const body=rows.slice(found.hi+1),validRows=body.filter(row=>String(row[codeIndex]||'').trim().startsWith('PS')).length;
    return {validRows,skippedRows:body.length-validRows};
  }
  let headerIndex=-1,dateColumns=[];
  rows.slice(0,12).some((row,index)=>{dateColumns=row.map((cell,column)=>/^\d{8}$/.test(String(cell||'').trim())?column:-1).filter(column=>column>=0);if(dateColumns.length){headerIndex=index;return true;}return false;});
  if(headerIndex<0)return {validRows:0,skippedRows:rows.length};
  const body=rows.slice(headerIndex+1),validRows=body.filter(row=>row.some((cell,column)=>dateColumns.includes(column)&&String(cell||'').trim()!=='')).length;
  return {validRows,skippedRows:body.length-validRows};
}

function buildDraft(selections){
  const result=calculateUploadedData(selections, FILE_INPUT, D);
  if(!result.data.codes.length) throw new Error('적용 가능한 AF 코드가 없습니다.');
  const summary=window.PSEOLogic.summarizeData(result.data);
  summary.validRows=result.validRows;
  summary.skippedRows=result.skippedRows;
  summary.duplicateRows=result.duplicateRows;
  summary.unmatchedCodes=result.data.codes.filter(code=>code.src!=='ledger'&&code.src!=='sheet').length;
  return {data:result.data,summary};
}

async function applyDraft(){
  if(!UPLOAD_SESSION.draft) return;
  const previous=D;
  try{
    const next=UPLOAD_SESSION.draft;
    await saveBackup(previous);
    D=next.data;setBaseData(D);prep();S.at=null;syncSelectors();srcNote();draw();
    await saveBackup(D);
    showToast('데이터를 갱신했습니다.','success');
    document.getElementById('refreshDialog').close();
  }catch(error){
    D=previous;setBaseData(D);prep();syncSelectors();srcNote();draw();
    showBanner('error','갱신을 적용하지 못해 이전 데이터를 유지했습니다.');
  }
}
```

`calculateUploadedData` 안에서 여러 실적 파일은 정규화된 `파일 종류|행 전체` 지문으로 완전히 같은 행만 한 번 처리한다. 날짜가 없는 대장·매핑 파일은 `perf`로 강제하지 않으며, 실적 파일에서 날짜 열이 없으면 그 파일만 오류로 기록하고 다른 정상 파일의 적용을 막지 않는다.

AF 별칭 인덱스는 원문 키와 `normalizeToken` 키를 함께 저장하고 실적의 `rawCode`도 같은 방식으로 조회한다. 이로써 앞뒤 공백, Unicode 호환 문자, 영문 대소문자, 연속 공백 차이는 같은 코드로 처리하며 원문 코드와 NBOS 명칭은 결과 레코드에 그대로 보존한다.

```javascript
function calculateUploadedData(selections,fileInput,current){
  const codes=fileInput.baseCodes.map(code=>({...code})),byCode={},aliases={},mapByCode={};
  const aliasKey=value=>window.PSEOLogic.normalizeToken(value);
  const addAlias=(value,code)=>{if(value){aliases[String(value).trim()]=code;aliases[aliasKey(value)]=code;}};
  const resolveAlias=value=>aliases[String(value).trim()]||aliases[aliasKey(value)]||String(value).trim();
  let validRows=0,skippedRows=0,duplicateRows=0;
  codes.forEach((code,index)=>{byCode[code.c]=index;addAlias(code.c,code.c);addAlias(code.nbos,code.c);});

  if(selections.led){
    const {rows}=selections.led,found=findHeader(rows,['AF코드','키워드']);
    if(!found)throw new Error('대장에서 AF코드와 키워드 열을 찾지 못했습니다.');
    const {hi,H}=found,index=name=>H.indexOf(name),codeIndex=index('AF코드'),nameIndex=index('키워드');
    const optional={ch:index('차수'),slug:index('slug'),live:index('현재 운영중'),upd:index('업데이트 날짜'),
      nw:index('최신 반영'),miss:index('미반영 분류'),why:index('미반영 사유')};
    const get=(row,column)=>column>=0?String(row[column]||'').trim():'';
    for(const row of rows.slice(hi+1)){
      const code=get(row,codeIndex);if(!code.startsWith('PS')){skippedRows++;continue;}
      const name=get(row,nameIndex).replace(/\s*추천\s*$/,''),parts=splitName(name);
      const record={c:code,n:name,t:code.startsWith('PSCAT')?'cat':parts.t,b:parts.b,g:parts.g,k:parts.k,src:'ledger',
        ch:get(row,optional.ch),slug:get(row,optional.slug),live:get(row,optional.live),upd:get(row,optional.upd).slice(0,10),
        new:get(row,optional.nw),miss:get(row,optional.miss),why:get(row,optional.why)};
      if(byCode[code]==null){byCode[code]=codes.length;codes.push(record);}else codes[byCode[code]]=record;
      addAlias(code,code);validRows++;
    }
  }

  if(selections.map){
    const {rows}=selections.map,found=findHeader(rows,['매체코드','매체코드명(NBOS등록)']);
    if(!found)throw new Error('AF 매핑에서 매체코드와 NBOS 등록명 열을 찾지 못했습니다.');
    const {hi,H}=found,codeIndex=H.indexOf('매체코드'),sheetIndex=H.indexOf('매체코드명(시트기재)'),
      nbosIndex=H.indexOf('매체코드명(NBOS등록)'),utmIndex=H.indexOf('utm_source');
    for(const row of rows.slice(hi+1)){
      const code=String(row[codeIndex]||'').trim();if(!code.startsWith('PS')){skippedRows++;continue;}
      const sheetName=sheetIndex>=0?String(row[sheetIndex]||'').trim():'',nbos=String(row[nbosIndex]||'').trim();
      const name=cleanMediaName(sheetName||nbos),parts=splitName(name),previous=byCode[code]==null?{}:codes[byCode[code]];
      addAlias(code,code);addAlias(sheetName,code);addAlias(nbos,code);if(utmIndex>=0)addAlias(row[utmIndex],code);
      const rawAliases=[sheetName,nbos,utmIndex>=0?String(row[utmIndex]||'').trim():''].filter(Boolean);
      const mapped={c:code,n:name,t:code.startsWith('PSCAT')?'cat':parts.t,b:parts.b,g:parts.g,k:parts.k,src:'sheet',nbos,aliases:rawAliases};
      mapByCode[code]=mapped;
      if(byCode[code]!=null)codes[byCode[code]]=previous.src==='ledger'?{...previous,nbos,aliases:rawAliases}:{...previous,...mapped};
      validRows++;
    }
  }

  let rows=fileInput.baseRows.map(row=>row.slice()),total={...fileInput.baseTotal};
  if(selections.perf.length){
    const seen=new Set(),acc={},totals={};
    for(const selection of selections.perf){
      const table=selection.rows;let headerIndex=-1,dateColumns=[];
      table.slice(0,12).some((row,index)=>{dateColumns=row.map((cell,column)=>/^\d{8}$/.test(String(cell||'').trim())?[column,String(cell).trim()]:null).filter(Boolean);if(dateColumns.length){headerIndex=index;return true;}return false;});
      if(headerIndex<0){skippedRows+=table.length;continue;}
      const header=table[headerIndex].map(cell=>String(cell||'').trim()),codeColumn=header.indexOf('AF코드');
      const metricMatch=selection.sourceName.match(/(?:^|[_\-\s])(pv|uv)(?=[_\-\s.\[]|$)/i),singleMetric=metricMatch?metricMatch[1].toLowerCase():'';
      for(const row of table.slice(headerIndex+1)){
        const fingerprint=JSON.stringify(row.map(cell=>String(cell||'').trim()));
        if(seen.has(fingerprint)){duplicateRows++;continue;}seen.add(fingerprint);
        const metric=singleMetric&&codeColumn>=0?singleMetric:String(row[0]||'').trim().toLowerCase();
        const rawCode=singleMetric&&codeColumn>=0?row[codeColumn]:row[1],code=resolveAlias(rawCode);
        if(!['pv','uv'].includes(metric)){skippedRows++;continue;}
        let used=false;
        for(const [column,date] of dateColumns){
          const number=Math.round(Number(String(row[column]||'').replace(/,/g,'')));if(!number)continue;
          const day=date.slice(0,4)+'-'+date.slice(4,6)+'-'+date.slice(6);used=true;
          if(code==='총계'||code==='총합계'){(totals[day]=totals[day]||{pv:0,uv:0})[metric]+=number;continue;}
          if(!code.startsWith('PS'))continue;
          if(byCode[code]==null){byCode[code]=codes.length;codes.push(mapByCode[code]||{c:code,n:String(rawCode||code),t:code.startsWith('PSCAT')?'cat':'br',b:'',g:'',k:'미매핑',src:'none'});}
          const key=day+'|'+byCode[code];(acc[key]=acc[key]||{pv:0,uv:0})[metric]+=number;
        }
        if(used)validRows++;else skippedRows++;
      }
    }
    rows=Object.entries(acc).map(([key,value])=>{const [day,index]=key.split('|');return [day,Number(index),value.pv,value.uv];}).sort((a,b)=>a[0].localeCompare(b[0]));
    total={};for(const day of Object.keys(totals).sort())total[day]=[totals[day].pv,totals[day].uv];
  }
  const dates=[...new Set(rows.map(row=>row[0]).filter(Boolean))].sort(),perfNames=selections.perf.map(item=>item.sourceName).join(' + ');
  const data={asof:dates.at(-1)||'',from:dates[0]||'',built:new Date().toISOString().slice(0,16).replace('T',' '),
    source:{perf:perfNames||fileInput.baseSource.perf||'–',map:[selections.led?.sourceName,selections.map?.sourceName].filter(Boolean).join(' + ')||fileInput.baseSource.map||'–'},
    metrics:['pv','uv'],codes,rows,total};
  return {data,validRows,skippedRows,duplicateRows};
}
```

- [ ] **Step 5: 파일 로그와 요약 렌더링 구현**

```javascript
function escapeHtml(value){
  return String(value??'').replace(/[&<>"']/g,char=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char]));
}
function renderUploadLog(){
  const host=document.getElementById('refreshFiles');
  host.innerHTML=UPLOAD_SESSION.files.map(file=>
    '<details class="file-row"'+(file.error?' open':'')+'><summary><b>'+escapeHtml(file.fileName)+'</b><span>'+
    (file.error?'오류':(file.kind||'확인'))+'</span></summary><p>'+
    (file.error?escapeHtml(file.error):escapeHtml(file.sheet||'단일 시트')+' · '+nf(file.readRows)+'행')+'</p></details>'
  ).join('');
  const summary=document.getElementById('refreshSummary');
  const value=UPLOAD_SESSION.summary;
  summary.hidden=!value;
  summary.innerHTML=value?(value.error?'<b>'+escapeHtml(value.error)+'</b>':
    '<b>적용 가능 '+nf(value.validRows)+'행</b><span>'+value.from+' ~ '+value.asof+
    ' · 중복 '+nf(value.duplicateRows)+' · 제외 '+nf(value.skippedRows)+' · 미매칭 코드 '+nf(value.unmatchedCodes)+'</span>'):'';
  document.getElementById('btnApplyRefresh').disabled=!UPLOAD_SESSION.draft;
}
```

- [ ] **Step 6: 테스트 실행 및 커밋**

Run: `node --test tests/dashboard_logic.test.mjs`

Expected: all tests PASS.

Run: `uv run --with openpyxl python -m unittest discover -s tests -v`

Expected: all tests PASS.

```bash
git add dashboard.html tests/test_dashboard_contract.py tests/dashboard_logic.test.mjs
git commit -m "feat: add reviewed data refresh workflow"
```

---

### Task 3: 버전형 자동 백업과 수동 백업 관리

**Files:**
- Modify: `dashboard.html:167-170, 225-235, 899-947`
- Modify: `dashboard_logic.js`
- Modify: `tests/dashboard_logic.test.mjs`
- Modify: `tests/test_dashboard_contract.py`

**Interfaces:**
- Consumes: Task 1의 `makeBackupEnvelope`, `readBackupEnvelope`, `sanitizeViewState`.
- Produces: `saveBackup(data):Promise<boolean>`, `loadBackup():Promise<Envelope|null>`, `clearBackup():Promise<void>`, `downloadBlob(blob,name):void`, `downloadBackup():void`, `importBackup(file):Promise<void>`, `restoreLastBackup():Promise<void>`.

- [ ] **Step 1: 백업 보존과 수동 관리 계약 테스트 작성**

```python
# tests/test_dashboard_contract.py에 추가
def test_backup_management_is_versioned_and_user_controllable(self):
    self.assertIn('id="btnBackup"', DASHBOARD)
    self.assertIn('id="backupMenu"', DASHBOARD)
    self.assertIn('id="backupImport"', DASHBOARD)
    for name in ('downloadBackup', 'importBackup', 'clearBackup'):
        self.assertRegex(DASHBOARD, rf"(?:async )?function {name}\(")
    self.assertIn("makeBackupEnvelope(data,S)", DASHBOARD)
    self.assertIn("readBackupEnvelope(saved)", DASHBOARD)
```

- [ ] **Step 2: 실패 확인**

Run: `uv run --with openpyxl python -m unittest tests.test_dashboard_contract.DashboardContractTests.test_backup_management_is_versioned_and_user_controllable -v`

Expected: FAIL because `btnBackup` is absent.

- [ ] **Step 3: 백업 메뉴와 가져오기 입력 추가**

```html
<div class="backup-tools">
  <button class="btn" type="button" id="btnBackup" aria-expanded="false" aria-controls="backupMenu">백업</button>
  <div class="backup-menu" id="backupMenu" hidden>
    <button type="button" data-backup="download">현재 데이터 내려받기</button>
    <button type="button" data-backup="restore">마지막 백업 복원</button>
    <button type="button" data-backup="import">백업 파일 불러오기</button>
    <button type="button" data-backup="clear" class="danger">브라우저 백업 초기화</button>
  </div>
  <input type="file" id="backupImport" accept="application/json,.json" hidden>
</div>
```

- [ ] **Step 4: IndexedDB 저장값을 버전형 봉투로 교체**

```javascript
async function saveBackup(data){
  try{
    const db=await openBackupDb();if(!db){showBanner('warn','이 브라우저에서는 자동 백업을 사용할 수 없습니다.');return false;}
    const envelope=window.PSEOLogic.makeBackupEnvelope(data,S);
    await new Promise((resolve,reject)=>{
      const tx=db.transaction(BACKUP_STORE,'readwrite');
      tx.objectStore(BACKUP_STORE).put({key:BACKUP_KEY,...envelope});
      tx.oncomplete=resolve;tx.onerror=()=>reject(tx.error);tx.onabort=()=>reject(tx.error);
    });
    db.close();return true;
  }catch(error){showBanner('warn','자동 백업에 실패했습니다. 현재 화면의 데이터는 유지됩니다.');return false;}
}

async function clearBackup(){
  const db=await openBackupDb();if(!db)return;
  await new Promise((resolve,reject)=>{
    const tx=db.transaction(BACKUP_STORE,'readwrite');
    tx.objectStore(BACKUP_STORE).delete(BACKUP_KEY);
    tx.oncomplete=resolve;tx.onerror=()=>reject(tx.error);
  });
  db.close();showToast('브라우저 백업을 초기화했습니다.','success');
}

function downloadBackup(){
  const envelope=window.PSEOLogic.makeBackupEnvelope(D,S);
  downloadBlob(new Blob([JSON.stringify(envelope,null,2)],{type:'application/json'}),
    'pseo_backup_'+(D.asof||ymd(new Date()))+'.json');
}

function downloadBlob(blob,name){
  const url=URL.createObjectURL(blob),anchor=document.createElement('a');
  anchor.href=url;anchor.download=name;anchor.click();
  setTimeout(()=>URL.revokeObjectURL(url),0);
}

async function importBackup(file){
  const envelope=window.PSEOLogic.readBackupEnvelope(JSON.parse(await file.text()));
  D=envelope.data;Object.assign(S,envelope.view);setBaseData(D);
  prep();syncSelectors();srcNote();draw();await saveBackup(D);
  showToast('백업 파일을 복원했습니다.','success');
}
```

`restoreLastBackup`은 `readBackupEnvelope(saved)` 검증을 통과한 값만 적용한다. 기본 데이터보다 백업 `savedAt`이 최신일 때 자동 복원하고, 손상·구버전이면 기본 데이터를 유지하며 경고 배너를 표시한다.

- [ ] **Step 5: 초기화 확인과 포커스 복귀 연결**

```javascript
document.getElementById('backupMenu').onclick=async event=>{
  const action=event.target.closest('[data-backup]')?.dataset.backup;if(!action)return;
  if(action==='download') downloadBackup();
  if(action==='restore') await restoreLastBackup(true);
  if(action==='import') document.getElementById('backupImport').click();
  if(action==='clear'&&confirm('브라우저에 저장된 마지막 정상 백업을 초기화할까요?')) await clearBackup();
  document.getElementById('btnBackup').focus();
};
```

- [ ] **Step 6: 테스트 실행 및 커밋**

Run: `node --test tests/dashboard_logic.test.mjs`

Expected: all tests PASS, including corrupt and version-mismatch envelopes.

Run: `uv run --with openpyxl python -m unittest discover -s tests -v`

Expected: all tests PASS.

```bash
git add dashboard.html dashboard_logic.js tests/dashboard_logic.test.mjs tests/test_dashboard_contract.py
git commit -m "feat: add versioned dashboard backup controls"
```

---

### Task 4: 데이터 상태, 화면 내 메시지, 드롭 안내, 보기 상태 복원

**Files:**
- Modify: `dashboard.html:15-158, 159-235, 788-836, 1048-1055`
- Modify: `tests/test_dashboard_contract.py`
- Modify: `tests/dashboard_logic.test.mjs`

**Interfaces:**
- Consumes: `PSEOLogic.summarizeData(D)`, `PSEOLogic.VIEW_STATE_KEY`, `PSEOLogic.sanitizeViewState(value)`.
- Produces: `renderDataStatus()`, `showBanner(level,message)`, `showToast(message,level)`, `saveViewState()`, `restoreViewState()`, 전역 `draw()` 이후 상태 동기화.

- [ ] **Step 1: 상태·피드백·보기 복원 계약 테스트 작성**

```python
# tests/test_dashboard_contract.py에 추가
def test_dashboard_exposes_non_blocking_status_feedback(self):
    for token in ('id="dataStatus"', 'id="statusBanner"', 'id="toastRegion"', 'id="dropOverlay"'):
        self.assertIn(token, DASHBOARD)
    self.assertIn('aria-live="polite"', DASHBOARD)
    self.assertIn('role="alert"', DASHBOARD)
    self.assertNotIn("alert(", DASHBOARD)

def test_view_state_is_saved_and_restored(self):
    self.assertRegex(DASHBOARD, r"function saveViewState\(\)")
    self.assertRegex(DASHBOARD, r"function restoreViewState\(\)")
    self.assertIn("localStorage.setItem", DASHBOARD)
    self.assertIn("localStorage.getItem", DASHBOARD)
```

- [ ] **Step 2: 실패 확인**

Run: `uv run --with openpyxl python -m unittest tests.test_dashboard_contract.DashboardContractTests.test_dashboard_exposes_non_blocking_status_feedback -v`

Expected: FAIL because `dataStatus` is absent.

- [ ] **Step 3: 상태 영역과 피드백 컨테이너 추가**

```html
<section class="data-status" id="dataStatus" aria-label="현재 데이터 상태"></section>
<div class="status-banner" id="statusBanner" role="alert" hidden></div>
<div class="toast-region" id="toastRegion" aria-live="polite" aria-atomic="true"></div>
<div class="drop-overlay" id="dropOverlay" hidden><div><b>파일을 놓아 갱신 준비</b><span>확인 후 적용할 수 있습니다.</span></div></div>
```

- [ ] **Step 4: 상태 및 메시지 렌더링 구현**

```javascript
function renderDataStatus(){
  const value=window.PSEOLogic.summarizeData(D);
  const source=D.source&&D.source.backup?'자동 복원':(D.source&&D.source.perf&&D.source.perf!=='–'?'업로드':'기본 데이터');
  document.getElementById('dataStatus').innerHTML=[
    ['기준일',value.asof||'–'],['기간',(value.from||'–')+' ~ '+(value.asof||'–')],
    ['데이터',nf(value.rowCount)+'행'],['AF 코드',nf(value.matchedCodes)+' 매칭 · '+nf(value.unmatchedCodes)+' 미매칭'],
    ['마지막 갱신',value.built||'–'],['출처',source]
  ].map(item=>'<span><small>'+item[0]+'</small><b>'+item[1]+'</b></span>').join('');
}

function showBanner(level,message){
  const host=document.getElementById('statusBanner');host.className='status-banner '+level;
  host.textContent=message;host.hidden=false;
}
function showToast(message,level='success'){
  const toast=document.createElement('div');toast.className='toast '+level;toast.textContent=message;
  document.getElementById('toastRegion').appendChild(toast);
  setTimeout(()=>toast.remove(),4000);
}
```

`draw()` 끝에서 `renderDataStatus();saveViewState();`를 호출하고, 앱 시작 시 `restoreViewState()`를 `restoreLastBackup()`보다 먼저 호출한다.

- [ ] **Step 5: 보기 상태 저장과 드래그 오버레이 구현**

```javascript
function saveViewState(){
  try{localStorage.setItem(window.PSEOLogic.VIEW_STATE_KEY,JSON.stringify(window.PSEOLogic.sanitizeViewState(S)));}catch(error){}
}
function restoreViewState(){
  try{Object.assign(S,window.PSEOLogic.sanitizeViewState(JSON.parse(localStorage.getItem(window.PSEOLogic.VIEW_STATE_KEY)||'{}')));}catch(error){}
}
let dragDepth=0;
document.addEventListener('dragenter',event=>{event.preventDefault();dragDepth++;document.getElementById('dropOverlay').hidden=false;});
document.addEventListener('dragleave',event=>{event.preventDefault();dragDepth--;if(dragDepth<=0){dragDepth=0;document.getElementById('dropOverlay').hidden=true;}});
document.addEventListener('drop',event=>{
  event.preventDefault();dragDepth=0;document.getElementById('dropOverlay').hidden=true;
  const files=[...event.dataTransfer.files];if(files.length)openRefreshDialog(files);
});
```

`prep()` 이후에는 저장된 `types`, `gender`, `brands`, `cats`, `at`을 현재 `DIMS`와 기간 목록의 교집합으로 제한한다. 교집합이 비면 해당 차원의 전체 기본값을 사용해 오래된 필터 때문에 빈 화면이 고착되지 않게 한다.

- [ ] **Step 6: 테스트 실행 및 커밋**

Run: `node --test tests/dashboard_logic.test.mjs`

Expected: all tests PASS.

Run: `uv run --with openpyxl python -m unittest discover -s tests -v`

Expected: all tests PASS.

```bash
git add dashboard.html tests/test_dashboard_contract.py tests/dashboard_logic.test.mjs
git commit -m "feat: show durable dashboard data status"
```

---

### Task 5: 실제 XLSX 내보내기

**Files:**
- Modify: `dashboard.html:1057-1064`
- Modify: `tests/test_dashboard_contract.py`

**Interfaces:**
- Consumes: vendored 전역 `XLSX`, 현재 필터 상태 `S`, `listRows()`, 현재 렌더링된 기간 표.
- Produces: `exportWorkbook(scope):void`, `tableToAoA(table):Array<Array<string|number>>`.

- [ ] **Step 1: 가짜 XLS 제거와 실제 워크북 계약 테스트 작성**

```python
# tests/test_dashboard_contract.py에 추가
def test_excel_export_builds_a_real_xlsx_workbook(self):
    self.assertNotIn("application/vnd.ms-excel", DASHBOARD)
    self.assertNotIn(".xls';", DASHBOARD)
    self.assertIn("XLSX.utils.book_new()", DASHBOARD)
    self.assertIn("XLSX.utils.aoa_to_sheet", DASHBOARD)
    self.assertIn("XLSX.writeFile", DASHBOARD)
    self.assertIn("전체 데이터", DASHBOARD)
    self.assertIn("현재 보기", DASHBOARD)
```

- [ ] **Step 2: 실패 확인**

Run: `uv run --with openpyxl python -m unittest tests.test_dashboard_contract.DashboardContractTests.test_excel_export_builds_a_real_xlsx_workbook -v`

Expected: FAIL because `application/vnd.ms-excel` is still present.

- [ ] **Step 3: 표와 원본 데이터를 실제 워크북으로 변환**

```javascript
function tableToAoA(table){
  return [...table.rows].map(row=>[...row.cells].map(cell=>{
    const text=cell.innerText.trim().replace(/,/g,'');
    return /^-?\d+(?:\.\d+)?$/.test(text)?Number(text):cell.innerText.trim();
  }));
}
function exportWorkbook(scope){
  if(typeof XLSX==='undefined'){showBanner('error','Excel 내보내기 모듈을 불러오지 못했습니다.');return;}
  const book=XLSX.utils.book_new();
  const visible=document.querySelector(scope==='list'?'#tblList table':'#tblPeriod table');
  XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet(tableToAoA(visible)),'현재 보기');
  const all=[['날짜','AF코드','페이지명','PV','UV']];
  for(const row of ROWS){const code=CODE[row.i]||{};all.push([row.d,code.c||'',code.n||'',row.pv,row.uv]);}
  XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet(all),'전체 데이터');
  XLSX.writeFile(book,'pseo_'+(scope==='list'?'페이지리스트_':'')+(D.asof||ymd(new Date()))+'.xlsx',{compression:true});
  showToast('Excel 파일을 만들었습니다.','success');
}
```

페이지 리스트 내보내기는 `listRows()` 결과로 별도 `전체 데이터` 시트를 만들고, 날짜 문자열 셀은 `yyyy-mm-dd` 형태를 유지한다.

- [ ] **Step 4: 버튼 연결 및 테스트**

```javascript
document.getElementById('btnXls').onclick=()=>exportWorkbook(S.tab==='list'?'list':'perf');
document.getElementById('btnListXls').onclick=()=>exportWorkbook('list');
```

Run: `uv run --with openpyxl python -m unittest discover -s tests -v`

Expected: all tests PASS.

- [ ] **Step 5: 커밋**

```bash
git add dashboard.html tests/test_dashboard_contract.py
git commit -m "feat: export real xlsx workbooks"
```

---

### Task 6: 반응형 표, 모달 접근성, 빈·부분·오류 상태 마감

**Files:**
- Modify: `dashboard.html:31-158, 224-301, 527-709`
- Modify: `tests/test_dashboard_contract.py`

**Interfaces:**
- Consumes: Task 2 모달, Task 4 배너·토스트·상태 영역.
- Produces: `openRefreshDialog(files)`, `closeRefreshDialog()`, 모달 포커스 트랩, `.tbl-wrap`, `.grid thead th`, `.grid th.stick`, 360px 미디어 규칙.

- [ ] **Step 1: 반응형·접근성 계약 테스트 작성**

```python
# tests/test_dashboard_contract.py에 추가
def test_tables_and_refresh_modal_remain_usable_on_small_screens(self):
    self.assertRegex(DASHBOARD, r"@media \(max-width:420px\)")
    self.assertRegex(DASHBOARD, r"\.tbl-wrap\{[^}]*overflow:auto")
    self.assertRegex(DASHBOARD, r"\.grid thead th\{[^}]*position:sticky")
    self.assertRegex(DASHBOARD, r"\.grid th\.stick\{[^}]*position:sticky")
    self.assertIn("trapDialogFocus", DASHBOARD)
    self.assertIn("UPLOAD_SESSION.lastFocus.focus()", DASHBOARD)
```

- [ ] **Step 2: 실패 확인**

Run: `uv run --with openpyxl python -m unittest tests.test_dashboard_contract.DashboardContractTests.test_tables_and_refresh_modal_remain_usable_on_small_screens -v`

Expected: FAIL because the 420px modal rules and focus trap are absent.

- [ ] **Step 3: 표 겹침과 작은 화면 CSS 정리**

```css
.tbl-wrap{position:relative;overflow:auto;max-height:74vh;overscroll-behavior:contain;border-radius:10px}
.grid thead th{position:sticky;top:0;z-index:3;background:var(--surface)}
.grid th.stick{position:sticky;left:0;z-index:2;background:var(--surface);box-shadow:1px 0 0 var(--grid)}
.grid thead th.stick{z-index:5}
dialog{width:min(680px,calc(100vw - 32px));max-height:min(760px,calc(100vh - 32px));padding:0;border:1px solid var(--border);border-radius:14px;background:var(--surface);color:var(--ink)}
.refresh-modal{display:grid;grid-template-rows:auto auto minmax(80px,1fr) auto auto;max-height:inherit}
.file-log{overflow:auto;min-height:80px}
@media (max-width:420px){
  .wrap{padding-inline:12px}.bar1{gap:6px 12px}.brand{width:100%}.menu{order:3;width:100%}
  .bk-tools{margin-left:auto}.data-status{grid-template-columns:repeat(2,minmax(0,1fr))}
  dialog{width:calc(100vw - 16px);max-height:calc(100vh - 16px)}
  .modal-actions{position:sticky;bottom:0;background:var(--surface)}
}
```

- [ ] **Step 4: 모달 포커스와 상태별 빈 화면 구현**

```javascript
function trapDialogFocus(event){
  if(event.key!=='Tab')return;
  const dialog=document.getElementById('refreshDialog');
  const focusable=[...dialog.querySelectorAll('button:not([disabled]),input:not([disabled]),select:not([disabled]),[tabindex="0"]')];
  if(!focusable.length)return;
  const first=focusable[0],last=focusable[focusable.length-1];
  if(event.shiftKey&&document.activeElement===first){event.preventDefault();last.focus();}
  if(!event.shiftKey&&document.activeElement===last){event.preventDefault();first.focus();}
}
function openRefreshDialog(files=[]){
  UPLOAD_SESSION.lastFocus=document.activeElement;
  const dialog=document.getElementById('refreshDialog');dialog.showModal();dialog.addEventListener('keydown',trapDialogFocus);
  document.getElementById('refreshDrop').focus();if(files.length)stageFiles(files);
}
function closeRefreshDialog(){
  const dialog=document.getElementById('refreshDialog');dialog.removeEventListener('keydown',trapDialogFocus);
  if(dialog.open)dialog.close();if(UPLOAD_SESSION.lastFocus)UPLOAD_SESSION.lastFocus.focus();
}
```

빈 데이터는 `emptyState`에서 `데이터 갱신`을 바로 제공한다. 일부 행 제외는 경고 배너, 파일 전체 오류는 모달 로그, IndexedDB 미지원은 백업 비활성 안내, 화면 재계산 실패는 이전 데이터 유지 메시지로 각각 분기한다.

- [ ] **Step 5: 자동 테스트 실행**

Run: `node --test tests/dashboard_logic.test.mjs`

Expected: all tests PASS.

Run: `uv run --with openpyxl python -m unittest discover -s tests -v`

Expected: all tests PASS.

- [ ] **Step 6: 커밋**

```bash
git add dashboard.html tests/test_dashboard_contract.py
git commit -m "fix: harden responsive dashboard states"
```

---

### Task 7: 운영 문서, 통합 검증, 배포 산출물

**Files:**
- Modify: `README.md`
- Modify: `HANDOFF.md`
- Modify: `tests/test_dashboard_contract.py`
- Verify only: `dashboard.html`, `dashboard_logic.js`, `app.py`, `vendor/xlsx.full.min.js`
- Generate without committing: `C:/Users/USER/Documents/Codex/2026-09-28/https-github-com-ryunj-lf-pseo/outputs/pseo_dashboard_fixed.html`

**Interfaces:**
- Consumes: Tasks 1-6의 완성된 업로드·백업·내보내기·반응형 흐름.
- Produces: 운영자가 그대로 따라 할 수 있는 사용 설명, 전체 회귀 테스트 결과, 로컬 스크립트가 모두 포함된 단일 HTML 산출물.

- [ ] **Step 1: 기존 지표와 로컬 의존성 불변 계약 추가**

```python
# tests/test_dashboard_contract.py에 추가
def test_core_metric_and_period_rules_remain_intact(self):
    for token in ("const MET=[", "function agg(grain)", "function val(o,k)",
                  "function prevId(p,grain)", "function weekLabel(id)"):
        self.assertIn(token, DASHBOARD)
    self.assertNotIn("cdn.sheetjs.com", DASHBOARD)
    self.assertIn('<script src="dashboard_logic.js"></script>', DASHBOARD)
    self.assertIn('read("dashboard_logic.js")', APP)
```

- [ ] **Step 2: README와 인수인계 문서 갱신**

`README.md`의 브라우저 사용법을 다음 순서로 명시한다.

```markdown
1. `데이터 갱신`을 눌러 파일을 선택하거나 화면에 끌어 놓습니다.
2. 파일별 시트·행 수·제외 사유와 적용 전 요약을 확인합니다.
3. `적용`을 누르면 정상 데이터만 반영되고 마지막 정상 상태가 자동 백업됩니다.
4. `백업`에서 현재 데이터 내려받기, 마지막 백업 복원, 백업 파일 불러오기, 초기화를 할 수 있습니다.
5. `엑셀`은 현재 보기와 전체 데이터를 담은 실제 `.xlsx` 파일을 만듭니다.
```

`HANDOFF.md`에는 CSV, XLSX, 날짜 없는 대장, 날짜 없는 AF 매핑, 손상 백업, 작은 화면, 키보드 모달의 수동 검증 항목을 체크박스로 기록한다.

- [ ] **Step 3: 전체 자동 검증**

Run: `node --test tests/dashboard_logic.test.mjs`

Expected: all JavaScript tests PASS.

Run: `uv run --with openpyxl python -m unittest discover -s tests -v`

Expected: all Python tests PASS.

Run: `uv run --with openpyxl python -m compileall -q app.py build_data.py`

Expected: exit code 0 with no output.

Run: `git diff --check`

Expected: exit code 0 with no output.

- [ ] **Step 4: 브라우저 통합 검증**

로컬 정적 서버에서 `dashboard.html`을 열고 다음을 순서대로 확인한다.

1. 초기 데이터 상태에 기준일·기간·행 수·AF 매칭·갱신 시각·출처가 표시된다.
2. UTF-16 탭 구분 PV/UV CSV, 대장 XLSX, AF 매핑 XLSX를 함께 올리면 파일별 로그 후 적용된다.
3. 날짜가 없는 대장 또는 AF 매핑만 올려도 기존 실적을 유지하면서 매핑이 바뀐다.
4. 지원 열이 없는 파일은 적용 버튼이 비활성화되고 기존 데이터가 유지된다.
5. 새로고침하면 마지막 정상 백업과 필터·보기 상태가 복원된다.
6. JSON 백업 내려받기, 초기화, 파일 복원이 순서대로 작동한다.
7. Excel 내보내기 파일을 다시 읽었을 때 `현재 보기`와 `전체 데이터` 시트가 존재한다.
8. 1440px, 1024px, 700px, 390px 폭에서 차트와 표가 겹치거나 잘리지 않는다.
9. 키보드만으로 갱신 모달을 열고 파일 선택·취소·적용하며 닫은 뒤 원래 버튼으로 포커스가 돌아온다.

- [ ] **Step 5: 단일 HTML 산출물 생성과 재검증**

`dashboard.html`의 로컬 SheetJS와 `dashboard_logic.js` 참조를 각각 파일 내용으로 인라인하고, 데이터 원본은 포함하지 않은 상태로 아래 경로에 저장한다.

```text
C:/Users/USER/Documents/Codex/2026-09-28/https-github-com-ryunj-lf-pseo/outputs/pseo_dashboard_fixed.html
```

산출물을 직접 열어 CSV와 XLSX 업로드, 자동 백업, 실제 XLSX 내보내기를 한 번씩 수행한다. 브라우저 개발자 콘솔에 오류가 없어야 한다.

- [ ] **Step 6: 문서 및 최종 통합 커밋**

```bash
git add README.md HANDOFF.md tests/test_dashboard_contract.py
git commit -m "docs: document reliable pSEO data operations"
```

- [ ] **Step 7: 배포 전 상태 확인**

Run: `git status --short`

Expected: no tracked or untracked repository changes. `outputs/pseo_dashboard_fixed.html`은 저장소 밖 산출물이므로 표시되지 않는다.

Run: `git log --oneline -8`

Expected: Tasks 1-7의 독립 커밋과 설계·계획 문서 커밋이 순서대로 표시된다. 원격 반영은 사용자의 실행 승인 범위에 배포가 포함될 때만 `main`을 push한다.
