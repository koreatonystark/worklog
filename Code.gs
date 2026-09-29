/**
 * 근무기록 (Daytime · OT) — Google Sheets 백엔드
 * 1) 기록용 Google Sheet 열기 → 확장 프로그램 → Apps Script → 이 코드 전체 붙여넣기
 * 2) 아래 TOKEN을 아무 문자열로 바꾸기 (앱 설정의 Token과 같게)
 * 3) 배포 → 새 배포 → 유형: 웹 앱 / 실행: 나 / 액세스: 모든 사용자(Anyone) → URL 복사
 */
const TOKEN = 'change-me-2026';

const REC_SHEET = 'Records';
const SET_SHEET = 'Settlements';
const REC_COLS = ['id','type','date','weekday','start','end','minutes','hours','status','note','settlementId','text','updatedAt'];
const SET_COLS = ['id','type','from','to','count','minutes','hours','createdAt','text'];

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.token !== TOKEN) return json({ ok: false, error: 'bad token' });
  if (p.action === 'ping') return json({ ok: true, sheet: SpreadsheetApp.getActive().getName() });
  return json({ ok: true, records: readAll(REC_SHEET, REC_COLS), settlements: readAll(SET_SHEET, SET_COLS) });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json({ ok: false, error: 'bad json' }); }
  if (body.token !== TOKEN) return json({ ok: false, error: 'bad token' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    upsert(REC_SHEET, REC_COLS, body.records || []);
    upsert(SET_SHEET, SET_COLS, body.settlements || []);
    SpreadsheetApp.flush();
    return json({ ok: true, at: new Date().toISOString() });
  } finally {
    lock.releaseLock();
  }
}

function sheet_(name, cols) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#1b2320').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.getRange('A:Z').setNumberFormat('@'); // 날짜·시간 자동 변환 방지 (텍스트로 저장)
    sh.setColumnWidth(cols.indexOf('text') + 1, 380);
  }
  return sh;
}

function readAll(name, cols) {
  const sh = sheet_(name, cols);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  return sh.getRange(2, 1, last - 1, head.length).getDisplayValues()
    .filter(r => r[0])
    .map(r => { const o = {}; head.forEach((h, i) => o[h] = r[i]); return o; });
}

function upsert(name, cols, items) {
  if (!items.length) return;
  const sh = sheet_(name, cols);
  const last = sh.getLastRow();
  const ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getDisplayValues().map(r => r[0]) : [];
  const index = {};
  ids.forEach((id, i) => index[id] = i + 2);
  const appends = [];
  items.forEach(it => {
    const row = cols.map(c => it[c] === undefined || it[c] === null ? '' : String(it[c]));
    if (index[it.id]) sh.getRange(index[it.id], 1, 1, cols.length).setValues([row]);
    else { appends.push(row); index[it.id] = -1; }
  });
  if (appends.length) sh.getRange(sh.getLastRow() + 1, 1, appends.length, cols.length).setValues(appends);
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
