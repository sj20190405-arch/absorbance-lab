/**
 * 바이오기초화학 실험실 — 수행평가 결과 수집용 Google Apps Script
 *
 * 설치: 스프레드시트 → 확장 프로그램 → Apps Script → 이 코드 붙여 넣기 → 저장
 * 배포: 배포 → 새 배포 → 유형 "웹 앱" / 실행 사용자 "나" / 액세스 권한 "모든 사용자"
 * 생성된 웹 앱 URL을 index.html 맨 위 SHEET_URL에 붙여 넣으세요.
 *
 * 교사 비밀번호: 처음에는 0405 입니다. 사이트의 교사 화면 → [비밀번호 변경]에서 바꿀 수 있고,
 * 바꾼 비밀번호는 이 스크립트의 속성(Script Properties)에 저장됩니다.
 * 비밀번호를 잊었다면 아래 resetTeacherKey 함수를 선택해 ▶실행하면 0405로 돌아갑니다.
 */
const DEFAULT_KEY = '0405';
const SHEET_NAME = '결과';
const TRASH_NAME = '삭제된 기록';
const HEADERS = ['수신 시각', '학년', '반', '번호', '이름', '구분', '항목', '점수', '상세', '학생 기기 시각', '기록 ID', '수행점수A', '수행점수B'];

function teacherKey_() {
  return PropertiesService.getScriptProperties().getProperty('TEACHER_KEY') || DEFAULT_KEY;
}
function resetTeacherKey() {
  PropertiesService.getScriptProperties().deleteProperty('TEACHER_KEY');
}
function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(name === TRASH_NAME ? ['삭제 시각'].concat(HEADERS) : HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// 학생 기록 저장
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const clip = (v, n) => String(v == null ? '' : v).slice(0, n);
    const num = v => (v === '' || v == null || isNaN(Number(v))) ? '' : Number(v);
    sheet_(SHEET_NAME).appendRow([
      new Date(), clip(d.grade, 2), clip(d.cls, 3), clip(d.no, 3), clip(d.name, 20),
      clip(d.kind, 20), clip(d.item, 40), Number(d.score) || 0, clip(d.detail, 300),
      clip(d.clientTime, 40), clip(d.id, 40), num(d.pa), num(d.pb)
    ]);
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

// 교사 요청: action = list | reset | setkey
function doGet(e) {
  const p = e.parameter;
  if ((p.key || '') !== teacherKey_()) return json_({ ok: false, error: '교사 비밀번호가 맞지 않습니다.' });
  const action = p.action || 'list';

  if (action === 'list') {
    const values = sheet_(SHEET_NAME).getDataRange().getValues();
    values.shift();
    return json_({ ok: true, rows: values.map(r => ({
      time: r[0], grade: r[1], cls: r[2], no: r[3], name: r[4], kind: r[5], item: r[6],
      score: r[7], detail: r[8], clientTime: r[9], id: r[10], pa: r[11], pb: r[12]
    })) });
  }

  if (action === 'setkey') {
    const nk = String(p.newkey || '');
    if (nk.length < 4) return json_({ ok: false, error: '새 비밀번호는 4자 이상이어야 합니다.' });
    PropertiesService.getScriptProperties().setProperty('TEACHER_KEY', nk);
    return json_({ ok: true });
  }

  if (action === 'reset') {
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const sh = sheet_(SHEET_NAME);
      const values = sh.getDataRange().getValues();
      const header = values.shift();
      const hit = r => p.scope === 'all' ||
        (String(r[1]) === p.grade && String(r[2]) === p.cls && String(r[3]) === p.no && String(r[4]) === p.name);
      const removed = values.filter(hit), keep = values.filter(r => !hit(r));
      if (removed.length) {
        const now = new Date();
        const trash = sheet_(TRASH_NAME);
        const rows = removed.map(r => [now].concat(r, Array(Math.max(0, HEADERS.length - r.length)).fill('')).slice(0, HEADERS.length + 1));
        trash.getRange(trash.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
        sh.clearContents();
        sh.getRange(1, 1, 1, header.length).setValues([header]);
        if (keep.length) sh.getRange(2, 1, keep.length, header.length).setValues(keep);
      }
      return json_({ ok: true, removed: removed.length });
    } finally {
      lock.releaseLock();
    }
  }
  return json_({ ok: false, error: '알 수 없는 요청입니다.' });
}
