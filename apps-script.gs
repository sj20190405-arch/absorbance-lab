/**
 * 바이오기초화학 실험실 — 학생 결과 수집용 Google Apps Script
 * 1) 아래 TEACHER_KEY를 선생님만 아는 비밀번호로 바꾸세요.
 * 2) 배포 → 새 배포 → 유형: 웹 앱 / 실행 사용자: 나 / 액세스 권한: 모든 사용자
 * 3) 생성된 웹 앱 URL을 index.html 맨 위 SHEET_URL에 붙여 넣으세요.
 */
const TEACHER_KEY = '여기에-교사-비밀번호';
const SHEET_NAME = '결과';
const HEADERS = ['수신 시각', '학년', '반', '번호', '이름', '구분', '항목', '점수', '상세', '학생 기기 시각', '기록 ID'];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const clip = (v, n) => String(v == null ? '' : v).slice(0, n);
    getSheet_().appendRow([
      new Date(), clip(d.grade, 2), clip(d.cls, 3), clip(d.no, 3), clip(d.name, 20),
      clip(d.kind, 20), clip(d.item, 40), Number(d.score) || 0, clip(d.detail, 300),
      clip(d.clientTime, 40), clip(d.id, 40)
    ]);
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  if ((e.parameter.key || '') !== TEACHER_KEY) {
    return json_({ ok: false, error: '교사 비밀번호가 맞지 않습니다.' });
  }
  const values = getSheet_().getDataRange().getValues();
  values.shift();
  const rows = values.map(r => ({
    time: r[0], grade: r[1], cls: r[2], no: r[3], name: r[4],
    kind: r[5], item: r[6], score: r[7], detail: r[8], clientTime: r[9], id: r[10]
  }));
  return json_({ ok: true, rows });
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
