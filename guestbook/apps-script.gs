/**
 * 성철 ♥ 주희 청첩장 API (Google Apps Script 웹 앱)
 * - 방명록 저장/조회
 * - 연락처·계좌 정보 제공 (GitHub 저장소에 개인정보를 두지 않기 위해 시트에서 읽어옴)
 * - 참석 의사 저장 ('참석 여부' 탭, 청첩장에는 공개하지 않음)
 *
 * 설정
 * 1. 구글 시트를 새로 만들고 [확장 프로그램 → Apps Script]를 연다.
 * 2. 기본 코드를 지우고 이 파일 내용을 붙여넣은 뒤 저장한다.
 * 3. 위쪽 함수 선택에서 setup 을 고르고 [실행] → 권한 승인. 시트에 '방명록', '연락처·계좌', '참석 여부' 탭이 생긴다.
 * 4. '연락처·계좌' 탭의 전화번호·은행·계좌번호 칸(D2:F7)을 채운다.
 * 5. [배포 → 새 배포 → 유형: 웹 앱] / 실행 계정: 나 / 액세스 권한: 모든 사용자 → 배포.
 * 6. 나오는 웹 앱 URL(https://script.google.com/macros/s/.../exec)을 index.html의 API_URL에 넣는다.
 *
 * 관리
 * - 방명록 글을 숨기려면 '숨김' 칸에 TRUE, 지우려면 행 삭제.
 * - 연락처·계좌를 고치면 청첩장에 바로 반영된다(재배포 필요 없음).
 * - 이 코드를 고친 뒤에는 [배포 → 배포 관리 → ✏️ → 버전: 새 버전 → 배포]로 갱신해야 URL이 그대로 유지된다.
 */
const GB_SHEET = '방명록';
const INFO_SHEET = '연락처·계좌';
const RSVP_SHEET = '참석 여부';
const INFO_ROWS = [
  ['groom', '신랑', '위성철'],
  ['groomFather', '신랑 아버지', '위명환'],
  ['groomMother', '신랑 어머니', '김매자'],
  ['bride', '신부', '남주희'],
  ['brideFather', '신부 아버지', '남기영'],
  ['brideMother', '신부 어머니', '허경아']
];
const MAX_NAME = 20;
const MAX_MESSAGE = 300;

function setup() {
  gbSheet_();
  infoSheet_();
  rsvpSheet_();
}

function rsvpSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(RSVP_SHEET);
  if (!sh) {
    sh = ss.insertSheet(RSVP_SHEET);
    sh.appendRow(['제출시각', '구분', '참석 여부', '성함', '인원(본인 포함)', '식사', '남기실 말씀']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function saveRsvp_(d) {
  const side = { groom: '신랑측', bride: '신부측' }[d.side];
  const attend = { yes: '참석', no: '불참' }[d.attend];
  const name = clean_(d.name, MAX_NAME);
  if (!side || !attend || !name || d.agree !== true) return json_({ ok: false, error: 'invalid' });
  const going = attend === '참석';
  const count = going ? Math.max(1, Math.min(10, parseInt(d.count, 10) || 1)) : 0;
  const meal = going ? ({ yes: '예정', no: '안 함', maybe: '미정' }[d.meal] || '미정') : '';
  const note = clean_(d.note, 100);

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    rsvpSheet_().appendRow([new Date(), side, attend, name, count, meal, note]);
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
}

function gbSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(GB_SHEET);
  if (!sh) {
    sh = ss.insertSheet(GB_SHEET);
    sh.appendRow(['작성시각', '이름', '메시지', '숨김']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function infoSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(INFO_SHEET);
  if (!sh) {
    sh = ss.insertSheet(INFO_SHEET);
    sh.getRange('A:F').setNumberFormat('@'); // 0으로 시작하는 번호가 숫자로 바뀌지 않게
    sh.getRange(1, 1, 1, 6).setValues([['키(수정 금지)', '관계', '이름', '전화번호', '은행', '계좌번호']]);
    sh.getRange(2, 1, INFO_ROWS.length, 3).setValues(INFO_ROWS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// 시트에서 수식으로 해석되지 않도록 막는다
function clean_(value, max) {
  let s = String(value || '').replace(/\r\n?/g, '\n').trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function readInfo_() {
  const rows = infoSheet_().getDataRange().getDisplayValues().slice(1);
  const info = {};
  rows.forEach(r => {
    if (!r[0]) return;
    info[r[0]] = { phone: String(r[3]).trim(), bank: String(r[4]).trim(), account: String(r[5]).trim() };
  });
  return info;
}

function readGuestbook_() {
  return gbSheet_().getDataRange().getValues().slice(1)
    .filter(r => r[1] && r[2] && String(r[3]).toUpperCase() !== 'TRUE')
    .map(r => ({
      at: new Date(r[0]).toISOString(),
      name: String(r[1]).replace(/^'/, ''),
      message: String(r[2]).replace(/^'/, '')
    }))
    .reverse();
}

function doGet() {
  return json_({ ok: true, items: readGuestbook_(), info: readInfo_() });
}

function doPost(e) {
  let d;
  try {
    d = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'bad_request' });
  }
  if (d.website) return json_({ ok: true }); // 스팸 봇용 숨은 칸
  if (d.type === 'rsvp') return saveRsvp_(d);
  const name = clean_(d.name, MAX_NAME);
  const message = clean_(d.message, MAX_MESSAGE);
  if (!name || !message) return json_({ ok: false, error: 'empty' });

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    gbSheet_().appendRow([new Date(), name, message, false]);
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
}
