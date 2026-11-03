/**
 * 조선유랑기 — 저장 서버 (Google Apps Script)
 * 구글 시트에 붙어 있는 Apps Script 프로젝트에 이 코드를 그대로 붙여 넣고
 * [배포 → 새 배포 → 웹 앱]으로 배포하세요.
 *   - 실행 사용자: 나
 *   - 액세스 권한: 모든 사용자
 * 배포 후 나오는 웹 앱 URL(…/exec)을 index.html 맨 위 SAVE_URL 에 붙여 넣으면 됩니다.
 *
 * 시트 'saves' 에 학생별로 한 줄씩 저장됩니다. (없으면 자동으로 만들어요)
 * 앞쪽 열은 선생님이 진행 상황을 한눈에 보기 위한 요약이고,
 * 마지막 열(데이터)은 게임이 불러올 때 쓰는 원본이니 지우거나 고치지 마세요.
 */
const SHEET_NAME = 'saves';
const HEADERS = ['ID', '학번', '이름', '닉네임', '직업', '레벨', '끝낸 부탁', '엽전', '플레이 시간(분)', '마지막 저장', '데이터(JSON)'];
const CLASS_NAMES = { yangban: '양반', jungin: '중인', pyeongmin: '평민', nobi: '노비' };

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#efe2c2');
  }
  return sh;
}

function findRow_(sh, id) {
  if (!id || sh.getLastRow() < 2) return 0;
  const hit = sh.getRange(2, 1, sh.getLastRow() - 1, 1).createTextFinder(id).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : 0;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

const validId_ = id => /^\d{4,6}_[^\s]{1,20}$/.test(id);

/** 불러오기: GET ?id=학번_이름 */
function doGet(e) {
  const id = String((e && e.parameter && e.parameter.id) || '').trim();
  if (!validId_(id)) return json_({ ok: true, data: null, msg: 'joseon-rpg save server' });
  const sh = getSheet_();
  const row = findRow_(sh, id);
  if (!row) return json_({ ok: true, data: null });
  try {
    return json_({ ok: true, data: JSON.parse(sh.getRange(row, HEADERS.length).getValue()) });
  } catch (err) {
    return json_({ ok: false, data: null, error: 'bad data' });
  }
}

/** 저장: POST {id, data} */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const body = JSON.parse(e.postData.contents);
    const id = String(body.id || '').trim();
    const d = body.data || {};
    if (!validId_(id)) return json_({ ok: false, error: 'bad id' });
    const raw = JSON.stringify(d);
    if (raw.length > 20000) return json_({ ok: false, error: 'too big' });

    const done = Object.keys(d.qs || {}).filter(k => d.qs[k] && d.qs[k].state === 'done').length;
    const rowData = [
      id, "'" + String(d.sid || ''), String(d.name || ''), String(d.nick || '').slice(0, 10),
      CLASS_NAMES[d.cid] || d.cid || '', Number(d.lv) || 1, done + ' / 4', Number(d.coins) || 0,
      Math.round((Number(d.play) || 0) / 3600), new Date(), raw,
    ];
    const sh = getSheet_();
    const row = findRow_(sh, id);
    if (row) sh.getRange(row, 1, 1, rowData.length).setValues([rowData]);
    else sh.appendRow(rowData);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
