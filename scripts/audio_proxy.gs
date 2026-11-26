// ===== 음성파일 스트리밍 + 싱크 보정값 저장 — Google Apps Script 백엔드 (v7) =====
//
// listen.html(학생 작품 감상)과 play.html(작곡하기 앱 제출 직후 "바로 듣기" 링크)이 공통으로
// 쓰는 백엔드. GitHub Pages는 정적 페이지라 구글 드라이브의 비공개 음성파일을 직접 재생할 수
// 없어서, 이 스크립트가 선생님 계정으로 파일을 읽어 base64로 감싸 돌려줌(doGet 기본 동작).
// 그 외에 "시작 지점/템포 수동 보정값" 저장·조회도 같이 맡고 있음(action=getOffset/saveOffset).
//
// v7에서 바뀐 것: 작곡하기 앱의 "드라이브로 제출" 기능이 이제 음성파일을 반(班)별 하위 폴더
// 안에 저장함(예: AUDIO_ROOT_FOLDER_ID/2반 제출/파일.mp3) — 예전엔 폴더 바로 아래(1단계)만
// 확인했지만, 이제는 몇 단계를 거슬러 올라가도 AUDIO_ROOT_FOLDER_ID 안에 있기만 하면 되도록
// isInAllowedFolder_를 전체 조상 폴더를 거슬러 올라가며 확인하게 고침.
//
// ---- 배포 방법 ----
// 기존에 쓰던 Apps Script 프로젝트(AUDIO_PROXY_URL로 배포된 것)의 코드를 전부 지우고 이 내용
// 으로 바꾼 뒤, "배포 > 배포 관리 > 수정(연필 아이콘) > 새 버전"으로 다시 배포해야 적용됨.
// (배포 주소 자체는 바뀌지 않으니 listen.html/play.html의 AUDIO_PROXY_URL은 그대로 둬도 됨.)

var ALLOWED_FOLDER_ID = '1Mj22wBv_maFrFfS7DxZ0t5DnFlfYXG4i';
var OFFSET_SHEET_NAME = 'SyncOffsets';
var OFFSET_SS_PROP_KEY = 'OFFSET_SS_ID';

// 파일의 조상 폴더들을 계속 거슬러 올라가며(몇 단계든) ALLOWED_FOLDER_ID 안에 있는지 확인함.
// 반별 하위 폴더(예: "2반 제출") 안에 들어있는 파일도 통과하도록 하기 위함.
function isInAllowedFolder_(file) {
  var queue = [];
  var it = file.getParents();
  while (it.hasNext()) queue.push(it.next());
  var seen = {};
  while (queue.length) {
    var folder = queue.shift();
    var id = folder.getId();
    if (seen[id]) continue;
    seen[id] = true;
    if (id === ALLOWED_FOLDER_ID) return true;
    var pit = folder.getParents();
    while (pit.hasNext()) queue.push(pit.next());
  }
  return false;
}

function getOffsetSheet_() {
  var props = PropertiesService.getScriptProperties();
  var ssId = props.getProperty(OFFSET_SS_PROP_KEY);
  var ss = null;
  if (ssId) {
    try { ss = SpreadsheetApp.openById(ssId); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create('음성 싱크 보정값 (자동 생성됨, 지우지 마세요)');
    props.setProperty(OFFSET_SS_PROP_KEY, ss.getId());
  }
  var sheet = ss.getSheetByName(OFFSET_SHEET_NAME);
  if (!sheet) {
    sheet = ss.getSheets()[0];
    sheet.setName(OFFSET_SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['fileId', 'offsetSeconds', 'tempoBPM', 'updatedAt']);
  }
  // 예전 버전(3칸: fileId/offsetSeconds/updatedAt)으로 만들어진 시트를 계속 쓰는 경우를 대비해
  // 헤더에 tempoBPM 칸이 없으면 추가해줌.
  var headers = sheet.getRange(1, 1, 1, Math.max(4, sheet.getLastColumn())).getValues()[0];
  if (headers[2] !== 'tempoBPM') {
    sheet.insertColumnBefore(3);
    sheet.getRange(1, 3).setValue('tempoBPM');
  }
  return sheet;
}

function findOffsetRow_(sheet, id) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === id) return i + 1;
  }
  return -1;
}

function doGet(e) {
  var action = e.parameter.action || 'audio';
  var result;
  try {
    if (action === 'getOffset') {
      var sheet = getOffsetSheet_();
      var row = findOffsetRow_(sheet, e.parameter.id);
      if (row === -1) {
        result = { ok: true, offset: null, tempo: null };
      } else {
        var vals = sheet.getRange(row, 2, 1, 2).getValues()[0];
        result = {
          ok: true,
          offset: (vals[0] === '' || vals[0] === null) ? null : vals[0],
          tempo: (vals[1] === '' || vals[1] === null) ? null : vals[1]
        };
      }

    } else if (action === 'saveOffset') {
      var lock = LockService.getScriptLock();
      lock.waitLock(30000);
      try {
        var sheet2 = getOffsetSheet_();
        var row2 = findOffsetRow_(sheet2, e.parameter.id);
        var now = new Date().toISOString();
        var offsetVal = Number(e.parameter.offset);
        var tempoVal = (e.parameter.tempo !== undefined && e.parameter.tempo !== '')
          ? Number(e.parameter.tempo) : '';
        if (row2 === -1) {
          sheet2.appendRow([e.parameter.id, offsetVal, tempoVal, now]);
        } else {
          sheet2.getRange(row2, 1, 1, 4).setValues([[e.parameter.id, offsetVal, tempoVal, now]]);
        }
        SpreadsheetApp.flush();
      } finally {
        lock.releaseLock();
      }
      result = { ok: true };

    } else {
      var id = e.parameter.id;
      if (!id) {
        result = { ok: false, error: 'missing id parameter' };
      } else {
        var file = DriveApp.getFileById(id);
        if (!isInAllowedFolder_(file)) {
          result = { ok: false, error: 'not allowed' };
        } else {
          var blob = file.getBlob();
          result = {
            ok: true,
            mimeType: blob.getContentType() || 'audio/mpeg',
            base64: Utilities.base64Encode(blob.getBytes())
          };
        }
      }
    }
  } catch (err) {
    result = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
