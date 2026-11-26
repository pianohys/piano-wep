// ===== 작곡하기 음원 제출함 — Google Apps Script 백엔드 (v2) =====
//
// 목적: 학생이 작곡하기 앱에서 "📤 선생님 드라이브로 제출"을 누르면, 녹음한 음원 파일이
// 이 스크립트를 통해 선생님의 구글 드라이브로 바로 올라감. 기기에 저장했다가 과제 제출
// 창에서 그 파일을 다시 찾아 올리는 단계를 아예 건너뛰어서, "분명 저장했는데 과제 제출할
// 때 파일이 안 보인다"는 문제를 피함(파일 선택창의 종류 필터링 등은 이 앱이 어찌할 수 없는
// 영역이라, 그 단계 자체를 없애는 쪽으로 해결).
//
// v2에서 추가된 것: 제출 성공 직후, 작곡하기 앱이 그 음성파일 ID + 지금 악보 상태를 합쳐서
// "바로 듣기" 링크(play.html)를 만들고, 그 링크를 이 스크립트가 전용 스프레드시트에 한 줄씩
// 자동으로 기록해둠(action=logLink) — 예전처럼 선생님이 엑셀/학생 명단과 음성파일을 일일이
// 수작업으로 매칭할 필요가 없어짐. 음성파일 저장 위치도, 기존에 듣기 페이지(listen.html)가
// 음성파일을 내려받을 때 쓰는 "공유 폴더"(AUDIO_ROOT_FOLDER_ID) 안으로 통일함 — 듣기 페이지
// 쪽 프록시가 그 폴더(및 하위 폴더) 안의 파일만 꺼내주도록 되어 있어서, 여기서도 같은 폴더
// 아래에 저장해야 제출 직후 바로 들을 수 있음.
//
// ---- 배포 방법 ----
// 1. script.google.com 에서 "새 프로젝트"를 만듦 (선생님 본인의 구글 계정으로).
// 2. 기본으로 들어있는 Code.gs의 내용을 전부 지우고, 이 파일(drive_upload.gs) 내용을
//    그대로 붙여넣음.
// 3. 화면 오른쪽 위 "배포" > "새 배포" 클릭.
// 4. 유형에서 톱니바퀴를 눌러 "웹 앱" 선택.
// 5. "실행 계정"은 "나"로, "액세스 권한이 있는 사용자"는 "모든 사용자"로 설정.
// 6. "배포" 클릭 → 생성된 웹 앱 URL(.../exec로 끝남)을 복사.
// 7. 작곡하기 앱(index.html, 작곡하기_ver2.html) 안의 DRIVE_UPLOAD_URL 상수에 그 주소를
//    붙여넣음 — 그러면 "📤 선생님 드라이브로 제출" 버튼이 나타남.
// 8. 코드를 수정할 때마다(처음 배포 이후) "배포 > 배포 관리 > 수정(연필 아이콘) > 새 버전"으로
//    다시 배포해야 변경 사항이 실제로 반영됨 — 그냥 저장만 하면 기존 배포에는 적용 안 됨.
// 9. (기존에 v1을 쓰고 있었다면) AUDIO_ROOT_FOLDER_ID 값이 듣기 페이지(listen.html)의 음성
//    프록시 스크립트에 있는 ALLOWED_FOLDER_ID와 반드시 같아야 함 — 다르면 제출은 되지만
//    "바로 듣기" 링크에서 목소리를 못 받아옴.
//
// ---- 저장 위치 ----
// AUDIO_ROOT_FOLDER_ID 폴더(듣기 페이지가 이미 읽을 수 있는 공유 폴더) 안에, 학생이 입력한
// "반"별로 "N반 제출" 하위 폴더가 자동으로 생겨서 그 안에 파일이 저장됨 — 반마다 폴더가
// 나뉘어 있어서 수합/채점하기 편함. 파일 이름 자체는 제목/학년/반/번호/이름이 들어간 기존
// 규칙 그대로(작곡하기 앱이 로컬에 저장할 때 쓰는 이름 규칙과 동일함).

var AUDIO_ROOT_FOLDER_ID = '1Mj22wBv_maFrFfS7DxZ0t5DnFlfYXG4i';
var LINK_LOG_SHEET_NAME = 'SubmissionLinks';
var LINK_LOG_SS_PROP_KEY = 'LINK_LOG_SS_ID';

function getOrCreateChildFolder_(parent, name) {
  var folders = parent.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return parent.createFolder(name);
}

function getOrCreateSubmitFolder_(clsName) {
  var root = DriveApp.getFolderById(AUDIO_ROOT_FOLDER_ID);
  var clsFolderName = sanitizeFileName_(clsName || '반미기재') + ' 제출';
  return getOrCreateChildFolder_(root, clsFolderName);
}

function sanitizeFileName_(name) {
  return String(name || 'submission').replace(/[\\/:*?"<>|]/g, '_');
}

function getLinkLogSheet_() {
  var props = PropertiesService.getScriptProperties();
  var ssId = props.getProperty(LINK_LOG_SS_PROP_KEY);
  var ss = null;
  if (ssId) {
    try { ss = SpreadsheetApp.openById(ssId); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create('학생 작품 제출 링크 목록 (자동 생성됨, 지우지 마세요)');
    props.setProperty(LINK_LOG_SS_PROP_KEY, ss.getId());
  }
  var sheet = ss.getSheetByName(LINK_LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.getSheets()[0];
    sheet.setName(LINK_LOG_SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['제출시각', '학년', '반', '번호', '이름', 'fileId', 'scoreShareId', '바로 듣기 링크']);
  }
  return sheet;
}

function doPost(e) {
  var result;
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('요청 본문이 비어 있어요.');
    }
    var data = JSON.parse(e.postData.contents);

    if (data.action === 'logLink') {
      var lock = LockService.getScriptLock();
      lock.waitLock(30000);
      try {
        var sheet = getLinkLogSheet_();
        sheet.appendRow([
          new Date(), data.grade || '', data.cls || '', data.num || '', data.name || '',
          data.fileId || '', data.scoreShareId || '', data.playLink || ''
        ]);
        SpreadsheetApp.flush();
      } finally {
        lock.releaseLock();
      }
      result = { ok: true };

    } else {
      if (!data.base64) throw new Error('오디오 데이터가 없어요.');

      var filename = sanitizeFileName_(data.filename || 'submission.mp3');
      var mimeType = data.mimeType || 'audio/mpeg';
      var bytes = Utilities.base64Decode(data.base64);
      var blob = Utilities.newBlob(bytes, mimeType, filename);

      var folder = getOrCreateSubmitFolder_(data.cls);
      var file = folder.createFile(blob);

      result = {
        ok: true,
        fileId: file.getId(),
        fileName: file.getName(),
        fileSize: file.getSize()
      };
    }
  } catch (err) {
    result = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// 배포된 웹 앱 주소를 브라우저로 직접 열어보면(GET 요청) 정상 동작 중인지 확인할 수 있음.
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, message: '작곡하기 제출 백엔드가 정상 동작 중이에요.' }))
    .setMimeType(ContentService.MimeType.JSON);
}
