// ===== 작곡하기 음원 제출함 — Google Apps Script 백엔드 =====
//
// 목적: 학생이 작곡하기 앱에서 "📤 선생님 드라이브로 제출"을 누르면, 녹음한 음원 파일이
// 이 스크립트를 통해 선생님의 구글 드라이브로 바로 올라감. 기기에 저장했다가 과제 제출
// 창에서 그 파일을 다시 찾아 올리는 단계를 아예 건너뛰어서, "분명 저장했는데 과제 제출할
// 때 파일이 안 보인다"는 문제를 피함(파일 선택창의 종류 필터링 등은 이 앱이 어찌할 수 없는
// 영역이라, 그 단계 자체를 없애는 쪽으로 해결).
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
//
// ---- 저장 위치 ----
// 드라이브에 "작곡하기 제출 파일"이라는 폴더가 자동으로 만들어지고(이미 있으면 그대로 씀),
// 그 안에 학생이 입력한 제목/학년/반/번호/이름이 들어간 파일 이름 그대로 저장됨 — 작곡하기
// 앱이 로컬에 저장할 때 쓰는 이름 규칙과 동일함.

var SUBMIT_FOLDER_NAME = '작곡하기 제출 파일';

function getOrCreateSubmitFolder_() {
  var folders = DriveApp.getFoldersByName(SUBMIT_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(SUBMIT_FOLDER_NAME);
}

function sanitizeFileName_(name) {
  return String(name || 'submission').replace(/[\\/:*?"<>|]/g, '_');
}

function doPost(e) {
  var result;
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('요청 본문이 비어 있어요.');
    }
    var data = JSON.parse(e.postData.contents);
    if (!data.base64) throw new Error('오디오 데이터가 없어요.');

    var filename = sanitizeFileName_(data.filename || 'submission.mp3');
    var mimeType = data.mimeType || 'audio/mpeg';
    var bytes = Utilities.base64Decode(data.base64);
    var blob = Utilities.newBlob(bytes, mimeType, filename);

    var folder = getOrCreateSubmitFolder_();
    var file = folder.createFile(blob);

    result = {
      ok: true,
      fileId: file.getId(),
      fileName: file.getName(),
      fileSize: file.getSize()
    };
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
