/**
 * ㈜아성다이소 안전보건팀 — 현장진단 체크리스트 수신 스크립트
 * Google Apps Script Web App
 *
 * ── 배포 방법 ──────────────────────────────────────────────────
 * 1. script.google.com 에서 새 프로젝트 생성
 * 2. 이 코드 붙여넣기
 * 3. SPREADSHEET_ID / PHOTO_FOLDER_ID 를 아래에 입력
 * 4. [배포] → [새 배포] → 유형: 웹 앱
 *    - 다음 사용자로 실행: 나 (스크립트 소유자)
 *    - 액세스 권한: 모든 사용자 (익명 포함)
 * 5. 배포 URL 복사 → index.html 상단 APPS_SCRIPT_URL 에 붙여넣기
 * ────────────────────────────────────────────────────────────────
 */

var SPREADSHEET_ID  = '1GNMt-Dw6iap6SzCJtKa2bETajIxknRwYacPjXV_MaFg';
var PHOTO_FOLDER_ID = '1db5XJF_PxXNvFVgqVK1S264pXQiGxiAY';
var SHEET_NAME      = '진단결과 집계';

// ── POST 수신 (체크리스트 제출) ────────────────────────────────
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);

    // 1) Sheets 행 추가
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(payload.row);

    // 2) 사진 Drive 업로드 (있을 경우)
    var uploaded = 0;
    if (payload.photos && payload.photos.length > 0) {
      var parent = DriveApp.getFolderById(PHOTO_FOLDER_ID);
      var folder = parent.createFolder(payload.folderName || ('진단_' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd_HHmmss')));

      payload.photos.forEach(function(photo) {
        try {
          var b64     = photo.url.replace(/^data:image\/\w+;base64,/, '');
          var decoded = Utilities.base64Decode(b64);
          var blob    = Utilities.newBlob(decoded, 'image/jpeg', photo.name);
          folder.createFile(blob);
          uploaded++;
        } catch (photoErr) { /* 개별 사진 실패 무시 */ }
      });

      // Sheets 마지막 행 끝에 Drive 폴더 링크 추가
      var lastRow = sheet.getLastRow();
      var lastCol = sheet.getLastColumn();
      sheet.getRange(lastRow, lastCol + 1).setValue(folder.getUrl());
    }

    return respond({ ok: true, photos: uploaded });

  } catch (err) {
    return respond({ ok: false, error: err.toString() });
  }
}

// ── GET 헬스체크 ───────────────────────────────────────────────
function doGet(e) {
  return respond({ status: 'alive' });
}

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
