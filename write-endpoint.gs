/**
 * 규제혁신 통합관리 — 팝업 저장을 구글시트에 반영하는 Apps Script
 * (시트 ID 직접 지정 방식: 시트에 붙이든 script.google.com에서 단독으로 만들든 동일하게 동작)
 *
 * [설치 — 5분]
 * 1. ck8117@gmail.com 로그인 상태에서 https://script.google.com/create 접속
 *    (또는 시트 > 확장 프로그램 > Apps Script)
 * 2. 이 코드 전체를 붙여넣고 저장
 * 3. 배포 > 새 배포 > 유형: 웹 앱 / 실행 계정: 나 / 액세스: 모든 사용자 > 배포
 * 4. 권한 승인(확인되지 않은 앱 경고 → 고급 → 이동 → 허용)
 * 5. 웹 앱 URL을 복사해 Claude에게 "저장 URL 연결해줘: <URL>" 전달
 */
var SHEET_ID = '1ZXlXeCkBfpXDeB5ix3-_7V-it2GjFTmxhpuSLmIvkJ8'; // 과제 로데이터(정상본)

function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  var key = body.key || {};
  var upd = body.update || {};
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  var values = sheet.getDataRange().getValues();
  var header = values[0];
  var col = {};
  header.forEach(function(h, i) { col[String(h).trim()] = i; });

  for (var r = 1; r < values.length; r++) {
    if (String(values[r][col['유형']]).trim() === String(key.유형).trim() &&
        String(values[r][col['제목']]).trim() === String(key.제목).trim()) {
      var map = { '상태': '진행상태', '결과': '결과', '후속조치': '후속조치', '비고': '비고' };
      for (var k in map) {
        if (upd[k] !== undefined && col[map[k]] !== undefined) {
          sheet.getRange(r + 1, col[map[k]] + 1).setValue(upd[k]);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ ok: true, row: r + 1 }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'row not found' }))
    .setMimeType(ContentService.MimeType.JSON);
}
