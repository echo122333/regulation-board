/**
 * 규제혁신 통합관리 — 팝업 저장을 구글시트에 반영하는 Apps Script
 *
 * [설치 방법 — 5분]
 * 1. "과제 로데이터" 구글시트 열기 → 확장 프로그램 > Apps Script
 * 2. 이 파일 내용을 전부 붙여넣고 저장
 * 3. 배포 > 새 배포 > 유형: 웹 앱
 *    - 실행 계정: 나
 *    - 액세스 권한: 모든 사용자
 * 4. 배포 후 나온 웹 앱 URL을 홈페이지 index.html의 WRITE_ENDPOINT = '' 에 넣기
 *    (또는 Claude에게 "저장 URL 연결해줘: <URL>" 이라고 요청)
 */
function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  var key = body.key || {};
  var upd = body.update || {};
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
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
