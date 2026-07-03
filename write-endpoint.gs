/**
 * 규제혁신 통합관리 — 팝업 저장을 구글시트에 반영하는 Apps Script (v2: 과거분 시트 포함)
 *
 * [업데이트 방법 — 기존 배포 URL 유지]
 * 1. script.google.com 에서 기존 프로젝트 열기
 * 2. 코드를 이 내용으로 전부 교체 후 저장
 * 3. 배포 > 배포 관리 > 연필(수정) > 버전: "새 버전" 선택 > 배포
 *    → URL이 그대로 유지되므로 홈페이지 수정 불필요
 */
var SHEET_IDS = [
  '1ZXlXeCkBfpXDeB5ix3-_7V-it2GjFTmxhpuSLmIvkJ8', // 과제 로데이터(정상본) — 2026년
  '19EPiytEgm8LEMyyXzqxcXj47jW_h6l5iOhIj_8xsN-M'  // 과거 로데이터(2023~2025)
];

function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  var key = body.key || {};
  var upd = body.update || {};
  for (var s = 0; s < SHEET_IDS.length; s++) {
    var sheet = SpreadsheetApp.openById(SHEET_IDS[s]).getSheets()[0];
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
        return ContentService.createTextOutput(JSON.stringify({ ok: true, sheet: s, row: r + 1 }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'row not found' }))
    .setMimeType(ContentService.MimeType.JSON);
}
