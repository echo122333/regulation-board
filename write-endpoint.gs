/**
 * 규제혁신 통합관리 — Apps Script v3 (수정 + 과제 추가 + 과거분 시트 지원)
 *
 * [업데이트 방법 — 기존 배포 URL 유지, 2분]
 * 1. script.google.com 에서 기존 프로젝트 열기
 * 2. 코드를 이 내용으로 전부 교체 후 저장
 * 3. 배포 > 배포 관리 > 연필(수정) > 버전: "새 버전" > 배포  → URL 그대로 유지됨
 */
var SHEET_IDS = [
  '1ZXlXeCkBfpXDeB5ix3-_7V-it2GjFTmxhpuSLmIvkJ8', // 과제 로데이터(정상본) — 2026년, 추가는 여기로
  '19EPiytEgm8LEMyyXzqxcXj47jW_h6l5iOhIj_8xsN-M'  // 과거 로데이터(2023~2025)
];

function doPost(e) {
  var body = JSON.parse(e.postData.contents);

  /* ===== 새 과제 추가 ===== */
  if (body.action === 'add') {
    var sheet = SpreadsheetApp.openById(SHEET_IDS[0]).getSheets()[0];
    var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var row = header.map(function(h) {
      var v = body.row[String(h).trim()];
      return v === undefined ? '' : v;
    });
    // 같은 유형+제목 중복 방지
    var values = sheet.getDataRange().getValues();
    var col = {};
    header.forEach(function(h, i) { col[String(h).trim()] = i; });
    for (var r = 1; r < values.length; r++) {
      if (String(values[r][col['유형']]).trim() === String(body.row['유형']).trim() &&
          String(values[r][col['제목']]).trim() === String(body.row['제목']).trim()) {
        return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'duplicate' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    sheet.appendRow(row);
    return ContentService.createTextOutput(JSON.stringify({ ok: true, added: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  /* ===== 기존 과제 수정 ===== */
  var key = body.key || {};
  var upd = body.update || {};
  for (var s = 0; s < SHEET_IDS.length; s++) {
    var sh = SpreadsheetApp.openById(SHEET_IDS[s]).getSheets()[0];
    var vals = sh.getDataRange().getValues();
    var hdr = vals[0];
    var c = {};
    hdr.forEach(function(h, i) { c[String(h).trim()] = i; });
    for (var i2 = 1; i2 < vals.length; i2++) {
      if (String(vals[i2][c['유형']]).trim() === String(key.유형).trim() &&
          String(vals[i2][c['제목']]).trim() === String(key.제목).trim()) {
        var map = { '상태': '진행상태', '결과': '결과', '후속조치': '후속조치', '비고': '비고' };
        for (var k in map) {
          if (upd[k] !== undefined && c[map[k]] !== undefined) {
            sh.getRange(i2 + 1, c[map[k]] + 1).setValue(upd[k]);
          }
        }
        return ContentService.createTextOutput(JSON.stringify({ ok: true, sheet: s, row: i2 + 1 }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'row not found' }))
    .setMimeType(ContentService.MimeType.JSON);
}
