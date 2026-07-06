/**
 * 규제혁신 통합관리 — Apps Script v5 (연도·차수 매칭 + 과제 추가 + 과거분 시트)
 *
 * ★ 반드시 이 버전으로 재배포해주세요. 현재 배포본(v1)은 과제 추가와 과거분 수정이 작동하지 않습니다.
 *
 * [업데이트 — 기존 URL 유지, 2분]
 * 1. script.google.com 기존 프로젝트 열기 → 코드 전체를 이 내용으로 교체 → 저장
 * 2. 배포 > 배포 관리 > 연필(수정) > 버전: "새 버전" > 배포  → URL 그대로 유지됨
 */
var SHEET_IDS = [
  '1ZXlXeCkBfpXDeB5ix3-_7V-it2GjFTmxhpuSLmIvkJ8', // 과제 로데이터(정상본) — 2026년, 추가는 여기로
  '1E8dxyd8ElV66CdaOHslk8IxRbjbmTC_FDAd_OmFM45E'  // 과거 로데이터(2023~2025) v2
];

function rowMatches(vals, c, key) {
  if (String(vals[c['유형']]).trim() !== String(key.유형 || '').trim()) return false;
  if (String(vals[c['제목']]).trim() !== String(key.제목 || '').trim()) return false;
  // 동명 과제 구분: 연도·차수가 전달되면 함께 대조
  if (key.연도 && c['제출연도'] !== undefined &&
      String(vals[c['제출연도']]).trim() !== String(key.연도).trim()) return false;
  if (key.차수 && c['차수/분기'] !== undefined &&
      String(vals[c['차수/분기']]).trim() !== String(key.차수).trim()) return false;
  return true;
}

function doPost(e) {
  var body = JSON.parse(e.postData.contents);

  /* ===== 새 과제 추가 (정상본 시트로) ===== */
  if (body.action === 'add') {
    var sheet = SpreadsheetApp.openById(SHEET_IDS[0]).getSheets()[0];
    var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var values = sheet.getDataRange().getValues();
    var col = {};
    header.forEach(function(h, i) { col[String(h).trim()] = i; });
    var addKey = { 유형: body.row['유형'], 제목: body.row['제목'],
                   연도: body.row['제출연도'], 차수: body.row['차수/분기'] };
    for (var r = 1; r < values.length; r++) {
      if (rowMatches(values[r], col, addKey)) {
        return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'duplicate' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    sheet.appendRow(header.map(function(h) {
      var v = body.row[String(h).trim()];
      return v === undefined ? '' : v;
    }));
    return ContentService.createTextOutput(JSON.stringify({ ok: true, added: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  /* ===== 기존 과제 수정 (두 시트 모두 검색) ===== */
  var key = body.key || {};
  var upd = body.update || {};
  for (var s = 0; s < SHEET_IDS.length; s++) {
    var sh = SpreadsheetApp.openById(SHEET_IDS[s]).getSheets()[0];
    var vals = sh.getDataRange().getValues();
    var hdr = vals[0];
    var c = {};
    hdr.forEach(function(h, i) { c[String(h).trim()] = i; });
    if (c['유형'] === undefined || c['제목'] === undefined) continue;
    for (var i2 = 1; i2 < vals.length; i2++) {
      if (rowMatches(vals[i2], c, key)) {
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
