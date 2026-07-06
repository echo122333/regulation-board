/**
 * 규제혁신 통합관리 — Apps Script v4 (폴더 자동 탐색: 시트가 바뀌어도 재배포 불필요)
 *
 * [업데이트 — 기존 URL 유지, 2분]
 * 1. script.google.com 기존 프로젝트 열기 → 코드 전체를 이 내용으로 교체 → 저장
 * 2. 배포 > 배포 관리 > 연필(수정) > 버전: "새 버전" > 배포
 * 3. 권한 승인 창이 다시 뜨면 허용 (Drive 접근 권한 추가)
 */
var FOLDER_ID = '1wxNoc4a_BOKEh8JUowhacA3uMV93ReQ8'; // '규제혁신 통합관리' 폴더
var MAIN_HINT = '정상본';   // 새 과제가 추가될 시트 제목에 포함된 단어
var SKIP_HINT = '연간일정'; // 과제 검색에서 제외할 시트

function getSheets() {
  var files = DriveApp.getFolderById(FOLDER_ID).getFilesByType(MimeType.GOOGLE_SHEETS);
  var main = null, others = [];
  while (files.hasNext()) {
    var f = files.next();
    if (f.getName().indexOf(SKIP_HINT) >= 0) continue;
    if (f.getName().indexOf(MAIN_HINT) >= 0) main = f.getId();
    else others.push(f.getId());
  }
  var ids = [];
  if (main) ids.push(main);
  return ids.concat(others);
}

function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  var ids = getSheets();

  /* ===== 새 과제 추가 (메인 시트로) ===== */
  if (body.action === 'add') {
    var sheet = SpreadsheetApp.openById(ids[0]).getSheets()[0];
    var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
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
    sheet.appendRow(header.map(function(h) {
      var v = body.row[String(h).trim()];
      return v === undefined ? '' : v;
    }));
    return ContentService.createTextOutput(JSON.stringify({ ok: true, added: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  /* ===== 기존 과제 수정 (모든 과제 시트 검색) ===== */
  var key = body.key || {};
  var upd = body.update || {};
  for (var s = 0; s < ids.length; s++) {
    var sh = SpreadsheetApp.openById(ids[s]).getSheets()[0];
    var vals = sh.getDataRange().getValues();
    var hdr = vals[0];
    var c = {};
    hdr.forEach(function(h, i) { c[String(h).trim()] = i; });
    if (c['유형'] === undefined || c['제목'] === undefined) continue;
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
