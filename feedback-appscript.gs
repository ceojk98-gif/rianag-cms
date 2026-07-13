/**
 * 리앙에이지 CMS 화면설계서 — 피드백 백엔드 (Google Apps Script)
 * ★ 화면(슬라이드)마다 "개별 시트 탭"에 따로 쌓이고, 탭 순서는 번호대로 자동 정렬됩니다.
 * ----------------------------------------------------------------------
 * · 첫 배포와 동일하게: 배포 → 배포 관리 → (연필) → 버전 "새 버전" → 배포  (URL 유지)
 * · 재배포 후 브라우저에서  <웹앱URL>?action=setup  을 1회 열면 00~11 탭이 한 번에 생성·정렬됩니다.
 */

var DEFAULT_TAB = '기타';

// 12개 화면 탭 (프론트 tabName 과 동일 규칙). '/' 는 시트 탭에서 공백으로 저장됨.
var SCREEN_TABS = [
  '00 개요 · 구성', '01 로그인', '02 대시보드', '03 프로젝트 목록',
  '04 프로젝트 등록/수정', "05 What's on 목록", "06 What's on 등록/수정",
  '07 문의 접수함', '08 문의 폼 설정', '09 채용 공고', '10 입사지원 접수함', '11 설정'
];

// 구글 시트 탭 이름에 못 쓰는 문자 정리 ( : \ / ? * [ ] )
function safeName_(name) {
  var n = String(name || DEFAULT_TAB).replace(/[:\\\/\?\*\[\]]/g, ' ').trim().slice(0, 90);
  return n || DEFAULT_TAB;
}

// "00 …" ~ "11 …" 형태의 탭을 번호 순으로 정렬(왼쪽부터). 기타 탭(시트1 등)은 뒤로 밀림.
function sortTabs_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var numbered = ss.getSheets().filter(function (s) { return /^\d\d\s/.test(s.getName()); });
  numbered.sort(function (a, b) { var x = a.getName(), y = b.getName(); return x < y ? -1 : (x > y ? 1 : 0); });
  for (var i = 0; i < numbered.length; i++) {
    ss.setActiveSheet(numbered[i]);
    ss.moveActiveSheet(i + 1); // 1-based
  }
}

// 화면(탭) 시트를 찾거나 없으면 만들어서 반환 (새로 만들면 정렬)
function getTab_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var title = safeName_(name);
  var sh = ss.getSheetByName(title);
  if (!sh) {
    sh = ss.insertSheet(title);
    sh.appendRow(['id', 'name', 'text', 'time', 'ts']);
    sh.setFrozenRows(1);
    sortTabs_();
  }
  return sh;
}

/** GET: ?action=list&sheet=07 문의 접수함   |   ?action=setup (12탭 일괄 생성·정렬) */
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'list';
  if (action === 'setup') return setup_();

  var name = (e && e.parameter && (e.parameter.sheet || e.parameter.screen)) || DEFAULT_TAB;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(safeName_(name));
  var rows = [];
  if (sh && sh.getLastRow() > 1) {
    var v = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
    for (var i = 0; i < v.length; i++) {
      var r = v[i];
      if (r[0] === '' && r[2] === '') continue;
      rows.push({ id: String(r[0]), name: r[1], text: r[2], time: r[3], ts: r[4] });
    }
  }
  return json_({ rows: rows });
}

// 12개 탭 한 번에 생성 + 안내용 시드 1줄 + 번호순 정렬
function setup_() {
  for (var i = 0; i < SCREEN_TABS.length; i++) {
    var sh = getTab_(SCREEN_TABS[i]);
    if (sh.getLastRow() < 2) {
      sh.appendRow(['seed' + i, '시스템', '(자동 생성) 이 화면의 피드백이 여기에 쌓입니다. 이 줄은 지워도 됩니다.',
        new Date().toISOString(), Date.now()]);
    }
  }
  sortTabs_();
  return json_({ ok: true, tabs: SCREEN_TABS.length });
}

/** POST (text/plain JSON): {action:'add'|'delete', sheet, ...} */
function doPost(e) {
  var body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) {}
  var name = body.sheet || body.screen || DEFAULT_TAB;

  if (body.action === 'add') {
    var sh = getTab_(name);
    sh.appendRow([body.id, body.name, body.text, body.time, body.ts]);
    return json_({ ok: true });
  }
  if (body.action === 'delete') {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh2 = ss.getSheetByName(safeName_(name));
    if (sh2) {
      var vals = sh2.getDataRange().getValues();
      for (var i = vals.length - 1; i >= 1; i--) {
        if (String(vals[i][0]) === String(body.id)) sh2.deleteRow(i + 1);
      }
    }
    return json_({ ok: true });
  }
  return json_({ ok: false, error: 'unknown action' });
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
