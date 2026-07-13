/**
 * 리앙에이지 CMS 화면설계서 — 피드백 백엔드 (Google Apps Script)
 * ★ 화면(슬라이드)마다 "개별 시트 탭"에 따로 쌓입니다.
 * ----------------------------------------------------------------------
 * 예) "02 대시보드", "07 문의 접수함", "09 채용 공고" … 화면마다 탭이 자동 생성됩니다.
 * 각 탭 컬럼: id | name | text | time | ts   (탭 이름 자체가 화면이므로 screen 컬럼은 없음)
 *
 * [설치]  구글 시트 → 확장 프로그램 → Apps Script → 이 코드 전체 붙여넣기
 *         → 배포 → 새 배포 → 유형 "웹 앱"
 *            · 실행 계정: 나(내 계정)
 *            · 액세스 권한: 모든 사용자
 *         → 배포 → 웹 앱 URL(...../exec) 복사
 *         → HTML 파일의  const SHEET_ENDPOINT = "여기에 붙여넣기"
 *
 *   코드 수정 후에는 배포 → 배포 관리 → 편집(연필) → "새 버전" → 배포 (URL 유지)
 */

// 첫 화면(개요/로그인 등)에도 대비해, 탭 이름이 없으면 이 이름을 씀
var DEFAULT_TAB = '기타';

// 구글 시트 탭 이름에 못 쓰는 문자 정리 ( : \ / ? * [ ] )
function safeName_(name) {
  var n = String(name || DEFAULT_TAB).replace(/[:\\\/\?\*\[\]]/g, ' ').trim().slice(0, 90);
  return n || DEFAULT_TAB;
}

// 화면(탭) 시트를 찾거나 없으면 만들어서 반환
function getTab_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var title = safeName_(name);
  var sh = ss.getSheetByName(title);
  if (!sh) {
    sh = ss.insertSheet(title);
    sh.appendRow(['id', 'name', 'text', 'time', 'ts']);
    sh.setFrozenRows(1);
  }
  return sh;
}

/** 목록 조회:  GET ?action=list&sheet=07 문의 접수함 */
function doGet(e) {
  var name = (e && e.parameter && (e.parameter.sheet || e.parameter.screen)) || DEFAULT_TAB;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(safeName_(name));
  var rows = [];
  if (sh && sh.getLastRow() > 1) {
    var values = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
    for (var i = 0; i < values.length; i++) {
      var r = values[i];
      if (r[0] === '' && r[2] === '') continue;
      rows.push({ id: String(r[0]), name: r[1], text: r[2], time: r[3], ts: r[4] });
    }
  }
  return json_({ rows: rows });
}

/** 등록/삭제:  POST (text/plain JSON)  {action:'add'|'delete', sheet, ...} */
function doPost(e) {
  var body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) {}
  var name = body.sheet || body.screen || DEFAULT_TAB;

  if (body.action === 'add') {
    var sh = getTab_(name); // 없으면 이 화면 탭을 새로 생성
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
