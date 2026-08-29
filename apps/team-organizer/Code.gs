/**
 * 체육쌤 스마트 팀 편성기 — 서버 코드 (Google Apps Script)
 *
 * 구글 스프레드시트를 명단 DB로 사용하는 웹앱입니다.
 *  - `명단` 시트       : 학급 / 번호 / 이름 / 성별 / 실력
 *  - `팀편성이력` 시트 : 편성 결과 누적 저장 (반복 조합 감소에 사용)
 */

var SHEET_ROSTER = '명단';
var SHEET_HISTORY = '팀편성이력';
var PROP_SPREADSHEET_ID = 'SPREADSHEET_ID';

var ROSTER_HEADERS = ['학급', '번호', '이름', '성별', '실력'];
var HISTORY_HEADERS = [
  '저장시각', '날짜', '학급', '활동', '팀수', '팀', '번호', '이름', '성별', '실력', '세션ID'
];

/** 이력에서 최근 몇 회차까지 반복 조합 계산에 쓸지 */
var HISTORY_SESSION_LIMIT = 30;
/** 오래된 회차일수록 가중치를 낮추는 감쇠 계수 */
var HISTORY_DECAY = 0.85;

/* ───────────────────────────── 웹앱 진입점 ───────────────────────────── */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('체육쌤 스마트 팀 편성기')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Index.html 안에서 CSS/JS 파일을 끼워 넣을 때 사용 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** 스프레드시트에 바인딩된 경우 메뉴를 추가한다. */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('팀 편성기')
      .addItem('시트 초기 설정', 'setupSheets')
      .addItem('명단 현황 보기', 'showRosterSummary')
      .addToUi();
  } catch (err) {
    // 컨테이너 바인딩이 아니면 무시
  }
}

/* ──────────────────────────── 스프레드시트 접근 ──────────────────────────── */

function getSpreadsheet_() {
  var active = null;
  try {
    active = SpreadsheetApp.getActiveSpreadsheet();
  } catch (err) {
    active = null;
  }
  if (active) return active;

  var id = PropertiesService.getScriptProperties().getProperty(PROP_SPREADSHEET_ID);
  if (!id) {
    throw new Error(
      '스프레드시트가 연결되지 않았습니다. 스크립트 속성에 ' + PROP_SPREADSHEET_ID +
      ' 값을 넣거나, 편집기에서 setSpreadsheetId("시트ID") 를 한 번 실행해 주세요.'
    );
  }
  return SpreadsheetApp.openById(id);
}

/**
 * 독립형(standalone) 스크립트에서 사용할 스프레드시트를 연결한다.
 * 편집기에서 한 번만 실행하면 된다.
 */
function setSpreadsheetId(spreadsheetId) {
  if (!spreadsheetId) throw new Error('스프레드시트 ID를 입력해 주세요.');
  SpreadsheetApp.openById(spreadsheetId); // 접근 가능 여부 확인
  PropertiesService.getScriptProperties().setProperty(PROP_SPREADSHEET_ID, spreadsheetId);
  return '연결 완료: ' + spreadsheetId;
}

/* ────────────────────────────── 초기 설정 ────────────────────────────── */

/**
 * `명단`, `팀편성이력` 시트를 만들고 머리글과 검증 규칙을 세팅한다.
 * 이미 있으면 건드리지 않는다.
 */
function setupSheets() {
  var ss = getSpreadsheet_();

  var roster = ss.getSheetByName(SHEET_ROSTER);
  if (!roster) {
    roster = ss.insertSheet(SHEET_ROSTER, 0);
    roster.getRange(1, 1, 1, ROSTER_HEADERS.length).setValues([ROSTER_HEADERS]);
    roster.getRange(2, 1, 4, ROSTER_HEADERS.length).setValues([
      ['1학년 1반', 1, '김체육', '남', '상'],
      ['1학년 1반', 2, '이운동', '여', '중'],
      ['1학년 1반', 3, '박활동', '남', '하'],
      ['1학년 2반', 1, '최건강', '여', '중']
    ]);
    roster.setFrozenRows(1);
    roster.getRange(1, 1, 1, ROSTER_HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#e8f5e9');
    roster.setColumnWidth(1, 120);
    roster.setColumnWidth(3, 110);

    var genderRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['남', '여'], true).setAllowInvalid(true).build();
    roster.getRange(2, 4, roster.getMaxRows() - 1, 1).setDataValidation(genderRule);

    var levelRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['상', '중', '하'], true).setAllowInvalid(true).build();
    roster.getRange(2, 5, roster.getMaxRows() - 1, 1).setDataValidation(levelRule);
  }

  var history = ss.getSheetByName(SHEET_HISTORY);
  if (!history) {
    history = ss.insertSheet(SHEET_HISTORY);
    history.getRange(1, 1, 1, HISTORY_HEADERS.length).setValues([HISTORY_HEADERS]);
    history.setFrozenRows(1);
    history.getRange(1, 1, 1, HISTORY_HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#fff3e0');
  }

  return '시트 준비 완료 (' + SHEET_ROSTER + ', ' + SHEET_HISTORY + ')';
}

function showRosterSummary() {
  var ui = SpreadsheetApp.getUi();
  var classes = getClasses();
  var total = getRosterRows_().length;
  ui.alert(
    '명단 현황',
    '학급 ' + classes.length + '개 / 학생 ' + total + '명\n\n' + classes.join(', '),
    ui.ButtonSet.OK
  );
}

/* ────────────────────────────── 명단 읽기 ────────────────────────────── */

function normalizeGender_(value) {
  var v = String(value == null ? '' : value).trim().toUpperCase();
  if (!v) return '';
  if (v.charAt(0) === '남' || v === 'M' || v === 'MALE' || v === '1') return '남';
  if (v.charAt(0) === '여' || v === 'F' || v === 'FEMALE' || v === '2') return '여';
  return '';
}

function normalizeLevel_(value) {
  var v = String(value == null ? '' : value).trim().toUpperCase();
  if (!v) return '중';
  if (v.charAt(0) === '상' || v === 'A' || v === 'HIGH' || v === '3') return '상';
  if (v.charAt(0) === '중' || v === 'B' || v === 'MID' || v === 'MIDDLE' || v === '2') return '중';
  if (v.charAt(0) === '하' || v === 'C' || v === 'LOW' || v === '1') return '하';
  return '중';
}

/** 머리글 이름이 조금 달라도 찾아준다. */
function findColumn_(headers, candidates) {
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i] == null ? '' : headers[i]).replace(/\s/g, '');
    for (var j = 0; j < candidates.length; j++) {
      if (h === candidates[j]) return i;
    }
  }
  return -1;
}

function getRosterRows_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(SHEET_ROSTER);
  if (!sheet) {
    throw new Error(
      '`' + SHEET_ROSTER + '` 시트를 찾을 수 없습니다. 메뉴 [팀 편성기 > 시트 초기 설정]을 먼저 실행해 주세요.'
    );
  }
  if (sheet.getLastRow() < 2) return [];

  var width = Math.min(Math.max(sheet.getLastColumn(), ROSTER_HEADERS.length), sheet.getMaxColumns());
  var values = sheet.getRange(1, 1, sheet.getLastRow(), width).getValues();
  var headers = values[0];

  var cCls = findColumn_(headers, ['학급', '반', '학반', '학급명']);
  var cNo = findColumn_(headers, ['번호', '출석번호', '학번']);
  var cName = findColumn_(headers, ['이름', '성명', '학생명']);
  var cGender = findColumn_(headers, ['성별']);
  var cLevel = findColumn_(headers, ['실력', '수준', '등급', '레벨']);

  if (cCls < 0 || cName < 0) {
    throw new Error('`' + SHEET_ROSTER + '` 시트 첫 줄에 `학급`, `이름` 머리글이 있어야 합니다.');
  }

  var rows = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var name = String(row[cName] == null ? '' : row[cName]).trim();
    var cls = String(row[cCls] == null ? '' : row[cCls]).trim();
    if (!name || !cls) continue;

    var no = cNo >= 0 ? String(row[cNo] == null ? '' : row[cNo]).trim() : '';
    rows.push({
      id: 'r' + (r + 1),
      key: makeStudentKey_(cls, no, name), // 이력 대조용 키
      cls: cls,
      no: no,
      name: name,
      gender: cGender >= 0 ? normalizeGender_(row[cGender]) : '',
      level: cLevel >= 0 ? normalizeLevel_(row[cLevel]) : '중'
    });
  }
  return rows;
}

/** 학급/번호/이름으로 만든, 행 위치가 바뀌어도 유지되는 학생 식별자 */
function makeStudentKey_(cls, no, name) {
  var n = String(no == null ? '' : no).trim();
  if (n !== '' && !isNaN(n)) n = String(Number(n)); // '01' 과 1 을 같은 값으로
  return String(cls).trim() + '|' + n + '|' + String(name).trim();
}

/** 시트에 등장한 순서대로 학급 목록을 돌려준다. */
function getClasses() {
  var seen = {};
  var out = [];
  var rows = getRosterRows_();
  for (var i = 0; i < rows.length; i++) {
    if (!seen[rows[i].cls]) {
      seen[rows[i].cls] = true;
      out.push(rows[i].cls);
    }
  }
  return out;
}

/** 웹앱 최초 로딩용 데이터 */
function getBootstrap() {
  var ss = getSpreadsheet_();
  var rows = getRosterRows_();
  var seen = {};
  var classes = [];
  for (var i = 0; i < rows.length; i++) {
    if (!seen[rows[i].cls]) {
      seen[rows[i].cls] = true;
      classes.push(rows[i].cls);
    }
  }
  return {
    classes: classes,
    totalStudents: rows.length,
    activities: getRecentActivities_(),
    spreadsheetUrl: ss.getUrl(),
    spreadsheetName: ss.getName(),
    today: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')
  };
}

/* ────────────────────────────── 이력 읽기 ────────────────────────────── */

function readHistoryRows_() {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_HISTORY);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var width = Math.min(Math.max(sheet.getLastColumn(), HISTORY_HEADERS.length), sheet.getMaxColumns());
  var values = sheet.getRange(1, 1, sheet.getLastRow(), width).getValues();
  var headers = values[0];

  var idx = {
    date: findColumn_(headers, ['날짜']),
    cls: findColumn_(headers, ['학급', '반']),
    activity: findColumn_(headers, ['활동', '종목', '활동명']),
    teamCount: findColumn_(headers, ['팀수']),
    team: findColumn_(headers, ['팀', '팀번호']),
    no: findColumn_(headers, ['번호']),
    name: findColumn_(headers, ['이름', '성명']),
    session: findColumn_(headers, ['세션ID', '세션'])
  };
  if (idx.cls < 0 || idx.team < 0 || idx.name < 0) return [];

  var out = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var name = String(row[idx.name] == null ? '' : row[idx.name]).trim();
    if (!name) continue;
    out.push({
      date: idx.date >= 0 ? formatCell_(row[idx.date]) : '',
      cls: String(row[idx.cls] == null ? '' : row[idx.cls]).trim(),
      activity: idx.activity >= 0 ? String(row[idx.activity] == null ? '' : row[idx.activity]).trim() : '',
      teamCount: idx.teamCount >= 0 ? (Number(row[idx.teamCount]) || 0) : 0,
      team: String(row[idx.team] == null ? '' : row[idx.team]).trim(),
      no: idx.no >= 0 ? String(row[idx.no] == null ? '' : row[idx.no]).trim() : '',
      name: name,
      session: idx.session >= 0 ? String(row[idx.session] == null ? '' : row[idx.session]).trim() : ''
    });
  }
  return out;
}

function formatCell_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value == null ? '' : value).trim();
}

function getRecentActivities_() {
  var rows = readHistoryRows_();
  var seen = {};
  var out = [];
  for (var i = rows.length - 1; i >= 0 && out.length < 15; i--) {
    var a = rows[i].activity;
    if (a && !seen[a]) {
      seen[a] = true;
      out.push(a);
    }
  }
  return out;
}

/**
 * 학급별 과거 편성 이력을 "같은 팀이었던 횟수(가중치)"로 압축한다.
 * 최근 회차일수록 가중치가 높다.
 */
function buildPairWeights_(className) {
  var rows = readHistoryRows_();
  var sessions = {};
  var order = [];

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    if (row.cls !== className) continue;

    var sid = row.session || (row.date + '|' + row.activity);
    if (!sessions[sid]) {
      sessions[sid] = {
        id: sid, date: row.date, activity: row.activity,
        teamCount: row.teamCount, teams: {}
      };
      order.push(sid);
    }
    var session = sessions[sid];
    if (!session.teams[row.team]) session.teams[row.team] = [];
    session.teams[row.team].push(makeStudentKey_(className, row.no, row.name));
  }

  // 최근 회차부터 HISTORY_SESSION_LIMIT 개만 사용
  var recent = order.slice(-HISTORY_SESSION_LIMIT).reverse();
  var weights = {};

  for (var s = 0; s < recent.length; s++) {
    var weight = Math.pow(HISTORY_DECAY, s);
    var teams = sessions[recent[s]].teams;
    for (var t in teams) {
      if (!teams.hasOwnProperty(t)) continue;
      var members = teams[t];
      for (var a = 0; a < members.length; a++) {
        for (var b = a + 1; b < members.length; b++) {
          var k = members[a] < members[b]
            ? members[a] + '::' + members[b]
            : members[b] + '::' + members[a];
          weights[k] = (weights[k] || 0) + weight;
        }
      }
    }
  }

  var summaries = [];
  for (var q = 0; q < recent.length && q < 10; q++) {
    var sess = sessions[recent[q]];
    var count = 0;
    var teamNames = 0;
    for (var tt in sess.teams) {
      if (!sess.teams.hasOwnProperty(tt)) continue;
      count += sess.teams[tt].length;
      teamNames++;
    }
    summaries.push({
      id: sess.id,
      date: sess.date,
      activity: sess.activity,
      teamCount: sess.teamCount || teamNames,
      students: count
    });
  }

  return { weights: weights, sessions: summaries, sessionCount: order.length };
}

/** 학급 선택 시 호출: 명단 + 이력 가중치를 한 번에 내려준다. */
function getClassData(className) {
  if (!className) throw new Error('학급을 선택해 주세요.');

  var rows = getRosterRows_();
  var students = [];
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].cls === className) students.push(rows[i]);
  }
  students.sort(function (a, b) {
    var na = parseInt(a.no, 10);
    var nb = parseInt(b.no, 10);
    if (!isNaN(na) && !isNaN(nb) && na !== nb) return na - nb;
    if (isNaN(na) !== isNaN(nb)) return isNaN(na) ? 1 : -1;
    return a.name < b.name ? -1 : (a.name > b.name ? 1 : 0);
  });

  var history = buildPairWeights_(className);
  return {
    className: className,
    students: students,
    pairWeights: history.weights,
    recentSessions: history.sessions,
    historySessionCount: history.sessionCount
  };
}

/* ────────────────────────────── 이력 저장 ────────────────────────────── */

/**
 * 편성 결과를 `팀편성이력` 시트에 저장한다.
 * payload = { className, activity, date, teams: [{ name, members: [{no,name,gender,level}] }] }
 */
function saveTeamResult(payload) {
  if (!payload || !payload.className) throw new Error('저장할 편성 결과가 없습니다.');
  if (!payload.teams || !payload.teams.length) throw new Error('저장할 팀이 없습니다.');

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName(SHEET_HISTORY);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_HISTORY);
      sheet.getRange(1, 1, 1, HISTORY_HEADERS.length).setValues([HISTORY_HEADERS]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, HISTORY_HEADERS.length).setFontWeight('bold').setBackground('#fff3e0');
    }

    var tz = Session.getScriptTimeZone();
    var now = new Date();
    var savedAt = Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss');
    var date = payload.date || Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    var activity = String(payload.activity || '').trim() || '활동 미기재';
    var sessionId = Utilities.getUuid().substring(0, 8);
    var teamCount = payload.teams.length;

    var rows = [];
    for (var t = 0; t < payload.teams.length; t++) {
      var team = payload.teams[t];
      var teamName = team.name || ((t + 1) + '팀');
      var members = team.members || [];
      for (var m = 0; m < members.length; m++) {
        var s = members[m];
        rows.push([
          savedAt, date, payload.className, activity, teamCount, teamName,
          s.no || '', s.name || '', s.gender || '', s.level || '', sessionId
        ]);
      }
    }
    if (!rows.length) throw new Error('저장할 학생이 없습니다.');

    var startRow = sheet.getLastRow() + 1;
    var shortage = startRow + rows.length - 1 - sheet.getMaxRows();
    if (shortage > 0) sheet.insertRowsAfter(sheet.getMaxRows(), shortage);

    sheet.getRange(startRow, 1, rows.length, HISTORY_HEADERS.length).setValues(rows);

    return {
      sessionId: sessionId,
      savedAt: savedAt,
      rows: rows.length,
      activity: activity,
      sheetUrl: ss.getUrl() + '#gid=' + sheet.getSheetId()
    };
  } finally {
    lock.releaseLock();
  }
}
