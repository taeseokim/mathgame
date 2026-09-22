// 노원수학문화축전 - 인덕과학기술고 바이브 코딩 체험 부스 대기자 관리
// Google Sheets를 DB로 사용하는 단일 Apps Script 웹앱 (참가자/스태프/전광판 페이지 통합)

var SHEET_NAME = 'Waitlist';

var STATUS = {
  WAITING: '대기중',
  CALLED: '호출됨',
  IN_PROGRESS: '진행중',
  DONE: '완료',
  NO_SHOW: '노쇼',
  CANCELLED: '취소'
};

var COL = { NUMBER: 1, NAME: 2, CONTACT: 3, STATUS: 4, CREATED: 5, CALLED: 6, DONE: 7 };

function doGet(e) {
  var page = ((e && e.parameter && e.parameter.page) || 'join').toLowerCase();
  var template;
  var title;

  if (page === 'staff') {
    template = 'Staff';
    title = '스태프 대시보드';
  } else if (page === 'board') {
    template = 'Board';
    title = '대기 현황판';
  } else {
    template = 'Index';
    title = '바이브 코딩 체험 접수';
  }

  return HtmlService.createTemplateFromFile(template)
    .evaluate()
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// 행사 전 스프레드시트에서 최초 1회 수동 실행: 시트 헤더와 설정값을 초기화한다.
function initializeSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  sheet.clear();
  sheet.appendRow(['번호', '이름', '연락처', '상태', '접수시각', '호출시각', '완료시각']);
  sheet.setFrozenRows(1);

  var props = PropertiesService.getScriptProperties();
  var existingCode = props.getProperty('STAFF_CODE');
  props.setProperties({
    NEXT_NUMBER: '1',
    STAFF_CODE: existingCode || Math.random().toString(36).slice(2, 8).toUpperCase(),
    MAX_WAITING: '30',
    AVG_MINUTES: '8'
  }, false);

  Logger.log('스태프 접근 코드: ' + props.getProperty('STAFF_CODE'));
}

function getSheet_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('시트를 찾을 수 없습니다. initializeSheet()를 먼저 실행하세요.');
  return sheet;
}

function avgMinutes_() {
  return Number(PropertiesService.getScriptProperties().getProperty('AVG_MINUTES') || 8);
}

function checkStaffCode(code) {
  var real = PropertiesService.getScriptProperties().getProperty('STAFF_CODE');
  return !!code && code === real;
}

function invalidateCache_() {
  CacheService.getScriptCache().remove('snapshot');
}

// 시트 전체를 5초간 캐시해서, 짧은 시간에 몰리는 폴링 요청이 매번 시트를 읽지 않도록 한다.
function getSnapshot_() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('snapshot');
  if (cached) return JSON.parse(cached);

  var data = getSheet_().getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (r[COL.NUMBER - 1] === '') continue;
    rows.push({
      number: Number(r[COL.NUMBER - 1]),
      name: r[COL.NAME - 1],
      status: r[COL.STATUS - 1],
      createdAt: r[COL.CREATED - 1] ? new Date(r[COL.CREATED - 1]).getTime() : 0,
      calledAt: r[COL.CALLED - 1] ? new Date(r[COL.CALLED - 1]).getTime() : 0
    });
  }
  cache.put('snapshot', JSON.stringify(rows), 5);
  return rows;
}

function countByStatus_(rows, status) {
  return rows.filter(function (r) { return r.status === status; }).length;
}

function registerParticipant(name) {
  name = (name || '').toString().trim().slice(0, 20);
  if (!name) throw new Error('이름을 입력해주세요.');

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var rows = getSnapshot_();
    var waitingCount = countByStatus_(rows, STATUS.WAITING);
    var maxWaiting = Number(PropertiesService.getScriptProperties().getProperty('MAX_WAITING') || 30);
    if (waitingCount >= maxWaiting) {
      throw new Error('현재 대기 인원이 많아 접수가 잠시 중단되었습니다. 스태프에게 문의해주세요.');
    }

    var props = PropertiesService.getScriptProperties();
    var next = Number(props.getProperty('NEXT_NUMBER') || '1');
    getSheet_().appendRow([next, name, '', STATUS.WAITING, new Date(), '', '']);
    props.setProperty('NEXT_NUMBER', String(next + 1));
    invalidateCache_();

    return { number: next, waitingAhead: waitingCount, estimatedWaitMin: waitingCount * avgMinutes_() };
  } finally {
    lock.releaseLock();
  }
}

function getStatus(number) {
  number = Number(number);
  var rows = getSnapshot_();
  var entry = rows.filter(function (r) { return r.number === number; })[0];
  if (!entry) return { found: false };

  var waitingAhead = rows.filter(function (r) {
    return r.status === STATUS.WAITING && r.createdAt < entry.createdAt;
  }).length;

  var calledEntry = rows.filter(function (r) {
    return r.status === STATUS.CALLED || r.status === STATUS.IN_PROGRESS;
  }).sort(function (a, b) { return b.calledAt - a.calledAt; })[0];

  return {
    found: true,
    number: entry.number,
    status: entry.status,
    waitingAhead: entry.status === STATUS.WAITING ? waitingAhead : 0,
    estimatedWaitMin: entry.status === STATUS.WAITING ? waitingAhead * avgMinutes_() : 0,
    currentCalledNumber: calledEntry ? calledEntry.number : null
  };
}

function maskName_(name) {
  name = (name || '').toString();
  if (name.length <= 1) return name;
  if (name.length === 2) return name.charAt(0) + '*';
  return name.charAt(0) + new Array(name.length - 1).join('*') + name.charAt(name.length - 1);
}

function getBoardData() {
  var rows = getSnapshot_();
  var active = rows.filter(function (r) {
    return r.status === STATUS.CALLED || r.status === STATUS.IN_PROGRESS;
  }).sort(function (a, b) { return b.calledAt - a.calledAt; });
  var waiting = rows.filter(function (r) {
    return r.status === STATUS.WAITING;
  }).sort(function (a, b) { return a.createdAt - b.createdAt; });

  return {
    current: active.length ? { number: active[0].number, name: maskName_(active[0].name) } : null,
    waitingList: waiting.slice(0, 15).map(function (r) {
      return { number: r.number, name: maskName_(r.name) };
    }),
    waitingCount: waiting.length,
    estimatedWaitMin: waiting.length * avgMinutes_()
  };
}

function getStaffQueue(code) {
  if (!checkStaffCode(code)) throw new Error('접근 코드가 올바르지 않습니다.');
  var rows = getSnapshot_();
  var active = rows.filter(function (r) {
    return r.status === STATUS.WAITING || r.status === STATUS.CALLED || r.status === STATUS.IN_PROGRESS;
  }).sort(function (a, b) { return a.number - b.number; });

  return {
    entries: active,
    stats: {
      waiting: countByStatus_(rows, STATUS.WAITING),
      inProgress: countByStatus_(rows, STATUS.IN_PROGRESS),
      done: countByStatus_(rows, STATUS.DONE),
      total: rows.length
    }
  };
}

function callNext(code) {
  if (!checkStaffCode(code)) throw new Error('접근 코드가 올바르지 않습니다.');
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_();
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][COL.STATUS - 1] === STATUS.WAITING) {
        var row = i + 1;
        sheet.getRange(row, COL.STATUS).setValue(STATUS.CALLED);
        sheet.getRange(row, COL.CALLED).setValue(new Date());
        invalidateCache_();
        return { number: data[i][COL.NUMBER - 1], name: data[i][COL.NAME - 1] };
      }
    }
    return null;
  } finally {
    lock.releaseLock();
  }
}

function updateStatus_(code, number, newStatus, timestampCol) {
  if (!checkStaffCode(code)) throw new Error('접근 코드가 올바르지 않습니다.');
  number = Number(number);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_();
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (Number(data[i][COL.NUMBER - 1]) === number) {
        var row = i + 1;
        sheet.getRange(row, COL.STATUS).setValue(newStatus);
        if (timestampCol) sheet.getRange(row, timestampCol).setValue(new Date());
        invalidateCache_();
        return true;
      }
    }
    return false;
  } finally {
    lock.releaseLock();
  }
}

function markInProgress(code, number) { return updateStatus_(code, number, STATUS.IN_PROGRESS, null); }
function markDone(code, number) { return updateStatus_(code, number, STATUS.DONE, COL.DONE); }
function markNoShow(code, number) { return updateStatus_(code, number, STATUS.NO_SHOW, null); }
function cancelEntry(code, number) { return updateStatus_(code, number, STATUS.CANCELLED, null); }
