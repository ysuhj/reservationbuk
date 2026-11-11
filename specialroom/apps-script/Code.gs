/**
 * 북가좌초등학교 특별실 현황·예약 API (Google Apps Script)
 * 화면은 GitHub Pages(docs 폴더)에 있고, 이 스크립트는 시트를 읽고 쓰는 API만 맡습니다.
 *
 * 설정: Apps Script 편집기 → 프로젝트 설정(톱니바퀴) → 스크립트 속성에
 *   SCHOOL_PASSWORD = 선생님들께 알려 줄 학교 비밀번호
 * 를 추가하세요. 비밀번호는 코드나 GitHub에 적지 않습니다.
 */
const CONFIG = {
  SPREADSHEET_ID: '1DR11TocphTfxy7Nyh8GU31NgpX8RUqpAbBLlzeAaSeY',
  SCHOOL_NAME: '북가좌초등학교',
  EXCLUDE_SHEETS: ['안내', '회의실(1관 1층)', '(전산실) 태블릿 대여', '시설 보수', '공지사항',
    '예약기록', '시정표', '소통보드'],
  LOG_SHEET: '예약기록',
  NOTICE_SHEET: '공지사항',
  MEMO_SHEET: '소통보드',
  BELL_SHEET: '시정표',
  // 처음 만들어지는 시정표 (시트에서 고치면 바로 반영)
  DEFAULT_BELL: [
    ['1교시', '09:00', '09:40'], ['2교시', '09:50', '10:30'], ['3교시', '10:50', '11:30'],
    ['4교시', '11:40', '12:20'], ['5교시', '13:20', '14:00'], ['6교시', '14:10', '14:50']
  ],
  FREE_MARKERS: ['', '-'],
  DEFAULT_CHANGE_COLOR: '#fff2cc',
  MAX_CONTENT: 30, MAX_NAME: 20, MAX_MEMO: 300,
  CACHE_SECONDS: 60
};

const LOG_HEADERS = ['기록시각', '특별실', '날짜', '교시', '내용', '신청자',
  '이전 내용', '이전 배경색', '이전 메모', '상태', '변경시각'];
const MEMO_HEADERS = ['ID', '작성시각', '작성자', '내용', '상태'];

/* ───────── API 입구 ───────── */
const API = {
  login: () => ({ ok: true }),
  getInit, getWeek, getDay, getMonth, reserve, modify, getMemos, addMemo, deleteMemo
};

function doGet() {
  return json_({ ok: true, data: '북가좌초 특별실 API가 동작 중입니다.' });
}

function doPost(e) {
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const pw = PropertiesService.getScriptProperties().getProperty('SCHOOL_PASSWORD');
    if (!pw) return json_({ ok: false, error: '관리자가 아직 학교 비밀번호(SCHOOL_PASSWORD)를 설정하지 않았습니다.' });
    if (String(req.pw || '') !== pw) {
      Utilities.sleep(700); // 비밀번호 마구 대입하기 늦추기
      return json_({ ok: false, code: 'AUTH', error: '학교 비밀번호가 맞지 않습니다.' });
    }
    if (!Object.prototype.hasOwnProperty.call(API, req.fn)) return json_({ ok: false, error: '알 수 없는 요청입니다.' });
    const data = API[req.fn].apply(null, Array.isArray(req.args) ? req.args : []);
    return json_({ ok: true, data });
  } catch (err) {
    return json_({ ok: false, error: (err && err.message) || String(err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ───────── 공통 ───────── */
let SS_ = null;
function ss_() { return SS_ || (SS_ = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)); }
const pad_ = n => String(n).padStart(2, '0');
const isoOf_ = (y, m, d) => `${y}-${pad_(m)}-${pad_(d)}`;
const weekday_ = (y, m, d) => new Date(Date.UTC(y, m - 1, d)).getUTCDay();
const isPeriod_ = s => /^\d+\s*교시$/.test(String(s).trim());
const isFree_ = s => CONFIG.FREE_MARKERS.includes(String(s).trim());
const txt_ = s => "'" + String(s);
const addIso_ = (iso, n) => {
  const [y, m, d] = iso.split('-').map(Number);
  const t = new Date(Date.UTC(y, m - 1, d) + n * 86400000);
  return isoOf_(t.getUTCFullYear(), t.getUTCMonth() + 1, t.getUTCDate());
};

function today_() {
  const iso = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  const [y, m, d] = iso.split('-').map(Number);
  return { iso, y, m, d };
}
function stamp_() { return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm'); }

function ensureSheet_(name, headers, hidden) {
  let sh = ss_().getSheetByName(name);
  if (!sh) {
    sh = ss_().insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    if (hidden) sh.hideSheet();
  }
  return sh;
}

/* ───────── 시트 해석 ───────── */
function listRooms_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('rooms');
  if (hit) return JSON.parse(hit);
  const rooms = ss_().getSheets().filter(sh => {
    if (CONFIG.EXCLUDE_SHEETS.includes(sh.getName())) return false;
    const n = Math.min(30, sh.getMaxRows());
    return sh.getRange(1, 2, n, 1).getDisplayValues().some(r => isPeriod_(r[0]));
  }).map(sh => sh.getName());
  cache.put('rooms', JSON.stringify(rooms), 600);
  return rooms;
}

function checkRoom_(room) {
  if (!listRooms_().includes(room)) throw new Error('특별실 이름을 찾을 수 없습니다: ' + room);
  return ss_().getSheetByName(room);
}

function readLegend_(v, bg) {
  const map = { '시간표 변경': 'change', '학교행사': 'event', '공휴일': 'holiday', '사용 가능한 시간표': 'available' };
  const out = {};
  for (let r = 0; r < Math.min(10, v.length); r++) {
    for (let c = 0; c < v[r].length; c++) {
      const key = map[String(v[r][c]).trim()];
      if (!key) continue;
      const own = bg[r][c].toLowerCase();
      const left = c > 0 ? bg[r][c - 1].toLowerCase() : '#ffffff';
      const color = own !== '#ffffff' ? own : left;
      if (color !== '#ffffff') out[key] = color;
    }
  }
  return out;
}

function yearFor_(month, day, weekday) {
  const t = today_();
  const start = t.m >= 3 ? t.y : t.y - 1;
  const base = month >= 3 ? start : start + 1;
  return [base, base + 1, base - 1].find(y => weekday_(y, month, day) === weekday) || null;
}

function parseRoom_(sh) {
  const lastRow = sh.getLastRow();
  const lastCol = Math.min(sh.getLastColumn(), 26);
  const rng = sh.getRange(1, 1, lastRow, lastCol);
  const v = rng.getDisplayValues();
  const bg = rng.getBackgrounds();
  const legend = readLegend_(v, bg);

  const sections = [];
  for (let r = 0; r < Math.min(12, lastRow) && !sections.length; r++) {
    for (let c = 0; c < lastCol; c++) {
      const m = String(v[r][c]).trim().match(/^(\d{1,2})월$/);
      if (m) sections.push({ col: c, month: Number(m[1]) });
    }
  }

  const days = {};
  for (let r = 1; r < lastRow; r++) {
    if (!isPeriod_(v[r][1]) || isPeriod_(v[r - 1][1])) continue;
    const dateRow = r - 1;
    const pRows = [];
    for (let k = r; k < lastRow && isPeriod_(v[k][1]); k++) pRows.push(k);
    sections.forEach(sec => {
      for (let i = 0; i < 5; i++) {
        const c = sec.col + i;
        if (c >= lastCol) continue;
        const dm = String(v[dateRow][c] || '').trim().match(/^(\d{1,2})\s*(.*)$/);
        if (!dm) continue;
        const day = Number(dm[1]);
        const y = yearFor_(sec.month, day, i + 1);
        if (!y) continue;
        days[isoOf_(y, sec.month, day)] = {
          note: dm[2] || '',
          slots: pRows.map(pr => ({
            period: String(v[pr][1]).trim().replace(/\s+/g, ''),
            row: pr + 1, col: c + 1,
            text: String(v[pr][c]).trim(),
            bg: bg[pr][c].toLowerCase()
          }))
        };
      }
    });
  }
  return { legend, days };
}

const cacheKey_ = room => 'p:' + Utilities.base64EncodeWebSafe(room);

// 여러 특별실을 한 번에 읽되, 1분 동안은 캐시 사용
function parsedMany_(rooms) {
  const cache = CacheService.getScriptCache();
  const hits = cache.getAll(rooms.map(cacheKey_));
  const out = {};
  rooms.forEach(room => {
    const k = cacheKey_(room);
    if (hits[k]) { out[room] = JSON.parse(hits[k]); return; }
    const p = parseRoom_(ss_().getSheetByName(room));
    try { cache.put(k, JSON.stringify(p), CONFIG.CACHE_SECONDS); } catch (e) { /* 너무 크면 캐시 생략 */ }
    out[room] = p;
  });
  return out;
}
function bust_(room) { CacheService.getScriptCache().remove(cacheKey_(room)); }

function slotStatus_(slot, legend) {
  if (legend.holiday && slot.bg === legend.holiday) return 'closed';
  return isFree_(slot.text) ? 'free' : 'taken';
}

/* ───────── 기록 ───────── */
function logSheet_() { return ensureSheet_(CONFIG.LOG_SHEET, LOG_HEADERS, true); }

function readLog_() {
  const sh = logSheet_();
  const n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, LOG_HEADERS.length).getDisplayValues().map((r, i) => ({
    row: i + 2, room: r[1], iso: r[2], period: r[3], content: r[4], name: r[5],
    prevValue: r[6], prevBg: r[7], prevNote: r[8], status: r[9]
  }));
}

// { room: { 'iso|교시': { booked, cleared } } }
function activeAll_() {
  const map = {};
  readLog_().forEach(e => {
    if (e.status !== '예약' && e.status !== '비움') return;
    const r = map[e.room] || (map[e.room] = {});
    const k = `${e.iso}|${e.period}`;
    const o = r[k] || (r[k] = {});
    if (e.status === '예약') o.booked = e; else o.cleared = e;
  });
  return map;
}

function slotView_(s, legend, act) {
  const status = slotStatus_(s, legend);
  const a = act && act.booked;
  const mine = !!a && a.content === s.text;
  const cl = act && act.cleared;
  return {
    period: s.period,
    text: isFree_(s.text) ? '' : s.text,
    status,
    by: mine ? a.name : '',
    isNew: status === 'taken' && (mine || (!!legend.change && s.bg === legend.change)),
    orig: mine && !isFree_(a.prevValue) ? a.prevValue : '',
    cleared: status === 'free' && cl ? cl.prevValue : ''
  };
}

function bell_() {
  const sh = ensureSheet_(CONFIG.BELL_SHEET, ['교시', '시작', '종료'], false);
  if (sh.getLastRow() < 2) {
    sh.getRange(2, 1, CONFIG.DEFAULT_BELL.length, 3).setNumberFormat('@').setValues(CONFIG.DEFAULT_BELL);
  }
  const out = {};
  sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), 3).getDisplayValues().forEach(r => {
    const p = String(r[0]).replace(/\s+/g, '');
    if (isPeriod_(p) && r[1]) out[p] = r[2] ? `${r[1]}–${r[2]}` : r[1];
  });
  return out;
}

function notices_() {
  const sh = ss_().getSheetByName(CONFIG.NOTICE_SHEET);
  if (!sh || sh.getLastRow() < 1) return [];
  return sh.getRange(1, 1, sh.getLastRow(), Math.min(4, sh.getLastColumn())).getDisplayValues()
    .filter(r => r[1] && r[2] && r[1] !== '공지사항')
    .map(r => ({ title: String(r[1]).replace(/\s*\n\s*/g, ' '), body: String(r[2]).replace(/\r/g, '') }));
}

/* ───────── 화면 데이터 ───────── */
function getInit() {
  return {
    today: today_().iso,
    rooms: listRooms_(),
    bell: bell_(),
    notices: notices_(),
    memos: getMemos(),
    sheetUrl: ss_().getUrl(),
    school: CONFIG.SCHOOL_NAME
  };
}

function weekDays_(mondayIso) {
  return ['월', '화', '수', '목', '금'].map((dow, i) => {
    const iso = addIso_(mondayIso, i);
    const [, m, d] = iso.split('-').map(Number);
    return { iso, dow, label: `${m}/${d}` };
  });
}

function getWeek(room, mondayIso) {
  checkRoom_(room);
  const p = parsedMany_([room])[room];
  const act = activeAll_()[room] || {};
  const today = today_().iso;
  let periods = [];
  const days = weekDays_(mondayIso).map(w => {
    const day = p.days[w.iso];
    if (day && day.slots.length > periods.length) periods = day.slots.map(s => s.period);
    return Object.assign(w, {
      note: day ? day.note : '', exists: !!day, past: w.iso < today,
      slots: day ? day.slots.map(s => slotView_(s, p.legend, act[`${w.iso}|${s.period}`])) : []
    });
  });
  return { room, monday: mondayIso, periods, days, hasData: days.some(x => x.exists) };
}

function getDay(iso) {
  const rooms = listRooms_();
  const parsed = parsedMany_(rooms);
  const act = activeAll_();
  let periods = [], note = '';
  const cols = rooms.map(room => {
    const p = parsed[room], day = p.days[iso];
    if (day && day.slots.length > periods.length) periods = day.slots.map(s => s.period);
    if (day && day.note && !note) note = day.note;
    const a = act[room] || {};
    return {
      room, exists: !!day,
      slots: day ? day.slots.map(s => slotView_(s, p.legend, a[`${iso}|${s.period}`])) : []
    };
  });
  const [y, m, d] = iso.split('-').map(Number);
  return {
    iso, note, periods, rooms: cols, past: iso < today_().iso,
    dow: '일월화수목금토'[weekday_(y, m, d)], hasData: cols.some(c => c.exists)
  };
}

// 월간: 날짜마다 교시별 상태를 한 글자로 (f 빈, t 기존, n 새, c 휴일)
function getMonth(room, year, month) {
  checkRoom_(room);
  const p = parsedMany_([room])[room];
  const act = activeAll_()[room] || {};
  const first = isoOf_(year, month, 1);
  const wd = weekday_(year, month, 1);
  let mon = addIso_(first, wd === 0 ? 1 : wd === 6 ? 2 : 1 - wd);
  const weeks = [];
  while (true) {
    const days = weekDays_(mon).map(w => {
      const day = p.days[w.iso];
      const inMonth = Number(w.iso.slice(5, 7)) === month;
      if (!day) return Object.assign(w, { inMonth, exists: false, note: '', marks: '', free: 0 });
      const views = day.slots.map(s => slotView_(s, p.legend, act[`${w.iso}|${s.period}`]));
      return Object.assign(w, {
        inMonth, exists: true, note: day.note,
        marks: views.map(v => v.status === 'closed' ? 'c' : v.status === 'free' ? 'f' : v.isNew ? 'n' : 't').join(''),
        free: views.filter(v => v.status === 'free').length
      });
    });
    weeks.push({ monday: mon, days });
    mon = addIso_(mon, 7);
    if (Number(mon.slice(5, 7)) !== month) break;
  }
  return { room, year, month, weeks, today: today_().iso };
}

/* ───────── 예약·변경 ───────── */
function checkText_(content, name) {
  if (!name) throw new Error('이름을 입력해 주세요.');
  if (name.length > CONFIG.MAX_NAME) throw new Error(`이름은 ${CONFIG.MAX_NAME}자까지 입력할 수 있습니다.`);
  if (content !== null) {
    if (!content) throw new Error('학급 또는 수업명을 입력해 주세요.');
    if (content.length > CONFIG.MAX_CONTENT) throw new Error(`학급/수업명은 ${CONFIG.MAX_CONTENT}자까지 입력할 수 있습니다.`);
  }
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function reserve(req) {
  const content = String(req.content || '').trim();
  const name = String(req.name || '').trim();
  const items = Array.isArray(req.items) ? req.items : [];
  checkText_(content, name);
  if (!items.length) throw new Error('예약할 시간을 선택해 주세요.');

  return withLock_(() => {
    const sh = checkRoom_(req.room);
    const p = parseRoom_(sh);
    const today = today_().iso;
    const color = p.legend.change || CONFIG.DEFAULT_CHANGE_COLOR;
    const targets = items.map(it => {
      const day = p.days[it.iso];
      const slot = day && day.slots.find(s => s.period === it.period);
      if (!slot) throw new Error(`${it.iso} ${it.period}은(는) 시트에 없는 시간입니다.`);
      if (it.iso < today) throw new Error(`${it.iso}은(는) 지난 날짜입니다.`);
      const st = slotStatus_(slot, p.legend);
      if (st === 'closed') throw new Error(`${it.iso} ${it.period}은(는) 휴일입니다.`);
      if (st !== 'free') throw new Error(`${req.room} ${it.iso} ${it.period}은(는) 방금 ‘${slot.text}’(으)로 예약되었습니다. 새로고침 후 다시 선택해 주세요.`);
      return { it, slot };
    });
    const log = logSheet_();
    const when = stamp_();
    targets.forEach(({ it, slot }) => {
      const cell = sh.getRange(slot.row, slot.col);
      const prevNote = cell.getNote();
      cell.setNumberFormat('@').setValue(content).setBackground(color).setNote(`웹 예약: ${name} (${when})`);
      log.appendRow([when, txt_(req.room), txt_(it.iso), txt_(it.period), txt_(content), txt_(name),
        txt_(slot.text), slot.bg, txt_(prevNote), '예약', '']);
    });
    SpreadsheetApp.flush();
    bust_(req.room);
    return { count: targets.length };
  });
}

/**
 * action: 'edit' 수정 | 'clear' 이 날만 비우기·예약 취소 | 'revert' 원래대로 | 'unclear' 비우기 취소
 */
function modify(req) {
  const name = String(req.name || '').trim();
  const action = req.action;
  const content = String(req.content || '').trim();
  checkText_(action === 'edit' ? content : null, name);
  if (!['edit', 'clear', 'revert', 'unclear'].includes(action)) throw new Error('알 수 없는 요청입니다.');

  return withLock_(() => {
    const sh = checkRoom_(req.room);
    const p = parseRoom_(sh);
    if (req.iso < today_().iso) throw new Error('지난 날짜는 바꿀 수 없습니다.');
    const day = p.days[req.iso];
    const slot = day && day.slots.find(s => s.period === req.period);
    if (!slot) throw new Error('시트에서 해당 시간을 찾을 수 없습니다.');
    const act = (activeAll_()[req.room] || {})[`${req.iso}|${req.period}`] || {};
    const cell = sh.getRange(slot.row, slot.col);
    const log = logSheet_();
    const when = stamp_();
    const color = p.legend.change || CONFIG.DEFAULT_CHANGE_COLOR;
    const done = () => { SpreadsheetApp.flush(); bust_(req.room); return { ok: true }; };

    if (action === 'unclear') {
      const cl = act.cleared;
      if (!cl || !isFree_(slot.text)) throw new Error('비우기를 취소할 수 없는 칸입니다. 새로고침해 주세요.');
      cell.setNumberFormat('@').setValue(cl.prevValue).setBackground(cl.prevBg || null).setNote(cl.prevNote || '');
      log.getRange(cl.row, 10, 1, 2).setValues([['비움취소', when]]);
      return done();
    }

    if (slotStatus_(slot, p.legend) !== 'taken') throw new Error('비어 있거나 휴일인 시간입니다. 새로고침해 주세요.');
    if (slot.text !== String(req.expect || '')) {
      throw new Error(`그사이 내용이 ‘${slot.text}’(으)로 바뀌었습니다. 새로고침 후 다시 시도해 주세요.`);
    }
    const entry = act.booked && act.booked.content === slot.text ? act.booked : null;
    const own = entry && entry.name === name;
    if (!own && !req.agreed) throw new Error('기존 사용 학급(선생님)과 협의했는지 확인해 주세요.');

    const orig = entry
      ? { value: entry.prevValue, bg: entry.prevBg, note: entry.prevNote }
      : { value: slot.text, bg: slot.bg, note: cell.getNote() };
    const origFree = isFree_(orig.value);
    if (entry) log.getRange(entry.row, 10, 1, 2).setValues([['변경됨', when]]);

    if (action === 'edit') {
      cell.setNumberFormat('@').setValue(content).setBackground(color)
        .setNote(`웹 수정: ${name} (${when})` + (origFree ? '' : `\n원래: ${orig.value}`));
      log.appendRow([when, txt_(req.room), txt_(req.iso), txt_(req.period), txt_(content), txt_(name),
        txt_(orig.value), orig.bg, txt_(orig.note), '예약', '']);
    } else if (action === 'revert' || (action === 'clear' && origFree)) {
      if (!entry) throw new Error('되돌릴 기록이 없는 칸입니다.');
      cell.setNumberFormat('@').setValue(orig.value).setBackground(orig.bg || null).setNote(orig.note || '');
      log.appendRow([when, txt_(req.room), txt_(req.iso), txt_(req.period), txt_(orig.value), txt_(name),
        txt_(slot.text), slot.bg, '', origFree ? '취소' : '되돌림', when]);
    } else {
      cell.setNumberFormat('@').setValue('').setBackground(color)
        .setNote(`웹 비움: ${name} (${when})\n원래: ${orig.value}`);
      log.appendRow([when, txt_(req.room), txt_(req.iso), txt_(req.period), '', txt_(name),
        txt_(orig.value), orig.bg, txt_(orig.note), '비움', '']);
    }
    return done();
  });
}

/* ───────── 소통 보드 ───────── */
function memoSheet_() { return ensureSheet_(CONFIG.MEMO_SHEET, MEMO_HEADERS, false); }

function getMemos() {
  const sh = memoSheet_();
  const n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, MEMO_HEADERS.length).getDisplayValues()
    .filter(r => r[4] !== '삭제')
    .map(r => ({ id: r[0], time: r[1], name: r[2], text: r[3] }))
    .reverse().slice(0, 30);
}

function addMemo(req) {
  const name = String(req.name || '').trim();
  const text = String(req.text || '').trim();
  if (!name) throw new Error('이름을 입력해 주세요.');
  if (!text) throw new Error('메모 내용을 입력해 주세요.');
  if (text.length > CONFIG.MAX_MEMO) throw new Error(`메모는 ${CONFIG.MAX_MEMO}자까지 쓸 수 있습니다.`);
  withLock_(() => memoSheet_().appendRow([Utilities.getUuid().slice(0, 8), stamp_(), txt_(name), txt_(text), '']));
  return getMemos();
}

function deleteMemo(req) {
  const name = String(req.name || '').trim();
  withLock_(() => {
    const sh = memoSheet_();
    const n = sh.getLastRow();
    const rows = n < 2 ? [] : sh.getRange(2, 1, n - 1, MEMO_HEADERS.length).getDisplayValues();
    const i = rows.findIndex(r => r[0] === req.id && r[4] !== '삭제');
    if (i < 0) throw new Error('이미 지워진 메모입니다.');
    if (rows[i][2] !== name) throw new Error('메모를 쓴 분의 이름과 달라 지울 수 없습니다.');
    sh.getRange(i + 2, 5).setValue('삭제');
  });
  return getMemos();
}
