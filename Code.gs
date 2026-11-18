/**
 * 개미증권 – 구글 시트 서버 (Google Apps Script)
 *
 * 하는 일
 *  - 회원가입 / 로그인 (비밀번호는 솔트+SHA-256 해시로만 저장)
 *  - 게임 저장 / 불러오기 (계정당 1개, 시트의 save 칸)
 *  - 랭킹 (총자산 순 상위 50명)
 *  - 멀티 모드: 방 만들기/참가/시작, 실시간 순위, 찌라시 (CacheService 사용, 빠름)
 *
 * 설치: 구글 시트 → 확장 프로그램 → Apps Script → 이 코드 붙여넣기 → 배포 → 웹 앱
 *       (실행 사용자: 나 / 액세스 권한: 모든 사용자)
 */

const SHEET_NAME = 'users';
const HEAD = ['id', 'nickname', 'pwHash', 'salt', 'token', 'equity', 'returnPct', 'day', 'tier', 'save', 'createdAt', 'updatedAt'];
const COL = {};
HEAD.forEach(function (h, i) { COL[h] = i + 1; }); // 시트 열 번호(1부터)
const MAX_SAVE = 49000; // 시트 한 칸 최대 5만 자

// ---------------- 진입점 ----------------
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  if (action === 'ranking') return json_(ranking_());
  return json_({ ok: true, server: '개미증권', time: new Date().toISOString() });
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: '잘못된 요청이에요' });
  }
  if (req.action === 'ranking') return json_(ranking_());
  if (req.action === 'room_sync') {
    try { return json_(roomSync_(req)); } catch (err) { return json_({ ok: false, error: '서버 오류: ' + err }); }
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return json_({ ok: false, error: '서버가 바빠요. 잠시 후 다시 시도해주세요' });
  try {
    const sh = sheet_();
    switch (req.action) {
      case 'signup': return json_(signup_(sh, req));
      case 'login':  return json_(login_(sh, req));
      case 'load':   return json_(load_(sh, req));
      case 'save':   return json_(save_(sh, req));
      case 'room_create': return json_(roomCreate_(req));
      case 'room_join':   return json_(roomJoin_(req));
      case 'room_leave':  return json_(roomLeave_(req));
      case 'room_start':  return json_(roomStart_(req));
      case 'room_rumor':  return json_(roomRumor_(req));
      default:       return json_({ ok: false, error: '알 수 없는 요청이에요' });
    }
  } catch (err) {
    return json_({ ok: false, error: '서버 오류: ' + err });
  } finally {
    lock.releaseLock();
  }
}

// ---------------- 기능 ----------------
function signup_(sh, r) {
  const id = String(r.id || '').trim();
  const nick = String(r.nickname || '').trim();
  const pw = String(r.password || '');
  if (!/^[A-Za-z0-9_]{3,16}$/.test(id)) return { ok: false, error: '아이디는 영문·숫자·_ 3~16자로 만들어주세요' };
  if (nick.length < 2 || nick.length > 10) return { ok: false, error: '닉네임은 2~10자로 만들어주세요' };
  if (/^[=+\-@]/.test(nick)) return { ok: false, error: '닉네임은 = + - @ 로 시작할 수 없어요' };
  if (pw.length < 4) return { ok: false, error: '비밀번호는 4자 이상이어야 해요' };
  if (findRow_(sh, COL.id, id) > 0) return { ok: false, error: '이미 사용 중인 아이디예요' };
  if (findRow_(sh, COL.nickname, nick) > 0) return { ok: false, error: '이미 사용 중인 닉네임이에요' };

  const salt = Utilities.getUuid();
  const token = Utilities.getUuid();
  const now = new Date();
  sh.appendRow([id, nick, hash_(pw, salt), salt, token, 10000000, 0, 1, '🐜 개미', '', now, now]);
  cacheToken_(id, token, nick);
  return { ok: true, id: id, nickname: nick, token: token, save: null };
}

function login_(sh, r) {
  const id = String(r.id || '').trim();
  const row = findRow_(sh, COL.id, id);
  const fail = { ok: false, error: '아이디 또는 비밀번호가 틀렸어요' };
  if (row < 0) return fail;
  const v = sh.getRange(row, 1, 1, HEAD.length).getValues()[0];
  if (hash_(String(r.password || ''), v[COL.salt - 1]) !== v[COL.pwHash - 1]) return fail;
  const token = Utilities.getUuid(); // 새 기기에서 로그인하면 이전 기기는 로그아웃
  sh.getRange(row, COL.token).setValue(token);
  cacheToken_(v[0], token, v[COL.nickname - 1]);
  return { ok: true, id: v[0], nickname: v[COL.nickname - 1], token: token, save: v[COL.save - 1] || null };
}

function load_(sh, r) {
  const a = auth_(sh, r);
  if (!a) return { ok: false, error: '다른 기기에서 로그인했거나 세션이 만료됐어요', relogin: true };
  return { ok: true, id: a.v[0], nickname: a.v[COL.nickname - 1], save: a.v[COL.save - 1] || null };
}

function save_(sh, r) {
  const a = auth_(sh, r);
  if (!a) return { ok: false, error: '다른 기기에서 로그인했거나 세션이 만료됐어요', relogin: true };
  const save = String(r.save || '');
  if (save.length > MAX_SAVE) return { ok: false, error: '저장 데이터가 너무 커요' };
  const equity = Number(r.equity), ret = Number(r.returnPct), day = Number(r.day);
  if (!isFinite(equity) || !isFinite(ret) || !isFinite(day)) return { ok: false, error: '잘못된 저장 데이터예요' };
  sh.getRange(a.row, COL.equity, 1, 5).setValues([[Math.round(equity), ret, day, String(r.tier || ''), save]]);
  sh.getRange(a.row, COL.updatedAt).setValue(new Date());
  return { ok: true, savedAt: Date.now() };
}

function ranking_() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return { ok: true, list: [] };
  const rows = sh.getRange(2, 1, last - 1, HEAD.length).getValues();
  const list = rows
    .filter(function (v) { return v[0]; })
    .map(function (v) {
      return { nickname: v[COL.nickname - 1], equity: Number(v[COL.equity - 1]) || 0, returnPct: Number(v[COL.returnPct - 1]) || 0,
               day: Number(v[COL.day - 1]) || 1, tier: v[COL.tier - 1] };
    })
    .sort(function (a, b) { return b.equity - a.equity; })
    .slice(0, 50);
  return { ok: true, list: list, total: rows.length };
}


// ================= 멀티 모드 =================
// 방 정보와 플레이어 상태는 CacheService(최대 6시간 보관)에 둬서 빠르게 주고받고,
// 끝난 게임 결과만 matches 시트에 기록해요.
const ROOM_TTL = 21600, MAX_PLAYERS = 6, MAX_RUMORS = 3, MPD = 390;
function cache_() { return CacheService.getScriptCache(); }
function getRoom_(code) { const s = cache_().get('room_' + code); return s ? JSON.parse(s) : null; }
function putRoom_(room) { cache_().put('room_' + room.code, JSON.stringify(room), ROOM_TTL); }
function cacheToken_(id, token, nick) {
  const k = String(id).toLowerCase();
  cache_().putAll({ ['tok_' + k]: token, ['nick_' + k]: String(nick) }, ROOM_TTL);
}
function authFast_(r) {
  const id = String(r.id || '').trim();
  if (!id || !r.token) return null;
  const k = id.toLowerCase();
  const c = cache_().getAll(['tok_' + k, 'nick_' + k]);
  if (c['tok_' + k] === r.token) return { id: id, nick: c['nick_' + k] };
  const a = auth_(sheet_(), r);
  if (!a) return null;
  cacheToken_(a.v[0], r.token, a.v[COL.nickname - 1]);
  return { id: a.v[0], nick: a.v[COL.nickname - 1] };
}
function clock_(room, now) {
  const block = MPD * room.msPerMin + room.breakMs;
  const el = now - room.startAt;
  if (el < 0) return { d: 0, m: 0, started: false };
  const d = Math.floor(el / block);
  return { d: d, m: Math.min(MPD, Math.floor((el % block) / room.msPerMin)), started: true, ended: d >= room.days };
}
function publicRoom_(room) {
  const o = JSON.parse(JSON.stringify(room));
  if (o.status !== 'ended') o.rumors.forEach(function (x) { delete x.by; delete x.nick; }); // 끝나기 전엔 유포자 비공개
  return o;
}
function newCode_() {
  const A = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  for (let t = 0; t < 20; t++) {
    let c = '';
    for (let i = 0; i < 5; i++) c += A[Math.floor(Math.random() * A.length)];
    if (!getRoom_(c)) return c;
  }
  throw new Error('방 코드 생성 실패');
}
function roomCreate_(r) {
  const u = authFast_(r);
  if (!u) return { ok: false, error: '로그인이 필요해요', relogin: true };
  const days = [3, 5, 7].indexOf(Number(r.days)) >= 0 ? Number(r.days) : 5;
  const diff = ['easy', 'normal', 'hard'].indexOf(r.diff) >= 0 ? r.diff : 'normal';
  const room = { code: newCode_(), host: u.id, status: 'lobby', days: days, diff: diff, msPerMin: 500, breakMs: 8000,
                 players: [{ id: u.id, nick: u.nick }], rumors: [], createdAt: Date.now(), seed: 0, startAt: 0 };
  putRoom_(room);
  return { ok: true, room: publicRoom_(room), serverTime: Date.now() };
}
function roomJoin_(r) {
  const u = authFast_(r);
  if (!u) return { ok: false, error: '로그인이 필요해요', relogin: true };
  const code = String(r.code || '').trim().toUpperCase();
  const room = getRoom_(code);
  if (!room) return { ok: false, error: '방을 찾을 수 없어요. 코드를 확인해주세요' };
  const inRoom = room.players.some(function (p) { return p.id === u.id; });
  if (!inRoom) {
    if (room.status !== 'lobby') return { ok: false, error: '이미 시작한 방이에요' };
    if (room.players.length >= MAX_PLAYERS) return { ok: false, error: '방이 꽉 찼어요 (최대 6명)' };
    room.players.push({ id: u.id, nick: u.nick });
    putRoom_(room);
  }
  return { ok: true, room: publicRoom_(room), serverTime: Date.now() };
}
function roomLeave_(r) {
  const u = authFast_(r);
  if (!u) return { ok: true };
  const room = getRoom_(String(r.code || '').toUpperCase());
  if (!room || room.status !== 'lobby') return { ok: true }; // 진행 중엔 나가도 자리 유지(재접속 가능)
  room.players = room.players.filter(function (p) { return p.id !== u.id; });
  if (!room.players.length) { cache_().remove('room_' + room.code); return { ok: true }; }
  if (room.host === u.id) room.host = room.players[0].id;
  putRoom_(room);
  return { ok: true };
}
function roomStart_(r) {
  const u = authFast_(r);
  if (!u) return { ok: false, error: '로그인이 필요해요', relogin: true };
  const room = getRoom_(String(r.code || '').toUpperCase());
  if (!room) return { ok: false, error: '방을 찾을 수 없어요' };
  if (room.host !== u.id) return { ok: false, error: '방장만 시작할 수 있어요' };
  if (room.status !== 'lobby') return { ok: false, error: '이미 시작했어요' };
  room.status = 'playing';
  room.seed = Math.floor(Math.random() * 2147483646) + 1;
  room.startAt = Date.now() + 5000; // 5초 카운트다운
  putRoom_(room);
  return { ok: true, room: publicRoom_(room), serverTime: Date.now() };
}
function roomRumor_(r) {
  const u = authFast_(r);
  if (!u) return { ok: false, error: '로그인이 필요해요', relogin: true };
  const room = getRoom_(String(r.code || '').toUpperCase());
  if (!room || room.status !== 'playing') return { ok: false, error: '진행 중인 게임이 아니에요' };
  if (!room.players.some(function (p) { return p.id === u.id; })) return { ok: false, error: '이 방의 참가자가 아니에요' };
  const used = room.rumors.filter(function (x) { return x.by === u.id; }).length;
  if (used >= MAX_RUMORS) return { ok: false, error: '찌라시를 모두 사용했어요' };
  const s = Number(r.stock), dir = Number(r.dir), tpl = Number(r.tpl);
  if (!(s >= 0 && s < 60) || (dir !== 1 && dir !== -1) || !(tpl >= 0 && tpl < 3)) return { ok: false, error: '잘못된 찌라시예요' };
  const c = clock_(room, Date.now());
  let day = c.d, min = c.m + 6; // 약 3초 뒤 모두에게 동시에 퍼짐
  if (min >= MPD - 10) { day += 1; min = 8; }
  if (day >= room.days) return { ok: false, error: '게임이 곧 끝나서 찌라시를 뿌릴 수 없어요' };
  const rumor = { rid: room.rumors.length + 1, by: u.id, nick: u.nick, s: s, dir: dir, tpl: tpl, day: day + 1, min: min };
  room.rumors.push(rumor);
  putRoom_(room);
  return { ok: true, rumor: { rid: rumor.rid, s: s, dir: dir, tpl: tpl, day: rumor.day, min: min }, left: MAX_RUMORS - used - 1 };
}
function roomSync_(r) {
  const u = authFast_(r);
  if (!u) return { ok: false, error: '로그인이 필요해요', relogin: true };
  const code = String(r.code || '').toUpperCase();
  let room = getRoom_(code);
  if (!room) return { ok: false, error: '방이 사라졌어요', gone: true };
  if (!room.players.some(function (p) { return p.id === u.id; })) return { ok: false, error: '이 방의 참가자가 아니에요', gone: true };
  const now = Date.now();
  const myKey = 'rp_' + code + '_' + u.id.toLowerCase();
  let mine = null;
  if (r.restore) { const m = cache_().get(myKey); mine = m ? JSON.parse(m) : null; }
  if (r.st && typeof r.st === 'object') {
    const st = r.st; st.t = now;
    const str = JSON.stringify(st);
    if (str.length < 90000) cache_().put(myKey, str, ROOM_TTL);
  }
  // 게임 시간이 끝났으면 종료 처리 + 결과 기록(한 번만)
  if (room.status === 'playing' && clock_(room, now).ended && now > room.startAt + 1) {
    const lock = LockService.getScriptLock();
    if (lock.tryLock(3000)) {
      try {
        room = getRoom_(code);
        if (room.status === 'playing') { room.status = 'ended'; room.endedAt = now; putRoom_(room); room._record = true; }
      } finally { lock.releaseLock(); }
    }
  }
  const keys = room.players.map(function (p) { return 'rp_' + code + '_' + p.id.toLowerCase(); });
  const got = cache_().getAll(keys);
  const players = room.players.map(function (p, i) {
    const st = got[keys[i]] ? JSON.parse(got[keys[i]]) : {};
    return { id: p.id, nick: p.nick, host: p.id === room.host, equity: st.equity || 10000000, ret: st.ret || 0,
             bankrupt: !!st.bankrupt, final: !!st.final, online: !!st.t && now - st.t < 8000 };
  });
  if (room._record) {
    delete room._record;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let mh = ss.getSheetByName('matches');
    if (!mh) { mh = ss.insertSheet('matches'); mh.appendRow(['code', 'endedAt', 'days', 'players', 'winner', 'results']); mh.setFrozenRows(1); }
    const alive = players.filter(function (p) { return !p.bankrupt; }).sort(function (a, b) { return b.equity - a.equity; });
    mh.appendRow([code, new Date(now), room.days, players.length, alive.length ? alive[0].nick : '(전원 파산)',
                  JSON.stringify(players.map(function (p) { return { nick: p.nick, equity: p.equity, bankrupt: p.bankrupt }; }))]);
  }
  return { ok: true, serverTime: now, room: publicRoom_(room), players: players, mine: mine };
}

// ---------------- 도우미 ----------------
function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
    sh.getRange('A:B').setNumberFormat('@'); // 아이디·닉네임은 항상 글자로
  }
  return sh;
}

function auth_(sh, r) {
  const row = findRow_(sh, COL.id, String(r.id || '').trim());
  if (row < 0 || !r.token) return null;
  const v = sh.getRange(row, 1, 1, HEAD.length).getValues()[0];
  if (v[COL.token - 1] !== r.token) return null;
  return { row: row, v: v };
}

function findRow_(sh, col, value) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const vals = sh.getRange(2, col, last - 1, 1).getValues();
  const want = String(value).toLowerCase();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).toLowerCase() === want) return i + 2;
  }
  return -1;
}

function hash_(pw, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pw, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(bytes);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
