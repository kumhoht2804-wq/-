/**
 * 개미증권 – 구글 시트 서버 (Google Apps Script)
 *
 * 하는 일
 *  - 회원가입 / 로그인 (비밀번호는 솔트+SHA-256 해시로만 저장)
 *  - 게임 저장 / 불러오기 (계정당 1개, 시트의 save 칸)
 *  - 랭킹 (총자산 순 상위 50명)
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

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return json_({ ok: false, error: '서버가 바빠요. 잠시 후 다시 시도해주세요' });
  try {
    const sh = sheet_();
    switch (req.action) {
      case 'signup': return json_(signup_(sh, req));
      case 'login':  return json_(login_(sh, req));
      case 'load':   return json_(load_(sh, req));
      case 'save':   return json_(save_(sh, req));
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
