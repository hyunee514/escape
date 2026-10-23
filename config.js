/**
 * 설비보전 방탈출 — 서버 (구글 앱스 스크립트 + 구글 시트)
 * 만든이: 성현쌤
 *
 * 하는 일: 회원가입 / 로그인 / 기록 저장 / 명예의 전당(랭킹: 최고 층 · 층별 최단 기록 · 완등 합계 · 정답률 · 최다 풀이)
 * 데이터: 이 스크립트가 붙어 있는 구글 시트의 'users' 시트에 저장돼요.
 *
 * ★ 처음 한 번만 하면 되는 설정 (자세한 순서는 '서버설정방법.md' 참고)
 *   1) 아래 ADMIN_ID / ADMIN_FIRST_PW 를 원하는 값으로 바꾸고
 *   2) 위쪽 함수 선택 칸에서 createAdminAccount 를 골라 ▶ 실행 → 권한 허용
 *   3) 실행이 끝나면 ADMIN_FIRST_PW 를 지워 두세요 (비밀번호가 코드에 남지 않게)
 *   4) 배포 → 새 배포 → 유형: 웹 앱 / 실행: 나 / 액세스: 모든 사용자 → 배포 → 웹 앱 URL 복사
 */

// ===== 설정 =====
const ADMIN_ID = "임성현08";               // 관리자 아이디 (일반 가입으로는 만들 수 없어요)
const ADMIN_FIRST_PW = "여기에-비밀번호";    // createAdminAccount() 실행할 때만 쓰고, 끝나면 지우세요
const ADMIN_NICK = "성현쌤";
const TOKEN_DAYS = 30;                     // 자동 로그인 유지 기간
const MAX_DEVICES = 100;                   // 한 계정으로 동시에 로그인해 둘 수 있는 기기 수
const SHEET_NAME = "users";
const RANK_TOP = 10;                       // 명예의 전당에 보여줄 인원
const MIN_ANSWERED_FOR_ACC = 50;           // 정답률 랭킹에 들어가는 최소 문제 수
const TOP_FLOOR = 30;                      // '완등 합계' 랭킹 기준 층 수 (층을 늘리면 config.js 의 MAX_FLOOR 와 함께 바꿔 주세요)
const MIN_FLOOR_SEC = 15;                  // 이 시간보다 빠른 층 기록은 부정으로 보고 받지 않아요 (층별 최소는 minSec_ 가 따로 계산)
const FLOOR_SLACK_SEC = 10;                // 서버 측정 시간과 기록 사이에 허용하는 오차(통신 지연 등)
const RANK_CACHE_SEC = 60;                 // 랭킹을 캐시에 두는 시간(초). 새 최단 기록이 저장되면 바로 비워요
const PROOF_TTL_SEC = 21600;               // 층 시작·클리어 확인을 서버가 기억하는 시간 (캐시 최대 6시간)

const COLS = ["id", "nick", "salt", "pwhash", "token", "tokenExp", "createdAt", "lastLogin",
  "maxFloor", "maxFloorAt", "totalBest", "answered", "correct", "bestJson", "runs", "updatedAt", "saveJson"];
const C = {}; COLS.forEach(function (n, i) { C[n] = i; });

// ===== 진입점 =====
function doGet(e) {
  return json_({ ok: true, message: "설비보전 방탈출 서버가 정상 작동 중이에요. (만든이: 성현쌤)", now: Date.now() });
}

function doPost(e) {
  var res;
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    // 층 시작/클리어 확인은 시트를 건드리지 않으니 잠금 없이 바로 처리해요 (동시 접속이 많아도 가벼워요)
    if (body.action === "floorstart") return json_(floorStart_(body));
    if (body.action === "floorclear") return json_(floorClear_(body));
    if (body.action === "ranks") return json_(ranks_(body));            // 읽기 전용 + 캐시라서 저장 요청을 기다리지 않아요
    var lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      switch (body.action) {
        case "register": res = register_(body); break;
        case "login": res = login_(body); break;
        case "resume": res = resume_(body); break;
        case "save": res = save_(body); break;
        case "changepw": res = changePw_(body); break;
        case "logout": res = logout_(body); break;
        case "reset": res = reset_(body); break;
        default: res = { ok: false, error: "알 수 없는 요청이에요." };
      }
    } finally { lock.releaseLock(); }
  } catch (err) {
    res = { ok: false, error: "서버 오류: " + (err && err.message ? err.message : err) };
  }
  return json_(res);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ===== 시트 도우미 =====
function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, COLS.length).setValues([COLS]);
    sh.setFrozenRows(1);
  }
  return sh;
}
function allRows_(sh) {
  var n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, COLS.length).getValues();
}
function rowAt_(sh, row) { return sh.getRange(row, 1, 1, COLS.length).getValues()[0]; }
function findRowById_(sh, id) {
  var n = sh.getLastRow() - 1; if (n < 1) return null;
  var col = sh.getRange(2, C.id + 1, n, 1).getValues(), key = String(id).toLowerCase();       // 아이디 열만 읽어요
  for (var i = 0; i < n; i++) if (String(col[i][0]).toLowerCase() === key) return { row: i + 2, vals: rowAt_(sh, i + 2) };
  return null;
}
function findRowByToken_(sh, token) {
  if (!token) return null;
  var n = sh.getLastRow() - 1; if (n < 1) return null;
  var col = sh.getRange(2, C.token + 1, n, 1).getValues(), tmp = [];                          // 토큰 열만 읽어요
  for (var i = 0; i < n; i++) {
    if (String(col[i][0]).indexOf(token) < 0) continue;
    tmp[C.token] = col[i][0];
    if (hasToken_(tmp, token)) return { row: i + 2, vals: rowAt_(sh, i + 2) };
  }
  return null;
}
// ----- 랭킹용 읽기 : 필요한 열만 읽고 60초 캐시 (큰 저장본 JSON 열은 읽지 않아요) -----
function cachePutBig_(key, str, ttl) {
  var size = 30000, n = Math.ceil(str.length / size), obj = {};
  for (var i = 0; i < n; i++) obj[key + ":" + i] = str.substr(i * size, size);
  obj[key + ":n"] = String(n);
  try { CacheService.getScriptCache().putAll(obj, ttl); } catch (e) {}
}
function cacheGetBig_(key) {
  var c = CacheService.getScriptCache(), n = Number(c.get(key + ":n")); if (!n) return null;
  var keys = []; for (var i = 0; i < n; i++) keys.push(key + ":" + i);
  var got = c.getAll(keys), out = "";
  for (var j = 0; j < n; j++) { var part = got[key + ":" + j]; if (part == null) return null; out += part; }
  return out;
}
function rankCacheClear_() { try { CacheService.getScriptCache().remove("ranks_rows:n"); } catch (e) {} }
function rankRows_(sh) {
  var hit = cacheGetBig_("ranks_rows");
  if (hit) { try { return JSON.parse(hit).map(function (a) { var r = []; for (var i = 0; i < COLS.length; i++) r.push(""); r[C.id] = a[0]; r[C.nick] = a[1]; r[C.maxFloor] = a[2]; r[C.maxFloorAt] = a[3]; r[C.totalBest] = a[4]; r[C.answered] = a[5]; r[C.correct] = a[6]; r[C.bestJson] = a[7]; return r; }); } catch (e) {} }
  var n = sh.getLastRow() - 1; if (n < 1) return [];
  var ab = sh.getRange(2, C.id + 1, n, 2).getValues();                                        // 아이디·닉네임
  var cd = sh.getRange(2, C.maxFloor + 1, n, C.bestJson - C.maxFloor + 1).getValues();        // 최고 층 ~ 층별 기록
  var slim = [], rows = [];
  for (var i = 0; i < n; i++) {
    var r = []; for (var k = 0; k < COLS.length; k++) r.push("");
    r[C.id] = ab[i][0]; r[C.nick] = ab[i][1];
    for (var j = 0; j < cd[i].length; j++) r[C.maxFloor + j] = cd[i][j];
    rows.push(r);
    slim.push([r[C.id], r[C.nick], r[C.maxFloor], r[C.maxFloorAt], r[C.totalBest], r[C.answered], r[C.correct], r[C.bestJson]]);
  }
  cachePutBig_("ranks_rows", JSON.stringify(slim), RANK_CACHE_SEC);
  return rows;
}
// 로그인 토큰은 기기 MAX_DEVICES대까지 동시에 유지돼요 (JSON 목록으로 저장)
function parseTokens_(v) { try { var a = JSON.parse(v || "[]"); return Array.isArray(a) ? a : []; } catch (e) { return []; } }
function hasToken_(vals, token) { var now = Date.now(); return parseTokens_(vals[C.token]).some(function (x) { return x.t === token && x.e > now; }); }
function addToken_(vals, token) {
  var now = Date.now(), a = parseTokens_(vals[C.token]).filter(function (x) { return x.e > now && x.t !== token; });
  a.push({ t: token, e: now + TOKEN_DAYS * 86400000 }); if (a.length > MAX_DEVICES) a = a.slice(a.length - MAX_DEVICES);
  vals[C.token] = JSON.stringify(a); vals[C.tokenExp] = 0;
}
function dropToken_(vals, token) { vals[C.token] = JSON.stringify(parseTokens_(vals[C.token]).filter(function (x) { return x.t !== token; })); }
function writeRow_(sh, row, vals) { sh.getRange(row, 1, 1, COLS.length).setValues([vals]); }

// ===== 보안 도우미 =====
function hash_(salt, pw) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + "|" + pw);
  return bytes.map(function (b) { b = b < 0 ? b + 256 : b; return ("0" + b.toString(16)).slice(-2); }).join("");
}
function isAdminId_(id) { return String(id).toLowerCase() === String(ADMIN_ID).toLowerCase(); }
function newToken_() { return Utilities.getUuid() + Utilities.getUuid().replace(/-/g, ""); }
function tooMany_(id) {
  var cache = CacheService.getScriptCache(), k = "fail_" + String(id).toLowerCase(), n = Number(cache.get(k) || 0);
  return n >= 6;
}
function noteFail_(id) {
  var cache = CacheService.getScriptCache(), k = "fail_" + String(id).toLowerCase(), n = Number(cache.get(k) || 0);
  cache.put(k, String(n + 1), 600);
}
function clearFail_(id) { CacheService.getScriptCache().remove("fail_" + String(id).toLowerCase()); }
function cleanNick_(s) { return String(s || "").replace(/[<>&"'`]/g, "").trim(); }

function userInfo_(vals) { return { id: vals[C.id], nick: vals[C.nick], admin: isAdminId_(vals[C.id]) }; }
function parseSave_(vals) { var s = vals[C.saveJson]; if (!s) return null; try { return JSON.parse(s); } catch (e) { return null; } }
function parseBest_(vals) { try { var o = JSON.parse(vals[C.bestJson] || "{}"); return o && typeof o === "object" ? o : {}; } catch (e) { return {}; } }

// ===== 회원가입 / 로그인 =====
function register_(b) {
  var id = String(b.id || "").trim(), pw = String(b.pw || ""), nick = cleanNick_(b.nick);
  if (!/^[A-Za-z0-9가-힣_]{3,16}$/.test(id)) return { ok: false, error: "아이디는 3~16자의 한글·영문·숫자·_ 만 쓸 수 있어요." };
  if (pw.length < 4 || pw.length > 30) return { ok: false, error: "비밀번호는 4~30자로 정해 주세요." };
  if (nick.length < 2 || nick.length > 12) return { ok: false, error: "별명은 2~12자로 정해 주세요." };
  if (isAdminId_(id)) return { ok: false, error: "사용할 수 없는 아이디예요." };
  var sh = sheet_(), rows = allRows_(sh);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][C.id]).toLowerCase() === id.toLowerCase()) return { ok: false, error: "이미 있는 아이디예요." };
    if (String(rows[i][C.nick]).toLowerCase() === nick.toLowerCase()) return { ok: false, error: "이미 있는 별명이에요. 다른 별명을 써 주세요." };
  }
  var salt = Utilities.getUuid(), token = newToken_(), now = Date.now();
  var vals = COLS.map(function () { return ""; });
  vals[C.id] = id; vals[C.nick] = nick; vals[C.salt] = salt; vals[C.pwhash] = hash_(salt, pw);
  addToken_(vals, token); vals[C.createdAt] = now; vals[C.lastLogin] = now;
  vals[C.maxFloor] = 0; vals[C.totalBest] = 0; vals[C.answered] = 0; vals[C.correct] = 0; vals[C.bestJson] = "{}"; vals[C.runs] = 0; vals[C.updatedAt] = now;
  sh.appendRow(vals);
  return { ok: true, token: token, user: userInfo_(vals), save: null, rev: now, now: now };
}

function login_(b) {
  var id = String(b.id || "").trim(), pw = String(b.pw || "");
  if (!id || !pw) return { ok: false, error: "아이디와 비밀번호를 입력해 주세요." };
  if (tooMany_(id)) return { ok: false, error: "비밀번호를 여러 번 틀렸어요. 10분 뒤에 다시 시도해 주세요." };
  var sh = sheet_(), f = findRowById_(sh, id);
  if (!f || hash_(f.vals[C.salt], pw) !== f.vals[C.pwhash]) { noteFail_(id); return { ok: false, error: "아이디 또는 비밀번호가 맞지 않아요." }; }
  clearFail_(id);
  var token = newToken_(), now = Date.now();
  addToken_(f.vals, token); f.vals[C.lastLogin] = now;
  writeRow_(sh, f.row, f.vals);
  return { ok: true, token: token, user: userInfo_(f.vals), save: parseSave_(f.vals), rev: Number(f.vals[C.updatedAt]) || 0, now: now };
}

function resume_(b) {
  var sh = sheet_(), f = findRowByToken_(sh, b.token);
  if (!f) return { ok: false, error: "auth" };
  var now = Date.now();
  f.vals[C.lastLogin] = now; addToken_(f.vals, b.token);
  writeRow_(sh, f.row, f.vals);
  return { ok: true, token: b.token, user: userInfo_(f.vals), save: parseSave_(f.vals), rev: Number(f.vals[C.updatedAt]) || 0, now: now };
}

function logout_(b) {
  var sh = sheet_(), f = findRowByToken_(sh, b.token);
  if (f) { dropToken_(f.vals, b.token); writeRow_(sh, f.row, f.vals); }
  return { ok: true };
}

function changePw_(b) {
  var sh = sheet_(), f = findRowByToken_(sh, b.token);
  if (!f) return { ok: false, error: "auth" };
  var np = String(b.newPw || "");
  if (hash_(f.vals[C.salt], String(b.oldPw || "")) !== f.vals[C.pwhash]) return { ok: false, error: "현재 비밀번호가 맞지 않아요." };
  if (np.length < 4 || np.length > 30) return { ok: false, error: "새 비밀번호는 4~30자로 정해 주세요." };
  var salt = Utilities.getUuid();
  f.vals[C.salt] = salt; f.vals[C.pwhash] = hash_(salt, np);
  f.vals[C.token] = "[]"; addToken_(f.vals, b.token);                  // 비밀번호를 바꾸면 다른 기기는 다시 로그인해야 해요
  writeRow_(sh, f.row, f.vals);
  return { ok: true };
}

// ===== 기록 저장 =====
function num_(x, lo, hi) { x = Number(x); if (!isFinite(x)) x = 0; return Math.max(lo, Math.min(hi, Math.floor(x))); }

// ===== 층 시간 검증 (클라이언트가 보낸 시간을 그대로 믿지 않기 위한 장치) =====
// 흐름: 층 시작 → 서버가 시각 기록 / 클리어 → 서버가 경과 시간을 재서 '인정 기록'을 캐시에 남김 / 저장 → 인정받은 기록만 최단 기록으로 반영
// ※ 보너스(시간 단축)는 클라이언트 규칙과 같아요: 층당 min(4, 1+floor(n/4))개 × (15 + min(15, n))초
function maxBonus_(n) { return Math.min(4, 1 + Math.floor(n / 4)) * (15 + Math.min(15, n)); }
function digits_(n) { return n <= 5 ? 3 : n <= 15 ? 4 : 5; }
function locks_(n) { return Math.min(6, Math.floor((n + 1) / 3)); }
// 층별 최소 기록 : 기본 15초 + 코드 숫자·잠긴 문 수에 비례. 실제 학생 기록을 보고 조정하세요.
function minSec_(n) { return Math.max(MIN_FLOOR_SEC, 10 + digits_(n) * 4 + locks_(n) * 3); }
function proofKey_(kind, token, fl) { return kind + ":" + String(token) + ":" + fl; }
function tokenOk_(t) { t = String(t || ""); return t.length >= 20 && t.length <= 200; }

function floorStart_(b) {
  var fl = Math.floor(Number(b.floor));
  if (!tokenOk_(b.token) || !(fl >= 1 && fl <= 99)) return { ok: false, error: "잘못된 요청이에요." };
  CacheService.getScriptCache().put(proofKey_("fs", b.token, fl), String(Date.now()), PROOF_TTL_SEC);
  return { ok: true, now: Date.now() };
}
function floorClear_(b) {
  var fl = Math.floor(Number(b.floor)), t = Math.round(Number(b.t) * 10) / 10;
  if (!tokenOk_(b.token) || !(fl >= 1 && fl <= 99) || !isFinite(t)) return { ok: false, error: "잘못된 요청이에요." };
  var cache = CacheService.getScriptCache(), now = Date.now(), st = Number(cache.get(proofKey_("fs", b.token, fl)));
  var floorMin = minSec_(fl), need = floorMin;
  if (st > 0) need = Math.max(floorMin, (now - st) / 1000 - maxBonus_(fl) - FLOOR_SLACK_SEC);   // 시작을 서버가 알고 있으면 실제 경과 시간으로 따져요
  if (t < need || t > 86400) return { ok: true, accepted: false, need: Math.ceil(need) };
  var old = Number(cache.get(proofKey_("ok", b.token, fl))) || 0;
  if (!old || t < old) cache.put(proofKey_("ok", b.token, fl), String(t), PROOF_TTL_SEC);
  cache.remove(proofKey_("fs", b.token, fl));                                                  // 한 번 시작으로 한 번만 인정
  return { ok: true, accepted: true };
}

function save_(b) {
  var sh = sheet_(), f = findRowByToken_(sh, b.token);
  if (!f) return { ok: false, error: "auth" };
  var s = b.save;
  if (!s || typeof s !== "object" || !s.stats) return { ok: false, error: "저장 형식이 올바르지 않아요." };
  var json = JSON.stringify(s);
  if (json.length > 45000) return { ok: false, error: "저장 데이터가 너무 커요." };
  var serverRev = Number(f.vals[C.updatedAt]) || 0;
  if (!b.force && Number(b.rev) !== serverRev && serverRev > 0) {
    return { ok: false, error: "conflict", save: parseSave_(f.vals), rev: serverRev };      // 다른 기기에서 더 최근에 저장됨
  }
  var a = num_(s.stats.a, 0, 1e7), c = num_(s.stats.c, 0, 1e7);
  if (c > a) return { ok: false, error: "저장 형식이 올바르지 않아요." };
  var prevA = Number(f.vals[C.answered]) || 0;
  if (!b.reset && a < prevA - 5) return { ok: false, error: "기록이 줄어드는 저장은 받을 수 없어요. (초기화는 '기록 초기화'를 쓰세요)" };

  // 층별 최단 기록 : 서버에 있던 기록과 비교해 더 빠른 것만 남겨요 (너무 빠른 기록은 거절)
  var oldBest = parseBest_(f.vals), inBest = (s.best && typeof s.best === "object") ? s.best : {}, best = {}, k, v;
  for (k in oldBest) { v = Number(oldBest[k]); if (v > 0) best[k] = v; }
  for (k in inBest) {
    var fl = Number(k); v = Number(inBest[k]);
    if (!(fl >= 1 && fl <= 99) || !isFinite(v)) continue;
    if (v < MIN_FLOOR_SEC || v > 86400) continue;
    v = Math.round(v * 10) / 10;
    if (!best[fl] || v < best[fl]) {
      // 새 최단 기록은 서버가 클리어를 직접 확인(floorclear)해 준 값보다 빠를 수 없어요
      var proof = Number(CacheService.getScriptCache().get(proofKey_("ok", b.token, fl))) || 0;
      if (!proof || v < proof - 0.05) continue;
      best[fl] = v;
    }
  }
  // 최고 층 = 1층부터 빠짐없이 기록이 있는 가장 높은 층
  var maxFloor = 0; while (best[maxFloor + 1]) maxFloor++;
  var total = 0; for (var i = 1; i <= maxFloor; i++) total += best[i];
  var now = Date.now();
  if (maxFloor > (Number(f.vals[C.maxFloor]) || 0)) f.vals[C.maxFloorAt] = now;
  f.vals[C.maxFloor] = maxFloor; f.vals[C.totalBest] = Math.round(total * 10) / 10;
  var bestJson = JSON.stringify(best); if (bestJson !== String(f.vals[C.bestJson] || "{}")) rankCacheClear_();
  f.vals[C.answered] = a; f.vals[C.correct] = c; f.vals[C.bestJson] = bestJson;
  f.vals[C.runs] = num_(s.runs, 0, 1e6); f.vals[C.updatedAt] = now;
  // 서버가 인정한 기록으로 저장본도 맞춰 줘요
  s.best = best; s.maxFloor = maxFloor; f.vals[C.saveJson] = JSON.stringify(s);
  writeRow_(sh, f.row, f.vals);
  return { ok: true, rev: now, save: s };
}

function reset_(b) {                 // 본인 기록 초기화
  var sh = sheet_(), f = findRowByToken_(sh, b.token);
  if (!f) return { ok: false, error: "auth" };
  ["maxFloor", "totalBest", "answered", "correct", "runs"].forEach(function (k) { f.vals[C[k]] = 0; });
  f.vals[C.maxFloorAt] = ""; f.vals[C.bestJson] = "{}"; f.vals[C.saveJson] = ""; f.vals[C.updatedAt] = Date.now();
  writeRow_(sh, f.row, f.vals); rankCacheClear_();
  return { ok: true, rev: f.vals[C.updatedAt] };
}

// ===== 명예의 전당 =====
function ranks_(b) {
  var sh = sheet_(), rows = rankRows_(sh), me = findRowByToken_(sh, b.token);
  rows = rows.filter(function (r) { return !isAdminId_(r[C.id]); });
  var n = function (r, k) { return Number(r[C[k]]) || 0; };
  var fmtSec = function (s) { var m = Math.floor(s / 60), x = s - m * 60; return m + ":" + (x < 10 ? "0" : "") + x.toFixed(1); };
  var acc = function (r) { return n(r, "answered") ? n(r, "correct") / n(r, "answered") : 0; };
  function build(key, title, unit, fmt, filter, cmp, valueOf, subOf) {
    var list = rows.filter(filter).sort(cmp);
    var out = list.slice(0, RANK_TOP).map(function (r, i) { return { rank: i + 1, nick: r[C.nick], value: valueOf(r), sub: subOf ? subOf(r) : "" }; });
    var myRank = 0;
    if (me) { for (var i = 0; i < list.length; i++) if (list[i][C.id] === me.vals[C.id]) { myRank = i + 1; break; } }
    return { key: key, title: title, unit: unit, fmt: fmt, list: out, myRank: myRank, total: list.length };
  }
  var cats = [
    build("height", "🏢 가장 높이 올라간 사람", "층", "", function (r) { return n(r, "maxFloor") > 0; },
      function (x, y) { return n(y, "maxFloor") - n(x, "maxFloor") || n(x, "maxFloorAt") - n(y, "maxFloorAt"); },
      function (r) { return n(r, "maxFloor"); }, function (r) { return "누적 최단 합계 " + fmtSec(n(r, "totalBest")); }),
    build("total", "⏱ 1~" + TOP_FLOOR + "층 완등 합계 시간", "", "time", function (r) { return n(r, "maxFloor") >= TOP_FLOOR; },
      function (x, y) { return sumTo_(x, TOP_FLOOR) - sumTo_(y, TOP_FLOOR); },
      function (r) { return Math.round(sumTo_(r, TOP_FLOOR) * 10) / 10; }, function (r) { return "푼 문제 " + n(r, "answered") + "개"; }),
    build("accuracy", "🎯 최고 정답률", "%", "", function (r) { return n(r, "answered") >= MIN_ANSWERED_FOR_ACC; },
      function (x, y) { return acc(y) - acc(x) || n(y, "answered") - n(x, "answered"); },
      function (r) { return Math.round(acc(r) * 1000) / 10; }, function (r) { return n(r, "answered") + "문제 중"; }),
    build("answered", "📚 최다 문제 풀이", "문제", "", function (r) { return n(r, "answered") > 0; },
      function (x, y) { return n(y, "answered") - n(x, "answered"); }, function (r) { return n(r, "answered"); }, function (r) { return "정답 " + n(r, "correct"); })
  ];
  // 층별 최단 기록
  var perFloor = {}, maxF = 0;
  rows.forEach(function (r) { var bb = parseBest_(r); for (var k in bb) { var fl = Number(k), v = Number(bb[k]); if (fl >= 1 && v > 0) { (perFloor[fl] = perFloor[fl] || []).push({ id: r[C.id], nick: r[C.nick], t: v }); if (fl > maxF) maxF = fl; } } });
  var floors = {};
  for (var fl in perFloor) {
    var arr = perFloor[fl].sort(function (x, y) { return x.t - y.t; }), myRankF = 0;
    if (me) { for (var i = 0; i < arr.length; i++) if (arr[i].id === me.vals[C.id]) { myRankF = i + 1; break; } }
    floors[fl] = { list: arr.slice(0, RANK_TOP).map(function (x, i) { return { rank: i + 1, nick: x.nick, value: x.t, sub: "" }; }), myRank: myRankF, total: arr.length };
  }
  return { ok: true, cats: cats, floors: floors, players: rows.filter(function (r) { return n(r, "answered") + n(r, "maxFloor") > 0; }).length, topFloor: TOP_FLOOR, now: Date.now() };
}
function sumTo_(r, top) { var bb = parseBest_(r), t = 0; for (var i = 1; i <= top; i++) t += Number(bb[i]) || 0; return t; }

// ===== 관리자 계정 만들기 (한 번만 실행) =====
function createAdminAccount() {
  var sh = sheet_();
  if (findRowById_(sh, ADMIN_ID)) { Logger.log("이미 관리자 계정이 있어요: " + ADMIN_ID); return; }
  var salt = Utilities.getUuid(), now = Date.now();
  var vals = COLS.map(function () { return ""; });
  vals[C.id] = ADMIN_ID; vals[C.nick] = ADMIN_NICK; vals[C.salt] = salt; vals[C.pwhash] = hash_(salt, ADMIN_FIRST_PW);
  vals[C.createdAt] = now; vals[C.lastLogin] = 0; vals[C.maxFloor] = 0; vals[C.totalBest] = 0; vals[C.answered] = 0; vals[C.correct] = 0; vals[C.bestJson] = "{}"; vals[C.runs] = 0;
  vals[C.updatedAt] = now; vals[C.tokenExp] = 0;
  sh.appendRow(vals);
  Logger.log("관리자 계정을 만들었어요: " + ADMIN_ID + " — 이제 ADMIN_FIRST_PW 값을 코드에서 지우세요.");
}

// ===== 도우미: 학생이 비밀번호를 잊었을 때 (편집기에서 직접 실행) =====
// 사용법: 아래 두 값을 바꾼 뒤, 함수 선택 칸에서 resetPasswordManually 를 골라 ▶ 실행하세요.
function resetPasswordManually() {
  var TARGET_ID = "학생아이디";        // 비밀번호를 바꿔 줄 학생의 아이디
  var NEW_PW = "새비밀번호";           // 임시 비밀번호 (학생에게 알려 주고, 로그인 후 '내 정보'에서 바꾸게 하세요)
  var sh = sheet_(), f = findRowById_(sh, TARGET_ID);
  if (!f) { Logger.log("그런 아이디가 없어요: " + TARGET_ID); return; }
  var salt = Utilities.getUuid();
  f.vals[C.salt] = salt; f.vals[C.pwhash] = hash_(salt, NEW_PW); f.vals[C.token] = "[]";
  writeRow_(sh, f.row, f.vals);
  Logger.log(TARGET_ID + " 의 비밀번호를 바꿨어요. (로그인된 기기는 모두 로그아웃돼요)");
}
