// 인증 모듈 - 클라이언트 사이드 PBKDF2 검증 + localStorage 세션
// 의존: users.js (window.USERS)

(function () {
  var SESSION_KEY = 'auth_session';
  var SESSION_HOURS = 8;
  var LOCK_KEY = 'auth_lock';
  var LOCK_THRESHOLD = 5;
  var LOCK_DURATION_MS = 60 * 1000;

  function hexToBytes(hex) {
    var bytes = new Uint8Array(hex.length / 2);
    for (var i = 0; i < hex.length; i += 2) {
      bytes[i / 2] = parseInt(hex.substr(i, 2), 16);
    }
    return bytes;
  }

  function bytesToHex(bytes) {
    var hex = '';
    for (var i = 0; i < bytes.length; i++) {
      hex += bytes[i].toString(16).padStart(2, '0');
    }
    return hex;
  }

  // PBKDF2 해시 (Web Crypto API)
  async function pbkdf2(password, saltHex, iterations) {
    var enc = new TextEncoder();
    var key = await crypto.subtle.importKey(
      'raw', enc.encode(password), { name: 'PBKDF2' }, false, ['deriveBits']
    );
    var bits = await crypto.subtle.deriveBits(
      { name: 'PBKDF2', salt: hexToBytes(saltHex), iterations: iterations, hash: 'SHA-256' },
      key, 256
    );
    return bytesToHex(new Uint8Array(bits));
  }

  function randomSaltHex() {
    var bytes = new Uint8Array(16);
    crypto.getRandomValues(bytes);
    return bytesToHex(bytes);
  }

  // 사용자의 현재 인증 데이터 (localStorage 변경 우선, 없으면 USERS 초기값)
  function getUserAuthData(userId) {
    var override = localStorage.getItem('pwd_override_' + userId);
    if (override) {
      try { return JSON.parse(override); } catch (e) {}
    }
    return window.USERS && window.USERS[userId] ? window.USERS[userId] : null;
  }

  // 로그인 시도 잠금 상태 체크
  function getLockState(userId) {
    var raw = localStorage.getItem(LOCK_KEY + '_' + userId);
    if (!raw) return { count: 0, lockedUntil: 0 };
    try { return JSON.parse(raw); } catch (e) { return { count: 0, lockedUntil: 0 }; }
  }

  function setLockState(userId, state) {
    localStorage.setItem(LOCK_KEY + '_' + userId, JSON.stringify(state));
  }

  function clearLockState(userId) {
    localStorage.removeItem(LOCK_KEY + '_' + userId);
  }

  // 로그인
  // returns { ok: true } or { ok: false, error: '...', lockedSec: N }
  async function login(userId, password) {
    userId = (userId || '').trim();
    if (!userId || !password) return { ok: false, error: 'ID와 비밀번호를 입력하세요.' };

    var lock = getLockState(userId);
    var now = Date.now();
    if (lock.lockedUntil > now) {
      var sec = Math.ceil((lock.lockedUntil - now) / 1000);
      return { ok: false, error: sec + '초 뒤에 다시 시도하세요.', lockedSec: sec };
    }

    var userData = getUserAuthData(userId);
    if (!userData) {
      // 등록되지 않은 ID에도 시도 횟수 누적 (timing attack 방지)
      var newLock = { count: lock.count + 1, lockedUntil: 0 };
      if (newLock.count >= LOCK_THRESHOLD) {
        newLock.lockedUntil = now + LOCK_DURATION_MS;
        newLock.count = 0;
      }
      setLockState(userId, newLock);
      return { ok: false, error: '잘못된 ID 또는 비밀번호입니다.' };
    }

    var inputHash = await pbkdf2(password, userData.salt, userData.iterations || 100000);
    if (inputHash !== userData.hash) {
      var newLock2 = { count: lock.count + 1, lockedUntil: 0 };
      if (newLock2.count >= LOCK_THRESHOLD) {
        newLock2.lockedUntil = now + LOCK_DURATION_MS;
        newLock2.count = 0;
      }
      setLockState(userId, newLock2);
      var remaining = LOCK_THRESHOLD - newLock2.count;
      return {
        ok: false,
        error: '잘못된 ID 또는 비밀번호입니다.' + (remaining > 0 && remaining <= 2 ? ' (' + remaining + '회 남음)' : '')
      };
    }

    // 성공 → 세션 생성
    clearLockState(userId);
    var session = {
      userId: userId,
      name: window.USERS[userId] ? window.USERS[userId].name : userId,
      expires: now + SESSION_HOURS * 3600 * 1000
    };
    localStorage.setItem(SESSION_KEY, JSON.stringify(session));
    return { ok: true, user: session };
  }

  function logout() {
    localStorage.removeItem(SESSION_KEY);
  }

  function getCurrentUser() {
    var raw = localStorage.getItem(SESSION_KEY);
    if (!raw) return null;
    try {
      var s = JSON.parse(raw);
      if (s.expires && s.expires < Date.now()) {
        localStorage.removeItem(SESSION_KEY);
        return null;
      }
      return s;
    } catch (e) { return null; }
  }

  // 미인증 시 login.html로 이동
  function requireAuth() {
    var user = getCurrentUser();
    if (!user) {
      var current = encodeURIComponent(location.pathname + location.search);
      location.replace('login.html?next=' + current);
      return null;
    }
    return user;
  }

  // 비밀번호 변경
  // returns { ok: true } or { ok: false, error: '...' }
  async function changePassword(oldPassword, newPassword) {
    var user = getCurrentUser();
    if (!user) return { ok: false, error: '로그인이 만료되었습니다.' };

    if (!newPassword || newPassword.length < 6) {
      return { ok: false, error: '새 비밀번호는 최소 6자 이상이어야 합니다.' };
    }
    if (newPassword === oldPassword) {
      return { ok: false, error: '새 비밀번호가 기존과 동일합니다.' };
    }

    var userData = getUserAuthData(user.userId);
    if (!userData) return { ok: false, error: '사용자 정보를 찾을 수 없습니다.' };

    var oldHash = await pbkdf2(oldPassword, userData.salt, userData.iterations || 100000);
    if (oldHash !== userData.hash) {
      return { ok: false, error: '현재 비밀번호가 일치하지 않습니다.' };
    }

    var newSalt = randomSaltHex();
    var newHash = await pbkdf2(newPassword, newSalt, 100000);
    var newData = { salt: newSalt, hash: newHash, iterations: 100000 };
    localStorage.setItem('pwd_override_' + user.userId, JSON.stringify(newData));
    return { ok: true };
  }

  // 사용자명 표시 — USERS에서 최신 이름/소속 조회 (변경 시 자동 반영)
  // 소속이 있으면 "생산3팀 이동준" 형태, 없으면 "이동준"만
  function getDisplayName() {
    var u = getCurrentUser();
    if (!u) return '';
    var info = window.USERS && window.USERS[u.userId] ? window.USERS[u.userId] : null;
    if (info) {
      var name = info.name || u.userId;
      return info.dept ? info.dept + ' ' + name : name;
    }
    return u.name || u.userId;
  }

  // 관리자 여부
  function isAdmin() {
    var u = getCurrentUser();
    if (!u) return false;
    return !!(window.USERS && window.USERS[u.userId] && window.USERS[u.userId].admin);
  }

  window.Auth = {
    login: login,
    logout: logout,
    getCurrentUser: getCurrentUser,
    requireAuth: requireAuth,
    changePassword: changePassword,
    getDisplayName: getDisplayName,
    isAdmin: isAdmin
  };

  // 세션 만료 자동 감시 (1분마다 체크) — 페이지를 켜둔 채로 만료 시각 도래해도 자동 로그아웃
  setInterval(function () {
    var raw = localStorage.getItem(SESSION_KEY);
    if (!raw) return;
    try {
      var s = JSON.parse(raw);
      if (s.expires && s.expires < Date.now()) {
        localStorage.removeItem(SESSION_KEY);
        // 로그인 페이지가 아닌 경우에만 redirect
        if (location.pathname.indexOf('login.html') === -1) {
          alert('로그인이 만료되었습니다. 다시 로그인해 주세요.');
          location.replace('login.html');
        }
      }
    } catch (e) {}
  }, 60 * 1000);
})();
