(function authClientBootstrap() {
  const ACCESS_PASSWORD = "123456";
  const AUTH_KEY = "dashboard_auth_v1";
  const AUTH_TTL_MS = 8 * 60 * 60 * 1000;

  function readAuthPayload() {
    try {
      const raw = localStorage.getItem(AUTH_KEY);
      if (!raw) {
        return null;
      }
      const parsed = JSON.parse(raw);
      if (!parsed || typeof parsed !== "object") {
        return null;
      }
      return parsed;
    } catch (_) {
      return null;
    }
  }

  function writeAuthPayload(payload) {
    localStorage.setItem(AUTH_KEY, JSON.stringify(payload));
  }

  function clearAuthPayload() {
    localStorage.removeItem(AUTH_KEY);
  }

  function getAuthStatus() {
    const payload = readAuthPayload();
    const now = Date.now();

    if (!payload || !Number.isFinite(payload.expiresAt) || payload.expiresAt <= now) {
      clearAuthPayload();
      return {
        authenticated: false,
        expiresAt: null,
        remainingMs: 0
      };
    }

    return {
      authenticated: true,
      expiresAt: payload.expiresAt,
      remainingMs: payload.expiresAt - now
    };
  }

  function hasDashboardAccess() {
    return Promise.resolve(getAuthStatus().authenticated);
  }

  function loginWithPassword(password) {
    const input = String(password || "");
    if (!input) {
      return Promise.resolve({
        ok: false,
        status: 400,
        data: { message: "密码不能为空。" }
      });
    }

    if (input !== ACCESS_PASSWORD) {
      return Promise.resolve({
        ok: false,
        status: 401,
        data: { message: "密码错误，请重试。" }
      });
    }

    const expiresAt = Date.now() + AUTH_TTL_MS;
    writeAuthPayload({ expiresAt });
    return Promise.resolve({
      ok: true,
      status: 200,
      data: {
        authenticated: true,
        expiresAt,
        remainingMs: AUTH_TTL_MS
      }
    });
  }

  function logoutFromServer() {
    clearAuthPayload();
    return Promise.resolve({ ok: true });
  }

  function requireDashboardAuth(redirectTo) {
    hasDashboardAccess().then((authenticated) => {
      if (!authenticated) {
        window.location.replace(redirectTo || "./index.html");
      }
    });
  }

  function bindLogoutButton(buttonId, redirectTo) {
    const button = document.getElementById(buttonId);
    if (!button) {
      return;
    }
    button.addEventListener("click", async function onLogout() {
      await logoutFromServer();
      window.location.replace(redirectTo || "./index.html");
    });
  }

  window.getAuthStatus = function getAuthStatusPublic() {
    return Promise.resolve(getAuthStatus());
  };
  window.hasDashboardAccess = hasDashboardAccess;
  window.loginWithPassword = loginWithPassword;
  window.logoutFromServer = logoutFromServer;
  window.requireDashboardAuth = requireDashboardAuth;
  window.bindLogoutButton = bindLogoutButton;
})();

