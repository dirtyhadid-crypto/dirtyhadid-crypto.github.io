(function authClientBootstrap() {
  async function getAuthStatus() {
    try {
      const response = await fetch("/api/auth/status", {
        method: "GET",
        headers: {
          Accept: "application/json"
        }
      });
      if (!response.ok) {
        return { authenticated: false };
      }
      return await response.json();
    } catch (_) {
      return { authenticated: false };
    }
  }

  async function hasDashboardAccess() {
    const data = await getAuthStatus();
    return Boolean(data.authenticated);
  }

  async function loginWithPassword(password) {
    try {
      const response = await fetch("/api/auth/login", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Accept: "application/json"
        },
        body: JSON.stringify({ password })
      });
      const data = await response.json().catch(() => ({}));
      return {
        ok: response.ok,
        status: response.status,
        data
      };
    } catch (_) {
      return {
        ok: false,
        status: 0,
        data: { message: "网络连接失败，请稍后重试。" }
      };
    }
  }

  async function logoutFromServer() {
    try {
      await fetch("/api/auth/logout", {
        method: "POST",
        headers: {
          Accept: "application/json"
        }
      });
    } catch (_) {
      // Ignore network errors and still try local redirect.
    }
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

  window.getAuthStatus = getAuthStatus;
  window.hasDashboardAccess = hasDashboardAccess;
  window.loginWithPassword = loginWithPassword;
  window.logoutFromServer = logoutFromServer;
  window.requireDashboardAuth = requireDashboardAuth;
  window.bindLogoutButton = bindLogoutButton;
})();
