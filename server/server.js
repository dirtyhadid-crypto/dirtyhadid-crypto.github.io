const crypto = require("crypto");
const express = require("express");
const fs = require("fs");
const path = require("path");

const app = express();

const PORT = Number(process.env.PORT || 3000);
const PASSWORD = String(process.env.DASHBOARD_PASSWORD || "123456");
const AUTH_TTL_MS = 8 * 60 * 60 * 1000;
const PUBLIC_DIR = path.join(__dirname, "..", "public");
const STORE_FILE = path.join(__dirname, "..", "data", "ip-auth-store.json");
const TRUST_PROXY = process.env.TRUST_PROXY === "1";

app.set("trust proxy", TRUST_PROXY);
app.use(express.json({ limit: "16kb" }));

const PROTECTED_PAGES = new Set([
  "/overview.html",
  "/amazon.html",
  "/influencer.html",
  "/social.html",
  "/media.html",
  "/dashboard.html"
]);

function normalizeIp(ipValue) {
  const ip = String(ipValue || "").trim();
  if (!ip) {
    return "unknown";
  }
  if (ip.startsWith("::ffff:")) {
    return ip.slice(7);
  }
  return ip;
}

function getClientIp(req) {
  return normalizeIp(req.ip || req.socket?.remoteAddress);
}

function readStore() {
  try {
    const content = fs.readFileSync(STORE_FILE, "utf8");
    const parsed = JSON.parse(content);
    if (parsed && typeof parsed === "object" && !Array.isArray(parsed)) {
      return parsed;
    }
  } catch (_) {
    // Ignore invalid or missing store file.
  }
  return {};
}

function ensureStoreDir() {
  fs.mkdirSync(path.dirname(STORE_FILE), { recursive: true });
}

function writeStore(store) {
  ensureStoreDir();
  fs.writeFileSync(STORE_FILE, JSON.stringify(store, null, 2), "utf8");
}

const ipAuthStore = readStore();

function pruneExpiredEntries() {
  const now = Date.now();
  let changed = false;

  Object.keys(ipAuthStore).forEach((ip) => {
    const expiresAt = Number(ipAuthStore[ip]);
    if (!Number.isFinite(expiresAt) || expiresAt <= now) {
      delete ipAuthStore[ip];
      changed = true;
    }
  });

  if (changed) {
    writeStore(ipAuthStore);
  }
}

function getAuthInfoByIp(ip) {
  const expiresAt = Number(ipAuthStore[ip]);
  const now = Date.now();
  if (!Number.isFinite(expiresAt) || expiresAt <= now) {
    if (ipAuthStore[ip]) {
      delete ipAuthStore[ip];
      writeStore(ipAuthStore);
    }
    return { authenticated: false, expiresAt: null, remainingMs: 0 };
  }
  return {
    authenticated: true,
    expiresAt,
    remainingMs: expiresAt - now
  };
}

function grantIpAccess(ip) {
  const expiresAt = Date.now() + AUTH_TTL_MS;
  ipAuthStore[ip] = expiresAt;
  writeStore(ipAuthStore);
  return expiresAt;
}

function clearIpAccess(ip) {
  if (!ipAuthStore[ip]) {
    return;
  }
  delete ipAuthStore[ip];
  writeStore(ipAuthStore);
}

function isPasswordValid(inputPassword) {
  const expected = Buffer.from(PASSWORD, "utf8");
  const actual = Buffer.from(String(inputPassword || ""), "utf8");
  if (expected.length !== actual.length) {
    return false;
  }
  return crypto.timingSafeEqual(expected, actual);
}

function authPageGuard(req, res, next) {
  if (req.method !== "GET") {
    next();
    return;
  }

  const pagePath = req.path === "/" ? "/index.html" : req.path;
  if (!PROTECTED_PAGES.has(pagePath)) {
    next();
    return;
  }

  const ip = getClientIp(req);
  const authInfo = getAuthInfoByIp(ip);
  if (!authInfo.authenticated) {
    res.redirect("/index.html");
    return;
  }

  next();
}

app.get("/api/auth/status", (req, res) => {
  const ip = getClientIp(req);
  const authInfo = getAuthInfoByIp(ip);
  res.json(authInfo);
});

app.post("/api/auth/login", (req, res) => {
  const inputPassword = String(req.body?.password || "");
  if (!inputPassword) {
    res.status(400).json({ message: "密码不能为空。" });
    return;
  }

  if (!isPasswordValid(inputPassword)) {
    res.status(401).json({ message: "密码错误，请重试。" });
    return;
  }

  const ip = getClientIp(req);
  const expiresAt = grantIpAccess(ip);
  res.json({
    authenticated: true,
    expiresAt,
    remainingMs: AUTH_TTL_MS
  });
});

app.post("/api/auth/logout", (req, res) => {
  const ip = getClientIp(req);
  clearIpAccess(ip);
  res.json({ ok: true });
});

app.use(authPageGuard);
app.use(express.static(PUBLIC_DIR, { index: "index.html" }));

app.get("*", (req, res) => {
  res.status(404).send("404 Not Found");
});

pruneExpiredEntries();
setInterval(pruneExpiredEntries, 60 * 1000);

app.listen(PORT, () => {
  const proxyMode = TRUST_PROXY ? "on" : "off";
  console.log(`Dashboard server running at http://localhost:${PORT}`);
  console.log(`IP auth ttl: ${AUTH_TTL_MS / (60 * 60 * 1000)} hours, trust proxy: ${proxyMode}`);
});
