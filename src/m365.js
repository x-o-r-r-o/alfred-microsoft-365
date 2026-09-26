#!/usr/bin/osascript -l JavaScript
// Microsoft 365 for Alfred: Teams, OneNote and Outlook through Microsoft Graph.
// Usage: osascript -l JavaScript m365.js <command> [query]
//   teams | onenote | outlook | account <query>   Script Filters (print Alfred JSON)
//   act <arg>                                     run the action named by $m365_action
//   poll | refresh <name>                         background jobs (spawned by the workflow)
//
// Security: access and refresh tokens live only in the macOS Keychain (or a test directory).
// They are sent to curl through its stdin config (-K -), never as command-line arguments,
// and are never written to the cache, to logs or to Alfred's output.
ObjC.import("Foundation");
ObjC.import("AppKit");
ObjC.import("Security");

const ENV = $.NSProcessInfo.processInfo.environment;
function env(name, fallback) {
  const v = ENV.objectForKey(name);
  return v.isNil() ? fallback : v.js;
}

const HOME = env("HOME", "/tmp");
const BUNDLE = env("alfred_workflow_bundleid", "io.github.x-o-r-r-o.microsoft-365");
const SETUP_URL = "https://github.com/x-o-r-r-o/alfred-microsoft-365#setup";
const trimSlash = (s) => String(s).replace(/\/+$/, "");
const LOGIN_BASE = trimSlash(env("M365_LOGIN_BASE", "https://login.microsoftonline.com"));
const GRAPH_BASE = trimSlash(env("M365_GRAPH_BASE", "https://graph.microsoft.com/v1.0"));
const CURL = env("M365_TEST_CURL", "/usr/bin/curl");
// Test mode: as soon as any test override is set, every side effect must be overridden too.
// A forgotten override then fails loudly instead of reaching Microsoft, the Keychain, the
// clipboard, the browser or Notification Center.
const TEST_MODE = (ObjC.deepUnwrap(ENV.allKeys) || []).some((k) => /^M365_(TEST_|LOGIN_BASE$|GRAPH_BASE$)/.test(k));
function testOnly(name, allowEmpty = false) {
  const v = env(name, null);
  if (TEST_MODE && (v === null || (v === "" && !allowEmpty))) throw new M365Error("graph", `Test mode: ${name} is not set`);
  return v;
}
// Minimum delegated permissions, one per feature (see README "Setup").
const CLIENT_ID = env("client_id", "").trim();
const TENANT = env("tenant", "").trim() || "common";
// Personal Microsoft accounts can't use presence, and asking for it can fail their sign-in.
const PERSONAL_ONLY = TENANT.toLowerCase() === "consumers";
const SCOPES = [
  "offline_access", // refresh token
  "User.Read", // your name and email
  "Calendars.Read", // Teams meetings and today's agenda
  "Mail.Read", // mail search
  "Notes.ReadWrite", // search OneNote pages and create new ones
  PERSONAL_ONLY ? null : "Presence.ReadWrite", // read and set your Teams status
  "People.Read", // find people to chat with
].filter(Boolean).join(" ");
const TEAMS_OPEN = env("teams_open", "app");
const ONENOTE_OPEN = env("onenote_open", "app");
const ONENOTE_SECTION = env("onenote_section", "").trim();
// A required keyword can still arrive empty (or padded) from the Workflow Configuration.
const KW_ACCOUNT = env("keyword_account", "").trim() || "m365";
const KW_ONENOTE = env("keyword_onenote", "").trim() || "onenote";
let BACKGROUND = false; // true inside background jobs: allows longer Retry-After waits
// Background refreshes stop starting requests after this, so they always finish well within their
// lock's 300 s expiry (worst case: one token refresh of 40 s and one request of 20 s after it).
let DEADLINE = Infinity;

// ---------- small helpers ----------

class M365Error extends Error {
  // kind: config | auth | network | throttle | consent | graph
  constructor(kind, message, extra = {}) {
    super(message);
    this.kind = kind;
    Object.assign(this, extra);
  }
}

function nowMs() {
  return Date.now();
}
// Wall clock for calendars; the tests pin it with M365_TEST_NOW.
function clock() {
  const t = env("M365_TEST_NOW", "");
  return t ? new Date(t) : new Date();
}
function sleep(sec) {
  if (sec > 0) $.NSThread.sleepForTimeInterval(sec);
}
// Lone UTF-16 surrogates make Alfred reject the JSON: keep pairs, replace strays.
function fixSurrogates(s) {
  return String(s).replace(/[\uD800-\uDBFF][\uDC00-\uDFFF]|[\uD800-\uDFFF]/g, (m) => (m.length === 2 ? m : "�"));
}
// Display text: no control characters or bidi overrides (a mail subject can use them to disguise itself).
function clean(s) {
  return fixSurrogates(String(s == null ? "" : s))
    .replace(/[‪-‮⁦-⁩‎‏؜]/g, "")
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F-\u009F]/g, "");
}
function oneLine(s, max = 120) {
  const t = clean(s).replace(/\s+/g, " ").trim();
  if (t.length <= max) return t;
  const cps = Array.from(t); // never split an emoji (surrogate pair) in half
  return cps.length > max ? cps.slice(0, max - 1).join("") + "…" : t;
}
// Case-, accent- and width-insensitive text for matching.
function norm(s) {
  return String(s || "").normalize("NFKD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
}
function hash(s) {
  let a = 5381, b = 52711;
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    a = (a * 33) ^ c;
    b = (b * 31) ^ c;
  }
  return (a >>> 0).toString(16) + (b >>> 0).toString(16);
}
function plural(n, word) {
  return `${n} ${word}${n === 1 ? "" : "s"}`;
}
function isHttpUrl(u) {
  return typeof u === "string" && /^https:\/\/[^\s]+$/i.test(u);
}
function formEncode(obj) {
  return Object.keys(obj).map((k) => `${encodeURIComponent(k)}=${encodeURIComponent(obj[k])}`).join("&");
}
function randomId() {
  return $.NSUUID.UUID.UUIDString.js.toLowerCase();
}

// ---------- files ----------

const FM = $.NSFileManager.defaultManager;
function mkdirp(p) {
  FM.createDirectoryAtPathWithIntermediateDirectoriesAttributesError(p, true, $({ NSFilePosixPermissions: 448 }), null);
  return p;
}
function exists(p) {
  return FM.fileExistsAtPath(p);
}
function readText(p) {
  const s = $.NSString.stringWithContentsOfFileEncodingError(p, $.NSUTF8StringEncoding, null);
  return s.isNil() ? null : s.js;
}
function writeText(p, text) {
  mkdirp(p.replace(/\/[^/]+$/, ""));
  const ok = $(text).writeToFileAtomicallyEncodingError(p, true, $.NSUTF8StringEncoding, null);
  if (ok) FM.setAttributesOfItemAtPathError($({ NSFilePosixPermissions: 384 }), p, null);
  return ok;
}
function readJSON(p) {
  const t = readText(p);
  if (t === null) return null;
  try {
    return JSON.parse(t);
  } catch (e) {
    return null;
  }
}
function writeJSON(p, v) {
  writeText(p, JSON.stringify(v));
}
function removePath(p) {
  if (exists(p)) FM.removeItemAtPathError(p, null);
}
function mtimeMs(p) {
  const a = FM.attributesOfItemAtPathError(p, null);
  if (a.isNil()) return 0;
  return a.fileModificationDate.timeIntervalSince1970 * 1000;
}

function cacheRoot() {
  return mkdirp(env("alfred_workflow_cache", `${HOME}/Library/Caches/com.runningwithcrayons.Alfred/Workflow Data/${BUNDLE}`));
}
function dataRoot() {
  return mkdirp(env("alfred_workflow_data", `${HOME}/Library/Application Support/Alfred/Workflow Data/${BUNDLE}`));
}
function account() {
  const a = readJSON(`${dataRoot()}/account.json`);
  return a && a.client_id === CLIENT_ID && String(a.tenant).toLowerCase() === TENANT.toLowerCase() ? a : null;
}
// Everything fetched from Graph is cached per account.
function cacheDir() {
  const a = account();
  const key = a && a.id ? String(a.id).replace(/[^A-Za-z0-9-]/g, "").slice(0, 64) || "default" : "default";
  return mkdirp(`${cacheRoot()}/${key}`);
}

// Advisory lock: mkdir is atomic. The owner's PID is stored inside, so a lock left behind by a
// killed process (Alfred terminates a Script Filter when you keep typing) is ignored at once;
// staleSec is a backstop.
ObjC.bindFunction("kill", ["int", ["int", "int"]]);
function pidAlive(pid) {
  if (!(pid > 0)) return false;
  if ($.kill(pid, 0) !== 0) return false; // lock owners are always this user's own processes
  // PIDs are reused: a live PID only counts if it still runs this script.
  const cmd = processCommand(pid);
  return cmd === null || /m365\.js/.test(cmd);
}
// The command line of a process, "" if it's gone, null if ps couldn't run.
function processCommand(pid) {
  const r = runWithStdin("/bin/ps", ["-o", "command=", "-p", String(pid)], "");
  if (r.code === -1 || !r.data) return null;
  return decode(r.data).trim();
}
function lockDir(name) {
  return `${cacheRoot()}/locks/${name}.lock`;
}
function lockStale(dir, staleSec) {
  if (nowMs() - mtimeMs(dir) > staleSec * 1000) return true;
  const pid = parseInt(readText(`${dir}/pid`) || "", 10);
  if (isNaN(pid)) return nowMs() - mtimeMs(dir) > 2000; // just created; the pid is written next
  return !pidAlive(pid);
}
function tryLock(name, staleSec) {
  mkdirp(`${cacheRoot()}/locks`);
  const dir = lockDir(name);
  const take = () => {
    if (!FM.createDirectoryAtPathWithIntermediateDirectoriesAttributesError(dir, false, $(), null)) return false;
    writeText(`${dir}/pid`, String($.NSProcessInfo.processInfo.processIdentifier));
    return true;
  };
  if (take()) return dir;
  if (exists(dir) && lockStale(dir, staleSec)) {
    removePath(dir);
    if (take()) return dir;
  }
  return null;
}
function lockHeld(name, staleSec) {
  const dir = lockDir(name);
  return exists(dir) && !lockStale(dir, staleSec);
}
function withLock(name, waitSec, staleSec, fn) {
  let lock = tryLock(name, staleSec);
  const until = nowMs() + waitSec * 1000;
  while (!lock && nowMs() < until) {
    sleep(0.1);
    lock = tryLock(name, staleSec);
  }
  try {
    return fn();
  } finally {
    if (lock) removePath(lock);
  }
}

// ---------- keychain ----------
// Security framework through the ObjC bridge: secrets never appear in a process's arguments.
// The dictionary keys are the string values of kSecClass, kSecAttrService and so on.

function kcQuery(acct) {
  const d = $.NSMutableDictionary.alloc.init;
  d.setObjectForKey($("genp"), $("class"));
  d.setObjectForKey($(BUNDLE), $("svce"));
  d.setObjectForKey($(acct), $("acct"));
  return d;
}
// A locked Keychain (or a cancelled unlock prompt) isn't "signed out": say so instead of
// offering a sign-in whose tokens couldn't be saved either.
function keychainError(status) {
  return new M365Error("keychain", `Couldn't read your sign-in from the Keychain (error ${status}). Unlock the login keychain and try again.`, { status });
}
function testSecretPath(acct) {
  const dir = testOnly("M365_TEST_KEYCHAIN_DIR") || "";
  return dir ? `${dir}/${acct.replace(/[^A-Za-z0-9._-]/g, "_")}.secret` : null;
}
function kcGet(acct) {
  const t = testSecretPath(acct);
  if (t) {
    const locked = Number(env("M365_TEST_KEYCHAIN_STATUS", "0"));
    if (locked) throw keychainError(locked);
    return readText(t);
  }
  const q = kcQuery(acct);
  q.setObjectForKey($.NSNumber.numberWithBool(true), $("r_Data"));
  q.setObjectForKey($("m_LimitOne"), $("m_Limit"));
  const r = Ref();
  const status = $.SecItemCopyMatching(q, r);
  if (status === -25300) return null; // errSecItemNotFound
  if (status !== 0) throw keychainError(status);
  const s = $.NSString.alloc.initWithDataEncoding(ObjC.castRefToObject(r[0]), $.NSUTF8StringEncoding);
  return s.isNil() ? null : s.js;
}
function kcSet(acct, value) {
  const t = testSecretPath(acct);
  if (t) {
    if (!writeText(t, value)) throw new M365Error("auth", "Couldn't save the sign-in to the Keychain");
    return;
  }
  const data = $(value).dataUsingEncoding($.NSUTF8StringEncoding);
  const q = kcQuery(acct);
  const upd = $.NSMutableDictionary.alloc.init;
  upd.setObjectForKey(data, $("v_Data"));
  let status = $.SecItemUpdate(q, upd);
  if (status === -25300) {
    q.setObjectForKey(data, $("v_Data"));
    q.setObjectForKey($("Microsoft 365 (Alfred) sign-in"), $("labl"));
    status = $.SecItemAdd(q, null);
  }
  if (status !== 0) throw new M365Error("auth", `Couldn't save the sign-in to the Keychain (error ${status})`);
}
function kcDel(acct) {
  const t = testSecretPath(acct);
  if (t) return removePath(t);
  $.SecItemDelete(kcQuery(acct));
}

// ---------- configuration ----------

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
function configProblem() {
  if (!CLIENT_ID) return "Set the Application (client) ID in the Workflow’s Configuration";
  if (!GUID.test(CLIENT_ID)) return "The Application (client) ID must look like 00000000-0000-0000-0000-000000000000";
  if (!/^[A-Za-z0-9][A-Za-z0-9.-]{0,99}$/.test(TENANT)) return "The tenant must be common, organizations, consumers, a domain or a Directory (tenant) ID";
  return null;
}
function requireConfig() {
  const p = configProblem();
  if (p) throw new M365Error("config", p);
}
const kcAccount = () => `oauth:${CLIENT_ID}@${TENANT.toLowerCase()}`;

// ---------- HTTP (curl with its configuration on stdin) ----------

function cfgQuote(s) {
  return '"' + String(s).replace(/\\/g, "\\\\").replace(/"/g, '\\"').replace(/\n/g, "\\n").replace(/\r/g, "\\r").replace(/\t/g, "\\t") + '"';
}

function runWithStdin(path, args, input) {
  const task = $.NSTask.alloc.init;
  task.executableURL = $.NSURL.fileURLWithPath(path);
  task.arguments = args;
  const inP = $.NSPipe.pipe, outP = $.NSPipe.pipe;
  task.standardInput = inP;
  task.standardOutput = outP;
  task.standardError = $.NSFileHandle.fileHandleWithNullDevice;
  if (!task.launchAndReturnError(null)) return { code: -1, data: null };
  inP.fileHandleForWriting.writeData($(input).dataUsingEncoding($.NSUTF8StringEncoding));
  inP.fileHandleForWriting.closeFile;
  const data = outP.fileHandleForReading.readDataToEndOfFile;
  task.waitUntilExit;
  return { code: task.terminationStatus, data };
}

function decode(data) {
  let s = $.NSString.alloc.initWithDataEncoding(data, $.NSUTF8StringEncoding);
  if (s.isNil()) s = $.NSString.alloc.initWithDataEncoding(data, $.NSISOLatin1StringEncoding);
  return s.isNil() ? "" : s.js;
}

// Returns {status, headers, text, json} or {net: curlExitCode}.
function http(method, url, { headers = {}, body = null, timeout = 20 } = {}) {
  if (!/^https?:\/\//i.test(url)) throw new M365Error("graph", `Refusing to request ${url}`);
  if (TEST_MODE) {
    testOnly("M365_LOGIN_BASE");
    testOnly("M365_GRAPH_BASE");
    // Only the local mock: never Microsoft, whatever the overrides say. Loopback literals only:
    // "localhost" can be remapped in /etc/hosts.
    const local = /^https?:\/\/(127(?:\.\d{1,3}){3}|\[::1\])(:\d+)?\//i.test(url);
    if (!local || (!sameOrigin(url, LOGIN_BASE) && !sameOrigin(url, GRAPH_BASE))) throw new M365Error("graph", `Test mode: refusing to request ${url}`);
  }
  const lines = ["silent", "include", "globoff", `url = ${cfgQuote(url)}`, `request = ${cfgQuote(method)}`,
    `max-time = ${timeout}`, "connect-timeout = 8", 'header = "Expect:"'];
  // Test mode: a proxy from the environment (http_proxy, ALL_PROXY…) must not see the mock's traffic.
  if (TEST_MODE) lines.push('noproxy = "*"');
  for (const [k, v] of Object.entries(headers)) lines.push(`header = ${cfgQuote(`${k}: ${String(v).replace(/[\r\n]+/g, " ")}`)}`);
  if (body !== null) {
    // data-binary would read a file for a leading "@"; our bodies are JSON, forms or HTML.
    if (String(body).startsWith("@")) throw new M365Error("graph", "Invalid request body");
    lines.push(`data-binary = ${cfgQuote(body)}`);
  }
  const r = runWithStdin(CURL, ["-K", "-"], lines.join("\n") + "\n");
  if (r.code !== 0) return { net: r.code };
  let text = decode(r.data);
  let status = 0, head = "";
  // Skip interim (1xx) responses; the final header block precedes the body.
  while (/^HTTP\/[\d.]+ \d{3}/.test(text)) {
    let idx = text.indexOf("\r\n\r\n"), sep = 4;
    if (idx < 0) {
      idx = text.indexOf("\n\n");
      sep = 2;
    }
    head = idx < 0 ? text : text.slice(0, idx);
    text = idx < 0 ? "" : text.slice(idx + sep);
    status = parseInt(head.match(/^HTTP\/[\d.]+ (\d{3})/)[1], 10);
    // More header blocks follow after 1xx responses and a proxy's "200 Connection established".
  }
  const hdrs = Object.create(null);
  for (const line of head.split(/\r?\n/).slice(1)) {
    const m = line.match(/^([^:]+):\s*(.*)$/);
    if (m) hdrs[m[1].toLowerCase()] = m[2].trim();
  }
  let json = null;
  try {
    json = text ? JSON.parse(text) : {};
  } catch (e) {
    json = null;
  }
  return { status, headers: hdrs, text, json };
}

function netError(code) {
  // Only connection problems mean "offline"; anything else (e.g. a malformed URL) is a bug to report.
  if ([5, 6, 7, 28, 35, 52, 55, 56].includes(code)) return new M365Error("network", "Can't reach Microsoft. Check your internet connection.", { code });
  return new M365Error("graph", `Request failed (curl error ${code})`, { code });
}

// A 2xx that isn't JSON comes from a captive portal or a proxy, not from Microsoft.
function notJSON(r) {
  return r.json === null;
}
function unexpectedResponse() {
  return new M365Error("network", "Unexpected answer from the network (captive portal or proxy?). Check your connection.");
}

function retryAfterSec(r) {
  const v = (r.headers || {})["retry-after"];
  if (!v) return 10;
  if (/^\d+$/.test(v)) return parseInt(v, 10);
  const t = Date.parse(v);
  return isNaN(t) ? 10 : Math.max(1, Math.ceil((t - nowMs()) / 1000));
}

// ---------- Microsoft identity platform ----------

// Clear messages for the AADSTS errors people actually hit with a self-registered app.
const AADSTS = {
  65001: ["consent", "The app doesn't have your consent for these permissions. Sign in again to consent, or ask your IT admin to grant admin consent."],
  65004: ["auth", "You declined to give the app permission. Sign in again and accept the permissions."],
  90094: ["consent", "Your organization requires an admin to approve this app. Ask your IT admin to grant admin consent for it."],
  90095: ["consent", "Your organization requires an admin to approve this app. Ask your IT admin to grant admin consent for it."],
  700016: ["config", "No app with this Application (client) ID in this tenant. Check the client ID and tenant in the Workflow’s Configuration."],
  7000218: ["config", "Turn on “Allow public client flows” (Authentication page of your app registration), then sign in again."],
  50194: ["config", "Your app registration is single-tenant: set its Directory (tenant) ID instead of “common” in the Workflow’s Configuration."],
  9002346: ["config", "Your app registration only allows personal Microsoft accounts: set the tenant to “consumers”."],
  9002331: ["config", "Your app registration only allows personal Microsoft accounts: set the tenant to “consumers”."],
  50059: ["config", "Microsoft couldn't tell which tenant to use. Set your Directory (tenant) ID in the Workflow’s Configuration."],
  90002: ["config", "Tenant not found. Check the tenant in the Workflow’s Configuration."],
  900023: ["config", "The tenant isn't valid. Check the tenant in the Workflow’s Configuration."],
  50020: ["auth", "This account doesn't belong to the app's tenant. Check the tenant or the app's supported account types."],
  50105: ["consent", "Your admin must assign you to this app before you can use it."],
  53003: ["auth", "Blocked by your organization's Conditional Access policy (it may block device code sign-in). Ask your IT admin."],
  70011: ["config", "Microsoft rejected the requested permissions. With a personal Microsoft account, set the tenant to “consumers”."],
  700082: ["auth", "Your sign-in expired after a long period of inactivity. Sign in again."],
  70008: ["auth", "Your sign-in expired. Sign in again."],
  50173: ["auth", "Your sign-in is no longer valid (password changed or session revoked). Sign in again."],
  50076: ["auth", "Your organization requires multi-factor authentication again. Sign in again."],
  50079: ["auth", "Your organization requires multi-factor authentication again. Sign in again."],
  50078: ["auth", "Your organization requires multi-factor authentication again. Sign in again."],
  70000: ["auth", "Your sign-in is no longer valid. Sign in again."],
  700084: ["auth", "Your sign-in expired. Sign in again."],
};

function aadCode(j) {
  if (!j) return null;
  if (Array.isArray(j.error_codes) && j.error_codes.length) return Number(j.error_codes[0]);
  const m = String(j.error_description || "").match(/AADSTS(\d+)/);
  return m ? Number(m[1]) : null;
}
function aadError(j, status) {
  const code = aadCode(j);
  if (code && AADSTS[code]) return new M365Error(AADSTS[code][0], AADSTS[code][1], { aadsts: code, error: j.error });
  let desc = String((j && j.error_description) || "").split(/\r?\n/)[0].replace(/^AADSTS\d+:\s*/, "")
    .replace(/\s*(Trace ID|Correlation ID|Timestamp):[\s\S]*$/, "");
  if (!desc) desc = (j && j.error) || `Sign-in failed (HTTP ${status})`;
  const kind = j && ["invalid_grant", "interaction_required", "login_required", "consent_required"].includes(j.error) ? "auth" : "config";
  return new M365Error(kind, oneLine(desc, 200), { aadsts: code, error: j && j.error });
}

function tokenPost(params) {
  return http("POST", `${LOGIN_BASE}/${TENANT}/oauth2/v2.0/token`, {
    headers: { "Content-Type": "application/x-www-form-urlencoded", Accept: "application/json" },
    body: formEncode(params),
  });
}

function loadTokens() {
  const s = kcGet(kcAccount());
  if (!s) return null;
  try {
    const t = JSON.parse(s);
    return t && t.refresh_token ? t : null;
  } catch (e) {
    return null;
  }
}
function saveTokenResponse(j, oldRefresh) {
  const life = Math.max(60, Number(j.expires_in) || 3600);
  const t = {
    access_token: j.access_token,
    refresh_token: j.refresh_token || oldRefresh,
    // Refresh a little early so a token never expires mid-request.
    expires_at: nowMs() + Math.max(30, life - 120) * 1000,
    scope: j.scope || "",
  };
  kcSet(kcAccount(), JSON.stringify(t));
  return t;
}

// staleAccess: the access token Graph just rejected (force a refresh unless another process already did).
function refreshTokens(staleAccess) {
  return withLock("token", 20, 30, () => {
    const t = loadTokens();
    if (!t) throw new M365Error("auth", "Not signed in");
    const usable = t.access_token && t.expires_at > nowMs();
    if (usable && (!staleAccess || t.access_token !== staleAccess)) return t;
    const r = tokenPost({ grant_type: "refresh_token", client_id: CLIENT_ID, refresh_token: t.refresh_token, scope: SCOPES });
    if (r.net) throw netError(r.net);
    if (notJSON(r)) throw unexpectedResponse();
    if (r.status === 200 && r.json.access_token) return saveTokenResponse(r.json, t.refresh_token);
    if (r.status === 429 || r.status >= 500) throw new M365Error("throttle", "Microsoft sign-in is busy. Try again shortly.", { retryAfter: retryAfterSec(r) });
    const e = aadError(r.json, r.status);
    if (e.kind === "auth" || e.kind === "consent") kcDel(kcAccount()); // the refresh token is dead
    if (e.kind === "config") e.kind = "auth";
    throw e;
  });
}

function accessToken() {
  const t = loadTokens();
  if (!t) throw new M365Error("auth", "Not signed in");
  if (t.access_token && t.expires_at > nowMs()) return t.access_token;
  return refreshTokens(null).access_token;
}

// ---------- Microsoft Graph ----------

function sameOrigin(url, base) {
  const o = (u) => (u.match(/^(https?:\/\/[^/?#]+)/i) || [])[1];
  const a = o(url), b = o(base);
  return !!a && !!b && a.toLowerCase() === b.toLowerCase();
}
function throttleFile() {
  return `${cacheRoot()}/throttled-until`;
}

function graph(method, path, opts = {}) {
  const url = /^https?:/i.test(path) ? path : GRAPH_BASE + path;
  // Never send the token anywhere but Graph (e.g. a crafted @odata.nextLink).
  if (!sameOrigin(url, GRAPH_BASE)) throw new M365Error("graph", "Refusing to follow a link outside Microsoft Graph");
  const until = Number(readText(throttleFile()) || 0);
  if (until > nowMs()) throw new M365Error("throttle", "Microsoft Graph asked us to slow down.", { retryAfter: Math.ceil((until - nowMs()) / 1000) });
  let token = accessToken();
  let refreshed = false, retries = 0;
  for (;;) {
    if (nowMs() > DEADLINE) throw new M365Error("timeout", "Microsoft 365 took too long to answer. Try again later.");
    const headers = Object.assign({ Authorization: `Bearer ${token}`, Accept: "application/json" }, opts.headers || {});
    let body = null;
    if (opts.json !== undefined) {
      headers["Content-Type"] = "application/json";
      body = JSON.stringify(opts.json);
    } else if (opts.body !== undefined) {
      body = opts.body;
    }
    const r = http(method, url, { headers, body, timeout: opts.timeout || 20 });
    if (r.net) throw netError(r.net);
    if (r.status === 401 && !refreshed) {
      refreshed = true;
      token = refreshTokens(token).access_token;
      continue;
    }
    if (r.status === 401) {
      // Keep the refresh token: this can be a claims challenge or a transient failure.
      throw new M365Error("auth", "Microsoft rejected your sign-in. Sign in again.");
    }
    if (r.status === 429 || r.status === 503 || r.status === 504) {
      const wait = retryAfterSec(r);
      if (retries < 2 && wait <= (BACKGROUND ? 30 : 2) && nowMs() + wait * 1000 < DEADLINE) {
        retries++;
        sleep(wait);
        continue;
      }
      writeText(throttleFile(), String(nowMs() + Math.min(wait, 300) * 1000));
      throw new M365Error("throttle", "Microsoft Graph asked us to slow down.", { retryAfter: wait });
    }
    if (r.status >= 400) throw graphError(r);
    if (notJSON(r)) throw unexpectedResponse();
    return r.json;
  }
}

function graphError(r) {
  const e = (r.json && r.json.error) || {};
  const code = String(e.code || "");
  const msg = oneLine(e.message || `HTTP ${r.status}`, 200);
  if (code === "10008") {
    return new M365Error("graph", "OneNote can't list your pages: a OneDrive library holds more than 5,000 OneNote items (error 10008)", { status: r.status, code });
  }
  if (r.status === 403) {
    return new M365Error("consent", `Permission denied: ${msg}`, { status: 403, code });
  }
  return new M365Error("graph", `Microsoft Graph error ${r.status}: ${msg}`, { status: r.status, code });
}

// Follow @odata.nextLink until done (or maxPages).
function graphAll(path, { maxPages = 20, headers } = {}) {
  let url = path, out = [], pages = 0;
  while (url && pages < maxPages) {
    const j = graph("GET", url, { headers });
    out = out.concat(j.value || []);
    url = j["@odata.nextLink"] || null;
    pages++;
  }
  return out;
}

// ---------- system: clipboard, open, notify, background jobs ----------

const CLIPBOARD_MAX = Number(env("M365_TEST_CLIPBOARD_MAX", "")) || 1000000;
function clipboardText() {
  let t = testOnly("M365_TEST_CLIPBOARD", true);
  if (t === null) {
    const str = $.NSPasteboard.generalPasteboard.stringForType($.NSPasteboardTypeString);
    t = str.isNil() ? "" : str.js;
  }
  if (t.length > CLIPBOARD_MAX) throw new M365Error("graph", `The clipboard is too large for a OneNote page (over ${plural(CLIPBOARD_MAX, "character")})`);
  return t;
}
function copyText(text, transient = false) {
  const f = testOnly("M365_TEST_CLIPBOARD_OUT");
  if (f) return writeText(f, text);
  const pb = $.NSPasteboard.generalPasteboard;
  pb.clearContents;
  pb.setStringForType($(text), $.NSPasteboardTypeString);
  // Tell clipboard managers (including Alfred's Clipboard History) not to keep one-time codes.
  if (transient) pb.setStringForType($(""), $("org.nspasteboard.TransientType"));
}
function testRecord(varName, line) {
  const f = env(varName, "");
  if (!f) return false;
  const prev = readText(f) || "";
  writeText(f, prev + line + "\n");
  return true;
}
function appCanOpen(url) {
  const fake = TEST_MODE ? env("M365_TEST_APPS", "") : null;
  if (fake !== null) return fake.split(",").includes(url.split(":")[0]);
  const u = $.NSURL.URLWithString(url);
  return !u.isNil() && !$.NSWorkspace.sharedWorkspace.URLForApplicationToOpenURL(u).isNil();
}
function openURL(url) {
  if (!/^(https:|msteams:|onenote:)/i.test(url) && !(env("M365_TEST_OPEN_FILE", "") && /^http:/.test(url))) {
    throw new M365Error("graph", "Refusing to open a link that isn't https");
  }
  testOnly("M365_TEST_OPEN_FILE");
  if (testRecord("M365_TEST_OPEN_FILE", url)) return;
  const u = $.NSURL.URLWithString(url);
  if (u.isNil() || !$.NSWorkspace.sharedWorkspace.openURL(u)) throw new M365Error("graph", "Couldn't open the link");
}
function openPreferApp(appUrl, webUrl, preferApp) {
  if (preferApp && appUrl && appCanOpen(appUrl)) return openURL(appUrl);
  if (webUrl) return openURL(webUrl);
  if (appUrl) return openURL(appUrl);
  throw new M365Error("graph", "No link to open");
}
// Background jobs can't use Alfred's Notification object directly: ask Alfred to run
// this workflow's "notify" External Trigger, which is connected to it.
function notify(text) {
  if (testRecord("M365_TEST_NOTIFY_FILE", text) || TEST_MODE) return;
  try {
    Application("com.runningwithcrayons.Alfred").runTrigger("notify", { inWorkflow: BUNDLE, withArgument: text });
  } catch (e) {
    const app = Application.currentApplication();
    app.includeStandardAdditions = true;
    app.displayNotification(text, { withTitle: "Microsoft 365" });
  }
}
function scriptPath() {
  const args = ObjC.deepUnwrap($.NSProcessInfo.processInfo.arguments) || [];
  let p = args.find((a) => /m365\.js$/.test(a)) || "./m365.js";
  if (!p.startsWith("/")) p = `${FM.currentDirectoryPath.js}/${p.replace(/^\.\//, "")}`;
  return p;
}
// Start `osascript m365.js <args>` detached (nohup, no pipes) so Alfred isn't kept waiting.
// Secrets for the job travel in its environment, never in its arguments.
function spawnDetached(args, extraEnv = {}) {
  const task = $.NSTask.alloc.init;
  task.executableURL = $.NSURL.fileURLWithPath("/bin/bash");
  // set -m puts the job in its own process group, so it survives Alfred killing the Script Filter's group.
  task.arguments = ["-c", 'set -m; /usr/bin/nohup /usr/bin/osascript -l JavaScript "$0" "$@" </dev/null >/dev/null 2>&1 &', scriptPath()].concat(args);
  const e = ObjC.deepUnwrap(ENV) || {};
  Object.assign(e, extraEnv);
  task.environment = $(e);
  task.standardInput = $.NSFileHandle.fileHandleWithNullDevice;
  task.standardOutput = $.NSFileHandle.fileHandleWithNullDevice;
  task.standardError = $.NSFileHandle.fileHandleWithNullDevice;
  if (task.launchAndReturnError(null)) task.waitUntilExit;
}

// ---------- cache with background refresh (stale-while-revalidate) ----------

function cacheFile(name) {
  return `${cacheDir()}/${name}.json`;
}
function errorFile(name) {
  return `${cacheDir()}/${name}.error.json`;
}
// Cache files can be corrupt, truncated or from an older version: accept only the shape we write.
function readCache(name) {
  const c = readJSON(cacheFile(name));
  if (!c || typeof c !== "object" || Array.isArray(c)) return null;
  if (typeof c.fetched_at !== "number" || !isFinite(c.fetched_at) || c.fetched_at > nowMs() + 60000) return null;
  const d = c.data;
  const ok = name === "presence" ? d && typeof d === "object" && !Array.isArray(d) : Array.isArray(d);
  return ok ? c : null;
}
function readCacheError(name) {
  const e = readJSON(errorFile(name));
  return e && typeof e === "object" && typeof e.at === "number" && typeof e.message === "string" ? e : null;
}
function errorToJSON(e) {
  return { kind: e.kind || "graph", message: e.message, retryAfter: e.retryAfter, status: e.status, at: nowMs() };
}
function errorFromJSON(j) {
  return new M365Error(j.kind, j.message, { retryAfter: j.retryAfter, status: j.status });
}
function storeCache(name, data) {
  writeJSON(cacheFile(name), { fetched_at: nowMs(), data });
  removePath(errorFile(name));
  if (/^(mail|people)-/.test(name)) pruneQueryCache();
}
// Search results are cached per query (one file per keystroke): drop the ones older than a day
// and keep at most the newest QUERY_CACHE_MAX.
const QUERY_CACHE_MAX = Number(env("M365_TEST_QUERY_CACHE_MAX", "")) || 200;
function pruneQueryCache() {
  const dir = cacheDir();
  const names = ObjC.deepUnwrap(FM.contentsOfDirectoryAtPathError(dir, null)) || [];
  const files = names.filter((n) => /^(mail|people)-[0-9a-f]+\.json$/.test(n))
    .map((n) => ({ p: `${dir}/${n}`, m: mtimeMs(`${dir}/${n}`) })).sort((a, b) => b.m - a.m);
  files.forEach((f, i) => {
    if (i >= QUERY_CACHE_MAX || nowMs() - f.m > 86400000) removePath(f.p);
  });
}

// Background job: refresh one cache entry, once at a time.
function refreshJob(name) {
  BACKGROUND = true;
  DEADLINE = nowMs() + (Number(env("M365_TEST_JOB_DEADLINE", "")) || 200) * 1000;
  const lock = tryLock(`refresh-${name}`, 300);
  if (!lock) return;
  try {
    storeCache(name, fetcherFor(name)());
  } catch (e) {
    writeJSON(errorFile(name), errorToJSON(e));
  } finally {
    removePath(lock);
  }
}
function startRefresh(name) {
  if (lockHeld(`refresh-${name}`, 300)) return;
  if (env("M365_TEST_BG_SYNC", "") === "1") return refreshJob(name);
  spawnDetached(["refresh", name]);
}

// Returns {data, age, rerun?, warning?, loading?}. Fresh data is returned as is; stale data is
// returned immediately while a background job refreshes it (Alfred reruns the filter).
// Without any data: fetch now, or (async) start a job and report loading.
function swr(name, ttlSec, { async = false } = {}) {
  const c = readCache(name);
  const age = c ? (nowMs() - c.fetched_at) / 1000 : Infinity;
  if (c && age < ttlSec && age >= 0) return { data: c.data, age };
  const err = readCacheError(name);
  const running = lockHeld(`refresh-${name}`, 300);
  const recentErr = err && nowMs() - err.at < 60000 && (!c || err.at > c.fetched_at);
  if (c) {
    if (running) return { data: c.data, age, rerun: 0.5 };
    if (recentErr) return { data: c.data, age, warning: errorFromJSON(err) };
    startRefresh(name);
    // The test mode refreshes synchronously: pick up its result straight away.
    const after = readCache(name);
    if (after && after.fetched_at !== c.fetched_at) return { data: after.data, age: 0 };
    const e2 = readCacheError(name);
    if (e2 && e2.at > c.fetched_at && !lockHeld(`refresh-${name}`, 300)) return { data: c.data, age, warning: errorFromJSON(e2) };
    return { data: c.data, age, rerun: 0.5 };
  }
  if (async) {
    if (recentErr && !running) throw errorFromJSON(err);
    if (!running) startRefresh(name);
    const after = readCache(name);
    if (after) return { data: after.data, age: 0 };
    const e2 = readCacheError(name);
    if (e2 && !lockHeld(`refresh-${name}`, 300)) throw errorFromJSON(e2);
    return { data: null, loading: true, rerun: 1 };
  }
  const data = fetcherFor(name)();
  storeCache(name, data);
  return { data, age: 0 };
}

function fetcherFor(name) {
  let m;
  if ((m = name.match(/^events-(\d{4})-(\d{2})-(\d{2})$/))) return () => fetchEvents(new Date(+m[1], +m[2] - 1, +m[3]));
  if (name === "presence") return fetchPresence;
  if (name === "onenote-index") return buildOneNoteIndex;
  if (name === "onenote-sections") return fetchSections;
  throw new M365Error("graph", `Unknown cache ${name}`);
}

// ---------- Alfred items ----------

function icon(name) {
  return { path: `icons/${name}.png` };
}
function info(title, subtitle, ic = "info", extra = {}) {
  return Object.assign({ title: oneLine(title), subtitle: oneLine(subtitle || "", 200), valid: false, icon: icon(ic) }, extra);
}
// An actionable row. action goes to the "act" script as $m365_action.
function row(title, subtitle, ic, action, arg, vars = {}, mods = {}, extra = {}) {
  const variables = Object.assign({ m365_action: action }, vars);
  const it = { title: oneLine(title), subtitle: oneLine(subtitle || "", 200), arg, valid: true, icon: icon(ic), variables, mods: {} };
  for (const [k, m] of Object.entries(mods)) {
    it.mods[k] = m.valid === false
      ? { valid: false, arg: "", subtitle: m.subtitle }
      : { valid: true, arg: m.arg !== undefined ? m.arg : arg, subtitle: oneLine(m.subtitle, 200), variables: Object.assign({ m365_action: m.action }, m.vars || {}) };
  }
  if (extra.text && extra.text.largetype !== undefined) extra.text.largetype = clean(extra.text.largetype);
  return Object.assign(it, extra);
}
function output(items, extra = {}) {
  return JSON.stringify(Object.assign({ skipknowledge: true, items }, extra), (k, v) => (typeof v === "string" ? fixSurrogates(v) : v));
}
function ago(ms) {
  const m = Math.round(ms / 60000);
  if (m < 1) return "just now";
  if (m < 60) return `${m} min ago`;
  const h = Math.round(m / 60);
  return h < 48 ? `${h} h ago` : `${Math.round(h / 24)} days ago`;
}
function signInRow(subtitle) {
  return row("Sign in to Microsoft 365", subtitle || "Opens microsoft.com/devicelogin and copies your sign-in code", "login", "login", "login");
}

// Turn any error into helpful rows.
function errorItems(e) {
  if (!(e instanceof M365Error)) return [info("Something went wrong", oneLine(String(e && e.message ? e.message : e), 200), "error")];
  switch (e.kind) {
    case "config":
      return [row(e.message, "↩ Open the setup guide", "error", "open", SETUP_URL)];
    case "auth":
      return pendingLogin() ? pendingRows() : [signInRow(e.message === "Not signed in" ? null : e.message)];
    case "consent":
      return [info(e.message, "Your organization may require admin consent for this app", "error"),
        row("Sign in again", "To consent to the workflow’s permissions", "login", "login", "login")];
    case "network":
      return [info("You're offline", e.message, "offline")];
    case "keychain":
      return [info("Your Keychain is locked", e.message, "error")];
    case "throttle":
      return [info("Microsoft 365 is busy", `Too many requests. Try again in ${plural(e.retryAfter || 10, "second")}.`, "error")];
    default:
      return [info(e.message, `Try again in a moment, or check your sign-in and permissions via the ${KW_ACCOUNT} keyword`, "error")];
  }
}
function warningRow(w, age) {
  if (w.kind === "auth") return signInRow(`${oneLine(w.message, 80)} (showing results from ${ago(age * 1000)})`);
  const base = w.kind === "network" ? "Offline" : w.kind === "throttle" ? "Microsoft 365 is busy" : oneLine(w.message, 80);
  return info(`${base}: showing results from ${ago(age * 1000)}`, "", w.kind === "network" ? "offline" : "error");
}

// ---------- sign-in (device code flow) ----------

function pendingFile() {
  return `${dataRoot()}/login-pending.json`;
}
function pendingLogin() {
  const p = readJSON(pendingFile());
  return p && p.client_id === CLIENT_ID && p.expires_at > nowMs() && typeof p.user_code === "string" && typeof p.verification_uri === "string" ? p : null;
}
function pendingRows() {
  const p = pendingLogin();
  const mins = Math.max(1, Math.round((p.expires_at - nowMs()) / 60000));
  return [
    row(`Enter code ${p.user_code} to sign in`, `Waiting for you at ${p.verification_uri.replace(/^https?:\/\/(www\.)?/, "")} · expires in ${mins} min · ↩ copy the code and open the page`,
      "login", "login-show", p.user_code),
    row("Cancel sign-in", "", "logout", "cancel-login", "cancel"),
  ];
}

function login() {
  requireConfig();
  const r = http("POST", `${LOGIN_BASE}/${TENANT}/oauth2/v2.0/devicecode`, {
    headers: { "Content-Type": "application/x-www-form-urlencoded", Accept: "application/json" },
    body: formEncode({ client_id: CLIENT_ID, scope: SCOPES }),
  });
  if (r.net) return netError(r.net).message;
  if (notJSON(r)) return unexpectedResponse().message;
  const j = r.json;
  if (r.status !== 200 || !j.device_code || !j.user_code) return aadError(j, r.status).message;
  const uri = j.verification_uri || "https://microsoft.com/devicelogin";
  const p = {
    id: randomId(),
    client_id: CLIENT_ID,
    user_code: String(j.user_code),
    verification_uri: uri,
    // The code lives 15 minutes; never poll longer than 30 whatever the server says.
    expires_at: nowMs() + Math.min(Math.max(Number(j.expires_in) || 900, 1), 1800) * 1000,
    interval: Math.min(Math.max(Number(j.interval) || 5, 1), 60),
  };
  writeJSON(pendingFile(), p); // no device_code on disk: it goes to the poller's environment
  copyText(p.user_code, true);
  spawnDetached(["poll"], { M365_DEVICE_CODE: j.device_code, M365_LOGIN_ID: p.id });
  try {
    openURL(uri);
  } catch (e) {
    return `Open ${uri} and enter the code ${p.user_code} (copied)`;
  }
  return `Code ${p.user_code} copied: paste it in the browser to sign in`;
}

// Background poller: one per sign-in attempt, stops on success, error, expiry,
// cancellation or when a newer attempt replaces it.
function poll() {
  BACKGROUND = true;
  const deviceCode = env("M365_DEVICE_CODE", "");
  const id = env("M365_LOGIN_ID", "");
  const scale = Number(env("M365_TEST_INTERVAL_SCALE", "1")) || 1;
  let p = readJSON(pendingFile());
  if (!deviceCode || !p || p.id !== id) return;
  let interval = p.interval;
  const hardStop = nowMs() + 1800 * 1000;
  const finish = (msg) => {
    const cur = readJSON(pendingFile());
    if (cur && cur.id === id) removePath(pendingFile());
    if (msg) notify(msg);
  };
  for (;;) {
    const left = (Math.min(p.expires_at, hardStop) - nowMs()) / 1000;
    sleep(Math.max(0, Math.min(interval * scale, left)));
    p = readJSON(pendingFile());
    if (!p || p.id !== id) return; // cancelled or superseded
    if (nowMs() >= p.expires_at || nowMs() >= hardStop) return finish(`The sign-in code expired. Use the ${KW_ACCOUNT} keyword to try again.`);
    const r = tokenPost({ grant_type: "urn:ietf:params:oauth:grant-type:device_code", client_id: CLIENT_ID, device_code: deviceCode });
    if (r.net || notJSON(r)) continue; // offline for a moment: keep trying until the code expires
    const j = r.json || {};
    if (r.status === 200 && j.access_token) {
      if (!j.refresh_token) return finish("Signed in, but Microsoft didn't return a refresh token (offline_access). Sign in again.");
      const cur = readJSON(pendingFile());
      if (!cur || cur.id !== id) return;
      try {
        saveTokenResponse(j, null);
      } catch (e) {
        return finish(e.message);
      }
      removePath(throttleFile()); // an old back-off must not break the first request
      let who = "";
      try {
        const me = graph("GET", "/me?$select=id,displayName,mail,userPrincipalName");
        const prev = account();
        if (prev && prev.id && prev.id !== me.id) removePath(cacheDir()); // someone else's data
        writeJSON(`${dataRoot()}/account.json`, {
          id: me.id, name: me.displayName || "", email: me.mail || me.userPrincipalName || "",
          client_id: CLIENT_ID, tenant: TENANT, signed_in_at: nowMs(),
        });
        who = me.displayName || me.mail || me.userPrincipalName || "";
      } catch (e) {
        writeJSON(`${dataRoot()}/account.json`, { id: "default", name: "", email: "", client_id: CLIENT_ID, tenant: TENANT, signed_in_at: nowMs() });
      }
      return finish(who ? `Signed in as ${who}` : "Signed in to Microsoft 365");
    }
    if (r.status === 429 || r.status >= 500) {
      interval = Math.min(interval + 5, 60);
      continue;
    }
    switch (j.error) {
      case "authorization_pending":
        continue;
      case "slow_down": // RFC 8628 §3.5: add 5 seconds to the interval for this and later requests
        interval = Math.min(interval + 5, 60);
        continue;
      case "authorization_declined":
        return finish("Sign-in was declined.");
      case "expired_token":
        return finish(`The sign-in code expired. Use the ${KW_ACCOUNT} keyword to try again.`);
      case "bad_verification_code":
        return finish(`Microsoft didn't recognize the sign-in code. Use the ${KW_ACCOUNT} keyword to try again.`);
      default:
        return finish(`Sign-in failed: ${aadError(j, r.status).message}`);
    }
  }
}

// Delegated permissions the last token response didn't grant (an admin can consent to fewer).
function missingScopes(granted) {
  if (!granted) return [];
  const have = String(granted).toLowerCase().split(/\s+/).map((s) => s.replace(/^https:\/\/graph\.microsoft\.com\//, ""));
  // X.ReadWrite and X.Read.All grant X.Read too.
  const covers = (h, s) => h === s || h.startsWith(`${s}.`) || h.startsWith(s.replace(/\.read$/, ".readwrite"));
  return SCOPES.split(" ").filter((s) => s !== "offline_access" && !have.some((h) => covers(h, s.toLowerCase())));
}

function logout() {
  kcDel(kcAccount());
  removePath(cacheDir());
  removePath(`${dataRoot()}/account.json`);
  removePath(pendingFile());
  return "Signed out of Microsoft 365";
}

function accountItems(query) {
  const cfg = configProblem();
  if (cfg) return [row(cfg, "↩ Open the setup guide", "error", "open", SETUP_URL)];
  const q = norm(query.trim());
  const rows = [];
  if (pendingLogin()) rows.push(...pendingRows());
  const t = loadTokens();
  const a = account();
  if (t) {
    const who = a && (a.name || a.email) ? `${a.name}${a.email ? ` (${a.email})` : ""}` : "your account";
    rows.push(info(`Signed in as ${who}`, `Tenant: ${TENANT}`, "account"));
    const missing = missingScopes(t.scope);
    if (missing.length) {
      rows.push(info(`Missing permission${missing.length > 1 ? "s" : ""}: ${missing.join(", ")}`,
        "Add them to your app registration (or ask for admin consent), then sign in again", "error"));
    }
    rows.push(row("Sign in again", "To switch accounts or consent to new permissions", "login", "login", "login"));
    rows.push(row("Refresh cached data", "Meetings, presence and the OneNote page index", "refresh", "clear-cache", "clear-cache"));
    rows.push(row("Sign out", "Removes the sign-in from the Keychain and clears the cache", "logout", "logout", "logout"));
  } else if (!pendingLogin()) {
    rows.push(signInRow());
  }
  rows.push(row("Open the setup guide", "How to register your app in Microsoft Entra ID", "info", "open", SETUP_URL));
  if (!q) return rows;
  const f = rows.filter((r) => norm(r.title).includes(q) || (q === "login" && r.variables && /^login/.test(r.variables.m365_action)) || (q === "logout" && r.variables && r.variables.m365_action === "logout"));
  return f.length ? f : rows;
}

// ---------- calendar ----------

function pad(n) {
  return String(n).padStart(2, "0");
}
function localDateKey(d) {
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}
// Graph returns "2026-09-26T09:00:00.0000000" (7 fractional digits, no zone) in the Prefer'd zone (UTC).
function parseGraphUTC(dt) {
  if (!dt) return NaN;
  let s = String(dt).replace(/(\.\d{3})\d+/, "$1");
  if (!/(Z|[+-]\d{2}:?\d{2})$/.test(s)) s += "Z";
  return Date.parse(s);
}
function dayRange(d) {
  const start = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const end = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1); // DST-safe
  return [start, end];
}

function fetchEvents(day) {
  const [start, end] = dayRange(day);
  const select = "id,subject,start,end,isAllDay,isCancelled,isOnlineMeeting,onlineMeeting,onlineMeetingProvider,onlineMeetingUrl,webLink,location,organizer,responseStatus,bodyPreview";
  const path = `/me/calendarView?startDateTime=${encodeURIComponent(start.toISOString())}&endDateTime=${encodeURIComponent(end.toISOString())}` +
    `&$select=${select}&$orderby=start/dateTime&$top=50`;
  const raw = graphAll(path, { maxPages: 10, headers: { Prefer: 'outlook.timezone="UTC"' } });
  const dayKey = localDateKey(start);
  const out = [];
  for (const e of raw) {
    if (e.isCancelled) continue;
    const join = (e.onlineMeeting && e.onlineMeeting.joinUrl) || e.onlineMeetingUrl || meetingLinkIn(e);
    const ev = {
      id: e.id,
      subject: e.subject || "",
      allDay: !!e.isAllDay,
      join: isHttpUrl(join) ? join : "",
      teams: e.onlineMeetingProvider === "teamsForBusiness" || /^https:\/\/teams\.(microsoft\.com|live\.com|cloud\.microsoft)\//i.test(join),
      web: isHttpUrl(e.webLink) ? e.webLink : "",
      location: (e.location && e.location.displayName) || "",
      organizer: (e.organizer && e.organizer.emailAddress && e.organizer.emailAddress.name) || "",
      response: (e.responseStatus && e.responseStatus.response) || "",
    };
    if (ev.allDay) {
      // All-day events are floating dates: compare calendar dates, not instants, and keep only
      // the ones that cover today (the UTC window can catch yesterday's or tomorrow's).
      ev.startDate = String((e.start || {}).dateTime || "").slice(0, 10);
      ev.endDate = String((e.end || {}).dateTime || "").slice(0, 10);
      if (!(ev.startDate <= dayKey && dayKey < (ev.endDate || ev.startDate + "~"))) continue;
    } else {
      ev.start = parseGraphUTC((e.start || {}).dateTime);
      ev.end = parseGraphUTC((e.end || {}).dateTime);
      if (isNaN(ev.start) || isNaN(ev.end)) continue;
      if (!(ev.start < end.getTime() && ev.end > start.getTime())) continue;
    }
    out.push(ev);
  }
  out.sort((a, b) => (a.allDay === b.allDay ? (a.start || 0) - (b.start || 0) : a.allDay ? -1 : 1));
  return out;
}

// Invitations from other organizations and other meeting services (Zoom, Google Meet, Webex…) often
// have no onlineMeeting: their join link is in the location or at the start of the body.
const MEETING_HOSTS = /^https:\/\/(?:[a-z0-9-]+\.)*(?:teams\.microsoft\.com|teams\.live\.com|teams\.cloud\.microsoft|zoom\.us|zoomgov\.com|meet\.google\.com|webex\.com|gotomeeting\.com|meet\.goto\.com|whereby\.com|chime\.aws)(?:[/?#]|$)/i;
function meetingLinkIn(e) {
  const loc = e.location || {};
  const texts = [loc.displayName, loc.locationUri, e.bodyPreview];
  for (const t of texts) {
    for (const u of String(t || "").match(/https:\/\/[^\s<>"'\u00A0]+/gi) || []) {
      const url = u.replace(/[)\].,;:!?>]+$/, "");
      if (MEETING_HOSTS.test(url) && isHttpUrl(url)) return url;
    }
  }
  return "";
}

function todayEvents() {
  const key = `events-${localDateKey(clock())}`;
  const res = swr(key, 300);
  // Keep only today's file.
  const dir = cacheDir();
  const names = ObjC.deepUnwrap(FM.contentsOfDirectoryAtPathError(dir, null)) || [];
  for (const n of names) if (/^events-\d{4}-\d{2}-\d{2}(\.error)?\.json$/.test(n) && !n.startsWith(key)) removePath(`${dir}/${n}`);
  return res;
}

let TIME_FMT = null;
function fmtDate(d, template) {
  const f = $.NSDateFormatter.alloc.init;
  const loc = env("M365_LOCALE", "");
  f.locale = loc ? $.NSLocale.localeWithLocaleIdentifier(loc) : $.NSLocale.currentLocale;
  f.setLocalizedDateFormatFromTemplate(template);
  return f.stringFromDate($.NSDate.dateWithTimeIntervalSince1970(d.getTime() / 1000)).js;
}
function fmtTime(ms) {
  if (!TIME_FMT) {
    TIME_FMT = $.NSDateFormatter.alloc.init;
    const loc = env("M365_LOCALE", "");
    TIME_FMT.locale = loc ? $.NSLocale.localeWithLocaleIdentifier(loc) : $.NSLocale.currentLocale;
    TIME_FMT.setLocalizedDateFormatFromTemplate("jmm");
  }
  return TIME_FMT.stringFromDate($.NSDate.dateWithTimeIntervalSince1970(ms / 1000)).js;
}
function duration(ms) {
  const m = Math.max(1, Math.round(ms / 60000));
  if (m < 60) return `${m} min`;
  const h = Math.floor(m / 60), r = m % 60;
  return r ? `${h} h ${r} min` : `${h} h`;
}
function eventWhen(ev, now) {
  if (ev.allDay) return { label: "All day", state: "allday" };
  const [dStart, dEnd] = dayRange(now);
  const s = ev.start < dStart.getTime() ? fmtDate(new Date(ev.start), "MMMd") + " " + fmtTime(ev.start) : fmtTime(ev.start);
  const e = ev.end > dEnd.getTime() ? fmtDate(new Date(ev.end), "MMMd") + " " + fmtTime(ev.end) : fmtTime(ev.end);
  const range = `${s}–${e}`;
  const t = now.getTime();
  if (t >= ev.end) return { label: `Ended · ${range}`, state: "ended" };
  if (t >= ev.start) return { label: `Now · ${range}`, state: "now" };
  return { label: `In ${duration(ev.start - t)} · ${range}`, state: "soon" };
}
const STATE_ORDER = { now: 0, soon: 1, allday: 2, ended: 3 };

function teamsLinks(httpsUrl) {
  // Work and school links (teams.microsoft.com, or teams.cloud.microsoft since 2024) open in the app.
  const m = /^https:\/\/teams\.(?:microsoft\.com|cloud\.microsoft)(\/l\/[^\s]*)$/i.exec(httpsUrl || "");
  return { app: m ? `msteams:${m[1]}` : "", web: httpsUrl || "" };
}

function meetingRow(ev, now) {
  const w = eventWhen(ev, now);
  const sub = [w.label, ev.organizer].filter(Boolean).join(" · ");
  const l = teamsLinks(ev.join);
  return row(ev.subject || "(No title)", sub, w.state === "ended" ? "meeting-ended" : "meeting", "join", ev.join,
    { m365_app_url: l.app, m365_web_url: l.web },
    {
      cmd: { action: "copy", arg: ev.join, subtitle: "Copy the join link" },
      alt: ev.web ? { action: "open", arg: ev.web, subtitle: "Open the event in Outlook" } : { valid: false, subtitle: "No Outlook link" },
    }, { text: { copy: ev.join, largetype: ev.subject || "" } });
}

// ---------- Teams ----------

const PRESENCES = [
  { key: "available", names: ["available", "online", "free"], label: "Available", availability: "Available", activity: "Available" },
  { key: "busy", names: ["busy"], label: "Busy", availability: "Busy", activity: "Busy" },
  { key: "dnd", names: ["dnd", "do not disturb", "donotdisturb"], label: "Do not disturb", availability: "DoNotDisturb", activity: "DoNotDisturb" },
  { key: "brb", names: ["brb", "be right back", "berightback"], label: "Be right back", availability: "BeRightBack", activity: "BeRightBack" },
  { key: "away", names: ["away", "appear away"], label: "Appear away", availability: "Away", activity: "Away" },
  { key: "offline", names: ["offline", "appear offline", "invisible"], label: "Appear offline", availability: "Offline", activity: "OffWork" },
  { key: "reset", names: ["reset", "clear", "automatic", "auto"], label: "Reset status", clear: true },
];

// "90m", "2h", "1 d", "1h30m" -> minutes
function parseDuration(s) {
  const m = String(s).trim().toLowerCase().match(/^(?:(\d+)\s*d(?:ays?)?)?\s*(?:(\d+)\s*h(?:rs?|ours?)?)?\s*(?:(\d+)\s*m(?:in(?:ute)?s?)?)?$/);
  if (!m || (!m[1] && !m[2] && !m[3])) return null;
  const mins = (+m[1] || 0) * 1440 + (+m[2] || 0) * 60 + (+m[3] || 0);
  return mins > 0 ? mins : null;
}
function isoDuration(mins) {
  const d = Math.floor(mins / 1440), h = Math.floor((mins % 1440) / 60), m = mins % 60;
  return `P${d ? d + "D" : ""}${h || m ? "T" : ""}${h ? h + "H" : ""}${m ? m + "M" : ""}`;
}
function humanDuration(mins) {
  if (mins % 1440 === 0) return plural(mins / 1440, "day");
  if (mins % 60 === 0) return plural(mins / 60, "hour");
  if (mins < 60) return plural(mins, "minute");
  return `${Math.floor(mins / 60)} h ${mins % 60} min`;
}

function fetchPresence() {
  const p = graph("GET", "/me/presence");
  const m = (p.statusMessage && p.statusMessage.message) || {};
  let text = String(m.content || "");
  if (String(m.contentType || "").toLowerCase() === "html") {
    text = text.replace(/<br\s*\/?>/gi, " ").replace(/<[^>]*>/g, "")
      .replace(/&nbsp;/g, " ").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, "&");
  }
  return { availability: p.availability || "", activity: p.activity || "", message: oneLine(text, 280) };
}
function readable(s) {
  return String(s || "").replace(/([a-z])([A-Z])/g, "$1 $2").replace(/^./, (c) => c.toUpperCase()).replace(/ ([A-Z])/g, (m, c) => " " + c.toLowerCase());
}

function statusItems(rest) {
  if (PERSONAL_ONLY) return [info("Teams status needs a work or school account", "Personal Microsoft accounts (tenant “consumers”) have no presence", "presence-offline")];
  const words = rest.trim().split(/\s+/).filter(Boolean);
  let mins = null;
  // Duration may be the last one or two words: "busy 2h", "busy 1 h", "dnd 1h 30m".
  for (let n = Math.min(3, words.length); n >= 1 && mins === null; n--) {
    const d = parseDuration(words.slice(-n).join(" "));
    if (d !== null && words.length - n >= 0) {
      mins = Math.min(d, 7 * 1440);
      words.splice(-n, n);
    }
  }
  const wanted = norm(words.join(" ")).replace(/\s+/g, "");
  const matches = PRESENCES.filter((p) => !wanted || p.names.concat([p.label]).some((n) => norm(n).replace(/\s+/g, "").startsWith(wanted)));
  const rows = [];
  try {
    const cur = swr("presence", 60);
    if (cur.data && cur.data.availability) {
      const a = cur.data.availability;
      // Teams only shows a status you set while you're signed in to a Teams app somewhere.
      const sub = a === "Offline" || a === "PresenceUnknown"
        ? "A status you set shows only while you’re signed in to Teams"
        : cur.data.activity && cur.data.activity !== a ? readable(cur.data.activity) : "";
      rows.push(info(`Current status: ${readable(a)}`, sub, "presence-" + presenceIcon(a)));
    }
    if (cur.data && cur.data.message) rows.push(info(`Status message: ${cur.data.message}`, "Change or clear it with “message”", "info"));
  } catch (e) {
    if (["auth", "network", "keychain"].includes(e.kind)) return errorItems(e);
    rows.push(info("Couldn't read your status", presenceHint(e), "error"));
  }
  if (!wanted && !mins) rows.push(Object.assign(info("Set a status message…", "Shows next to your name in Teams, like “message Out for lunch :: 1h”", "info"), { autocomplete: "message ", valid: false }));
  if (!matches.length) {
    rows.push(info("Unknown status", "Try available, busy, dnd, brb, away, offline or reset, optionally followed by a duration like 2h", "error"));
    return rows;
  }
  for (const p of matches) {
    if (p.clear) {
      rows.push(row(p.label, "Let Teams set your status automatically again", "presence-reset", "presence", p.key, { m365_presence: p.key }));
      continue;
    }
    const dflt = p.availability === "Busy" || p.availability === "DoNotDisturb" ? "1 day" : "7 days";
    const sub = mins ? `Set for ${humanDuration(mins)}` : `Set until you change it (up to ${dflt}) · add a duration like 2h`;
    rows.push(row(p.label, sub, "presence-" + p.key, "presence", p.key, { m365_presence: p.key, m365_minutes: mins ? String(mins) : "" }));
  }
  return rows;
}
function presenceIcon(availability) {
  const p = PRESENCES.find((x) => x.availability === availability);
  return p ? p.key : "reset";
}
function presenceHint(e) {
  if (e.status === 403 || e.status === 404 || e.status === 400) return "Presence works with work or school accounts only";
  return e.message;
}

function setPresence(key, mins) {
  const p = PRESENCES.find((x) => x.key === key);
  if (!p) return "Unknown status";
  try {
    if (p.clear) {
      graph("POST", "/me/presence/clearUserPreferredPresence", { json: {} });
    } else {
      const body = { availability: p.availability, activity: p.activity };
      if (mins) body.expirationDuration = isoDuration(mins);
      graph("POST", "/me/presence/setUserPreferredPresence", { json: body });
    }
  } catch (e) {
    if (e.kind === "graph" || e.kind === "consent") return `Couldn't set your status: ${presenceHint(e)}`;
    return e.message === "Not signed in" ? `Sign in first via the ${KW_ACCOUNT} keyword` : e.message;
  }
  removePath(cacheFile("presence"));
  if (p.clear) return "Status reset: Teams sets it automatically";
  const dflt = p.availability === "Busy" || p.availability === "DoNotDisturb" ? "1 day" : "7 days";
  return `Status set to ${p.label} ${mins ? `for ${humanDuration(mins)}` : `(expires in ${dflt})`}`;
}

const STATUS_MESSAGE_MAX = 280; // Teams' limit
function statusMessageItems(rest) {
  if (PERSONAL_ONLY) return [info("Teams status messages need a work or school account", "Personal Microsoft accounts (tenant “consumers”) have no presence", "presence-offline")];
  const m = rest.match(/^([\s\S]*?)\s*::\s*(.*)$/);
  const text = clean(m ? m[1] : rest).replace(/\s+/g, " ").trim();
  let mins = null;
  if (m && m[2].trim()) {
    mins = parseDuration(m[2]);
    if (mins === null) return [info("Unknown duration", "Add a duration after “::” like 30m, 2h or 1d", "error")];
    mins = Math.min(mins, 7 * 1440);
  }
  const rows = [];
  let current = "";
  try {
    current = (swr("presence", 60).data || {}).message || "";
  } catch (e) {
    if (["auth", "network", "keychain"].includes(e.kind)) return errorItems(e);
  }
  if (!text) {
    if (current) rows.push(info(`Status message: ${current}`, "Type a new message to replace it", "info"));
    rows.push(info("Type your status message", "Add “:: 2h” to clear it after a while", "info"));
    if (current) rows.push(row("Clear the status message", "Removes the message next to your name in Teams", "presence-reset", "status-message", "clear", { m365_message: "", m365_minutes: "" }));
    return rows;
  }
  if ([...text].length > STATUS_MESSAGE_MAX) return [info(`The message is too long (${[...text].length} of ${STATUS_MESSAGE_MAX} characters)`, "Teams allows 280 characters", "error")];
  const sub = mins ? `Clears after ${humanDuration(mins)}` : "Shows next to your name in Teams until you change it · add “:: 2h” to clear it after a while";
  rows.push(row(`Set status message “${text}”`, sub, "info", "status-message", text, { m365_message: text, m365_minutes: mins ? String(mins) : "" }));
  return rows;
}

function setStatusMessage(text, mins) {
  const msg = String(text || "").trim();
  const body = { statusMessage: { message: { content: msg, contentType: "text" } } };
  if (msg && mins) {
    // dateTime without a zone designator, interpreted in timeZone.
    body.statusMessage.expiryDateTime = { dateTime: new Date(nowMs() + mins * 60000).toISOString().replace(/Z$/, ""), timeZone: "UTC" };
  }
  try {
    graph("POST", "/me/presence/setStatusMessage", { json: body });
  } catch (e) {
    if (e.kind === "graph" || e.kind === "consent") return `Couldn't set your status message: ${presenceHint(e)}`;
    return e.message === "Not signed in" ? `Sign in first via the ${KW_ACCOUNT} keyword` : e.message;
  }
  removePath(cacheFile("presence"));
  if (!msg) return "Status message cleared";
  return `Status message set${mins ? ` for ${humanDuration(mins)}` : ""}`;
}

// $search values are KQL phrases in double quotes; escape backslashes and quotes.
function searchParam(q) {
  return encodeURIComponent(`"${q.replace(/\\/g, "\\\\").replace(/"/g, '\\"')}"`);
}

function peopleSearch(q) {
  const name = `people-${hash(norm(q))}`;
  const c = readCache(name);
  if (c && nowMs() - c.fetched_at < 600000) return c.data;
  const res = graph("GET", `/me/people?$search=${searchParam(q)}&$top=15&$select=id,displayName,scoredEmailAddresses,jobTitle,department,userPrincipalName,personType`);
  const people = [];
  const seen = new Set();
  for (const p of res.value || []) {
    if (p.personType && p.personType.class && p.personType.class !== "Person") continue;
    const email = ((p.scoredEmailAddresses || [])[0] || {}).address || p.userPrincipalName || "";
    if (!email || seen.has(email.toLowerCase())) continue;
    seen.add(email.toLowerCase());
    people.push({ name: p.displayName || email, email, job: p.jobTitle || "", dept: p.department || "" });
  }
  storeCache(name, people);
  return people;
}

function chatLinks(email) {
  const q = `users=${encodeURIComponent(email)}`;
  return { app: `msteams:/l/chat/0/0?${q}`, web: `https://teams.microsoft.com/l/chat/0/0?${q}` };
}

function personRow(p) {
  const l = chatLinks(p.email);
  const call = `https://teams.microsoft.com/l/call/0/0?users=${encodeURIComponent(p.email)}&withVideo=true`;
  return row(p.name, [p.email, p.job, p.dept].filter(Boolean).join(" · "), "person", "chat", l.web,
    { m365_app_url: l.app, m365_web_url: l.web },
    {
      cmd: { action: "copy", arg: p.email, subtitle: `Copy ${p.email}` },
      ctrl: { action: "copy", arg: l.web, subtitle: "Copy the Teams chat link" },
      alt: { action: "chat", arg: call, subtitle: "Start a video call", vars: { m365_app_url: call.replace("https://teams.microsoft.com", "msteams:"), m365_web_url: call } },
    }, { text: { copy: p.email, largetype: p.name } });
}

function teamsItems(query) {
  requireConfig();
  if (!loadTokens()) return errorItems(new M365Error("auth", "Not signed in"));
  const q = query.trim();
  const sm = q.match(/^(status|presence)\b\s*(.*)$/i);
  if (sm) return { items: statusItems(sm[2]) };
  const mm = q.match(/^message\b\s*([\s\S]*)$/i);
  if (mm) return { items: statusMessageItems(mm[1]) };
  const now = clock();
  const rows = [];
  let extra = {};
  let meetings = [];
  try {
    const ev = todayEvents();
    if (ev.rerun) extra.rerun = ev.rerun;
    if (ev.warning) rows.push(warningRow(ev.warning, ev.age));
    meetings = (ev.data || []).filter((e) => e.join && e.teams && e.response !== "declined");
  } catch (e) {
    if (["auth", "config", "keychain"].includes(e.kind)) return { items: errorItems(e) };
    rows.push(...errorItems(e));
  }
  const words = norm(q).split(/\s+/).filter(Boolean);
  const hit = (ev) => words.every((w) => norm(`${ev.subject} ${ev.organizer}`).includes(w));
  const list = meetings.filter(hit).map((ev) => ({ ev, w: eventWhen(ev, now) }));
  list.sort((a, b) => STATE_ORDER[a.w.state] - STATE_ORDER[b.w.state] || (a.ev.start || 0) - (b.ev.start || 0));
  for (const { ev } of list) rows.push(meetingRow(ev, now));
  if (!q) {
    if (!meetings.length && !rows.length) rows.push(info("No Teams meetings today", "Type a name to chat with someone", "meeting-ended"));
    rows.push(Object.assign(info("Set your status…", "Available, busy, do not disturb, be right back, away, offline", "presence-available"), { autocomplete: "status ", valid: false }));
    return { items: rows, extra };
  }
  if (q.length < 2) {
    if (!list.length) rows.push(info("Keep typing to find people", "Type at least two letters of a name or email address", "person"));
  } else {
    try {
      const people = peopleSearch(q);
      for (const p of people) rows.push(personRow(p));
      if (!people.length && !list.length) rows.push(info(`No meetings or people match “${oneLine(q, 60)}”`, "", "person"));
    } catch (e) {
      if (e.kind === "auth" || e.kind === "keychain") return { items: errorItems(e) };
      rows.push(...errorItems(e).map((r) => Object.assign(r, { subtitle: r.subtitle || "People search" })));
    }
  }
  return { items: rows, extra };
}

// ---------- Outlook ----------

function mailDate(ms, now) {
  const d = new Date(ms);
  if (localDateKey(d) === localDateKey(now)) return fmtTime(ms);
  if (d.getFullYear() === now.getFullYear()) return fmtDate(d, "MMMd");
  return fmtDate(d, "yMMMd");
}

function mailSearch(q) {
  const name = `mail-${hash(q)}`;
  const c = readCache(name);
  if (c && nowMs() - c.fetched_at < 120000) return c.data;
  const res = graph("GET", `/me/messages?$search=${searchParam(q)}&$top=25&$select=id,subject,from,receivedDateTime,bodyPreview,webLink,isRead,hasAttachments`);
  const mails = (res.value || []).map((m) => ({
    subject: m.subject || "",
    from: (m.from && m.from.emailAddress && (m.from.emailAddress.name || m.from.emailAddress.address)) || "",
    fromEmail: (m.from && m.from.emailAddress && m.from.emailAddress.address) || "",
    received: Date.parse(m.receivedDateTime || "") || 0,
    preview: oneLine(m.bodyPreview || "", 100),
    web: isHttpUrl(m.webLink) ? m.webLink : "",
    unread: m.isRead === false,
    attach: !!m.hasAttachments,
  }));
  storeCache(name, mails);
  return mails;
}

function mailRow(m, now) {
  const sub = [m.from, m.received ? mailDate(m.received, now) : "", m.preview].filter(Boolean).join(" · ");
  const title = `${m.unread ? "● " : ""}${m.subject || "(No subject)"}${m.attach ? " 📎" : ""}`;
  const r = m.web
    ? row(title, sub, m.unread ? "mail-unread" : "mail", "open", m.web, {}, {
      cmd: m.fromEmail ? { action: "copy", arg: m.fromEmail, subtitle: `Copy ${m.fromEmail}` } : { valid: false, subtitle: "No sender address" },
    }, { text: { copy: m.web, largetype: m.subject || "" } })
    : info(title, sub, "mail");
  return r;
}

function agendaRow(ev, now) {
  const w = eventWhen(ev, now);
  const sub = [w.label, ev.location, ev.response === "declined" ? "Declined" : ""].filter(Boolean).join(" · ");
  const l = teamsLinks(ev.join);
  const joinMod = ev.join
    ? { action: "join", arg: ev.join, subtitle: "Join the online meeting", vars: { m365_app_url: l.app, m365_web_url: l.web } }
    : { valid: false, subtitle: "Not an online meeting" };
  if (!ev.web) return info(ev.subject || "(No title)", sub, "calendar");
  return row(ev.subject || "(No title)", sub, w.state === "ended" ? "meeting-ended" : "calendar", "open", ev.web, {}, { cmd: joinMod },
    { text: { copy: ev.web, largetype: ev.subject || "" } });
}

function outlookItems(query) {
  requireConfig();
  if (!loadTokens()) return { items: errorItems(new M365Error("auth", "Not signed in")) };
  const q = query.trim();
  const now = clock();
  if (!q) {
    const rows = [];
    let extra = {};
    try {
      const ev = todayEvents();
      if (ev.rerun) extra.rerun = ev.rerun;
      if (ev.warning) rows.push(warningRow(ev.warning, ev.age));
      const list = (ev.data || []).slice();
      for (const e of list) rows.push(agendaRow(e, now));
      if (!list.length) rows.push(info("Nothing on your calendar today", "Type to search your mail", "calendar"));
    } catch (e) {
      return { items: errorItems(e) };
    }
    rows.push(info("Type to search your mail", "Words, from:name, subject:word, hasAttachments:yes…", "mail"));
    return { items: rows, extra };
  }
  try {
    const mails = mailSearch(q);
    if (!mails.length) return { items: [info(`No mail matches “${oneLine(q, 60)}”`, "Try other words", "mail")] };
    return { items: mails.map((m) => mailRow(m, now)) };
  } catch (e) {
    return { items: errorItems(e) };
  }
}

// ---------- OneNote ----------

const ONENOTE_PAGE_SELECT = "id,title,links,lastModifiedDateTime";

function pageEntry(p, section, notebook) {
  const links = p.links || {};
  const c = ((links.oneNoteClientUrl || {}).href) || "";
  const w = ((links.oneNoteWebUrl || {}).href) || "";
  return {
    id: p.id,
    t: p.title || "",
    s: section || (p.parentSection && p.parentSection.displayName) || "",
    n: notebook || (p.parentNotebook && p.parentNotebook.displayName) || "",
    m: Date.parse(p.lastModifiedDateTime || "") || 0,
    c: /^onenote:/i.test(c) ? c : "",
    w: isHttpUrl(w) ? w : "",
  };
}

// Graph has no usable full-text search for OneNote pages ($search was beta-only and is gone,
// and $filter on title only works per request). So we keep a title index in the cache and
// search it locally. Listing all pages at once fails for accounts with many sections
// (OneNote error 20266); then we list them section by section.
function buildOneNoteIndex() {
  const MAX = 5000;
  let pages;
  try {
    pages = graphAll(`/me/onenote/pages?$select=${ONENOTE_PAGE_SELECT}&$expand=parentSection($select=id,displayName),parentNotebook($select=id,displayName)&$orderby=lastModifiedDateTime%20desc&$top=100`, { maxPages: MAX / 100 })
      .map((p) => pageEntry(p));
  } catch (e) {
    if (!(e.kind === "graph" && (e.code === "20266" || /maximum (number of )?sections/i.test(e.message)))) throw e;
    pages = [];
    const sections = fetchSections();
    for (const s of sections) {
      if (pages.length >= MAX) break;
      const list = graphAll(`/me/onenote/sections/${encodeURIComponent(s.id)}/pages?$select=${ONENOTE_PAGE_SELECT}&$top=100`, { maxPages: 20 });
      for (const p of list) pages.push(pageEntry(p, s.name, s.notebook));
    }
    pages.sort((a, b) => b.m - a.m);
  }
  return pages.slice(0, MAX);
}

function fetchSections() {
  return graphAll("/me/onenote/sections?$select=id,displayName&$expand=parentNotebook($select=id,displayName)&$top=100", { maxPages: 20 })
    .map((s) => ({ id: s.id, name: s.displayName || "", notebook: (s.parentNotebook && s.parentNotebook.displayName) || "" }));
}

const WORDCHAR = /[\p{L}\p{N}]/u;
function atWordStart(text, w) {
  let i = text.indexOf(w);
  while (i >= 0) {
    if (i === 0 || !WORDCHAR.test(text[i - 1])) return true;
    i = text.indexOf(w, i + 1);
  }
  return false;
}
function scorePage(p, words) {
  const t = norm(p.t), path = norm(`${p.n} ${p.s}`);
  let score = 0;
  for (const w of words) {
    if (t.startsWith(w)) score += 30;
    else if (atWordStart(t, w)) score += 20;
    else if (t.includes(w)) score += 10;
    else if (path.includes(w)) score += 3;
    else return -1;
  }
  return score;
}

function pageRow(p, now) {
  const preferApp = ONENOTE_OPEN !== "web";
  const primary = preferApp && p.c ? p.c : p.w || p.c;
  const sub = [[p.n, p.s].filter(Boolean).join(" › "), p.m ? `edited ${mailDate(p.m, now)}` : ""].filter(Boolean).join(" · ");
  return row(p.t || "Untitled Page", sub, "page", "open-note", primary, { m365_app_url: p.c, m365_web_url: p.w }, {
    cmd: p.w ? { action: "open", arg: p.w, subtitle: "Open in OneNote on the web" } : { valid: false, subtitle: "No web link" },
    alt: p.w ? { action: "copy", arg: p.w, subtitle: "Copy the web link" } : { valid: false, subtitle: "No web link" },
  }, { text: { copy: p.w || p.c, largetype: p.t || "" } });
}

function newPageItems(rest) {
  const m = rest.match(/^(.*?)\s*::\s*([\s\S]*)$/);
  const title = (m ? m[1] : rest).trim();
  const inline = m ? m[2] : null;
  const where = ONENOTE_SECTION ? `in ${ONENOTE_SECTION}` : "in your default section";
  if (!title) return [info("Type a title for the new page", `Creates a page ${where} with the clipboard as its text · add “:: text” to type the text instead`, "new-page")];
  let body, src;
  if (inline !== null) {
    body = inline;
    src = inline.trim() ? `with “${oneLine(inline, 40)}”` : "empty";
  } else {
    try {
      body = clipboardText();
    } catch (e) {
      return [info(e.message, "Type the text after “::” instead", "error")];
    }
    src = body.trim() ? `with the clipboard (${plural([...body].length, "character")})` : "empty (the clipboard is empty)";
  }
  const vars = { m365_title: title, m365_body_source: inline !== null ? "inline" : "clipboard", m365_body: inline !== null ? inline : "" };
  return [row(`Create “${title}”`, `New page ${where}, ${src}`, "new-page", "create-page", title, vars, {
    cmd: { action: "create-page", subtitle: "Create the page and open it", vars: Object.assign({}, vars, { m365_open: "1" }) },
    alt: { action: "create-page", subtitle: "Create an empty page", vars: Object.assign({}, vars, { m365_body_source: "none", m365_body: "" }) },
  })];
}

function onenoteItems(query) {
  requireConfig();
  if (!loadTokens()) return { items: errorItems(new M365Error("auth", "Not signed in")) };
  const q = query.trim();
  const nm = q.match(/^new(?:\s+([\s\S]*))?$/i);
  const now = clock();
  if (nm) {
    // Also offer pages whose title matches, e.g. a page called "New ideas" (cached index only).
    const idx = readCache("onenote-index");
    const words = norm(q).split(/\s+/).filter(Boolean);
    const hits = idx && Array.isArray(idx.data)
      ? idx.data.map((p) => ({ p, s: scorePage(p, words) })).filter((x) => x.s >= 0).sort((a, b) => b.s - a.s || b.p.m - a.p.m).slice(0, 10).map((x) => pageRow(x.p, now))
      : [];
    return { items: newPageItems(nm[1] || "").concat(hits) };
  }
  let res;
  try {
    res = swr("onenote-index", 1800, { async: true });
  } catch (e) {
    return { items: errorItems(e) };
  }
  if (res.loading) return { items: [info("Indexing your OneNote pages…", "The first time can take a moment", "refresh")], extra: { rerun: res.rerun } };
  const rows = [];
  const extra = res.rerun ? { rerun: res.rerun } : {};
  if (res.warning) rows.push(warningRow(res.warning, res.age));
  const pages = res.data || [];
  const words = norm(q).split(/\s+/).filter(Boolean);
  let hits;
  if (!words.length) hits = pages.slice(0, 20);
  else {
    hits = pages.map((p) => ({ p, s: scorePage(p, words) })).filter((x) => x.s >= 0)
      .sort((a, b) => b.s - a.s || b.p.m - a.p.m).slice(0, 50).map((x) => x.p);
  }
  for (const p of hits) rows.push(pageRow(p, now));
  if (!hits.length) rows.push(info(pages.length ? `No page titles match “${oneLine(q, 60)}”` : "No OneNote pages found", `Searched ${plural(pages.length, "page title")}`, "page"));
  rows.push(Object.assign(info("Create a new page…", `${KW_ONENOTE} new <title> · the clipboard becomes the page text`, "new-page"), { autocomplete: "new ", valid: false }));
  if (!q) rows.push(row("Rebuild the page index", `${plural(pages.length, "page")}, updated ${ago(res.age * 1000)}`, "refresh", "reindex", "reindex"));
  return { items: rows, extra };
}

function escapeXml(s) {
  return String(s)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, "") // not allowed in XML
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}
function localISO(d) {
  const off = -d.getTimezoneOffset();
  const sign = off >= 0 ? "+" : "-";
  const a = Math.abs(off);
  return `${localDateKey(d)}T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}${sign}${pad(Math.floor(a / 60))}:${pad(a % 60)}`;
}
function pageHTML(title, text, when) {
  const paras = String(text || "").replace(/\r\n?/g, "\n").split(/\n{2,}/).filter((p) => p.trim() !== "")
    .map((p) => `<p>${p.split("\n").map(escapeXml).join("<br/>")}</p>`);
  return `<!DOCTYPE html>\n<html>\n<head>\n<title>${escapeXml(title)}</title>\n<meta name="created" content="${localISO(when)}" />\n</head>\n<body>\n${paras.join("\n")}\n</body>\n</html>\n`;
}

function resolveSectionPath() {
  if (!ONENOTE_SECTION) return "/me/onenote/pages";
  const want = norm(ONENOTE_SECTION).replace(/\s*\/\s*/g, "/");
  const sections = swr("onenote-sections", 3600).data || [];
  const byPath = sections.filter((s) => norm(`${s.notebook}/${s.name}`) === want);
  const byName = sections.filter((s) => norm(s.name) === want);
  const s = byPath[0] || (byName.length === 1 ? byName[0] : null);
  if (s) return `/me/onenote/sections/${encodeURIComponent(s.id)}/pages`;
  if (byName.length > 1) throw new M365Error("graph", `Several sections are named “${ONENOTE_SECTION}”: use Notebook/Section in the Workflow’s Configuration`);
  if (ONENOTE_SECTION.includes("/")) throw new M365Error("graph", `Section “${ONENOTE_SECTION}” not found`);
  // A plain name: OneNote finds it in the default notebook (or creates it), but can't create
  // a section whose name has any of these characters.
  if (/[?*\\/:<>|&#"%~]/.test(ONENOTE_SECTION)) throw new M365Error("graph", `Section “${ONENOTE_SECTION}” not found (and OneNote can't create a section with that name)`);
  return `/me/onenote/pages?sectionName=${encodeURIComponent(ONENOTE_SECTION)}`;
}

function createPage() {
  const title = env("m365_title", "").trim();
  if (!title) return "Type a title for the new page";
  const src = env("m365_body_source", "clipboard");
  const body = src === "inline" ? env("m365_body", "") : src === "clipboard" ? clipboardText() : "";
  let page;
  const post = () => graph("POST", resolveSectionPath(), { body: pageHTML(title, body, clock()), headers: { "Content-Type": "text/html; charset=utf-8" }, timeout: 40 });
  try {
    try {
      page = post();
    } catch (e) {
      // The cached section list may point at a deleted or moved section: reload it once.
      if (!(e.status === 404 && ONENOTE_SECTION)) throw e;
      removePath(cacheFile("onenote-sections"));
      page = post();
    }
  } catch (e) {
    if (e.kind === "graph" && e.status === 507) return "That section is full: choose another one in the Workflow’s Configuration";
    return `Couldn't create the page: ${e.message}`;
  }
  const entry = pageEntry(page, page.parentSection && page.parentSection.displayName, "");
  entry.m = nowMs();
  // Make the new page searchable right away.
  const idx = readCache("onenote-index");
  if (idx && Array.isArray(idx.data)) {
    idx.data.unshift(entry);
    writeJSON(cacheFile("onenote-index"), idx);
  }
  if (env("m365_open", "") === "1") {
    try {
      openPreferApp(entry.c, entry.w, ONENOTE_OPEN !== "web");
    } catch (e) {
      /* the page exists; opening is best effort */
    }
  }
  return `Created “${title}”`;
}

// ---------- actions ----------

function act(arg) {
  const a = env("m365_action", "");
  try {
    switch (a) {
      case "login":
        return login();
      case "login-show": {
        const p = pendingLogin();
        if (!p) return `The sign-in code expired. Use the ${KW_ACCOUNT} keyword to try again.`;
        copyText(p.user_code, true);
        openURL(p.verification_uri);
        return `Code ${p.user_code} copied`;
      }
      case "cancel-login":
        removePath(pendingFile());
        return "Sign-in cancelled";
      case "logout":
        return logout();
      case "clear-cache":
        removePath(cacheDir());
        removePath(throttleFile());
        startRefresh("onenote-index");
        return "Cleared cached data";
      case "reindex":
        startRefresh("onenote-index");
        return "Rebuilding the OneNote page index";
      case "join":
      case "chat":
        openPreferApp(env("m365_app_url", ""), env("m365_web_url", "") || arg, TEAMS_OPEN !== "web");
        return "";
      case "open-note":
        openPreferApp(env("m365_app_url", ""), env("m365_web_url", "") || arg, ONENOTE_OPEN !== "web");
        return "";
      case "open":
        openURL(arg);
        return "";
      case "copy":
        copyText(arg);
        return `Copied ${oneLine(arg, 60)}`;
      case "presence":
        return setPresence(env("m365_presence", arg), Number(env("m365_minutes", "")) || null);
      case "status-message":
        return setStatusMessage(env("m365_message", ""), Number(env("m365_minutes", "")) || null);
      case "create-page":
        return createPage();
      default:
        if (isHttpUrl(arg)) {
          openURL(arg);
          return "";
        }
        return "";
    }
  } catch (e) {
    return e.message || String(e);
  }
}

// ---------- entry point ----------

function run(argv) {
  const cmd = argv[0] || "";
  const query = argv[1] || "";
  // osascript prints "\n" for an empty string but nothing for undefined: an action without a message
  // must print nothing, or the Notification ("only show if populated") may pop up empty.
  if (cmd === "act") return act(query) || undefined;
  if (cmd === "poll") return void poll();
  if (cmd === "refresh") return void refreshJob(query);
  const filters = { teams: teamsItems, outlook: outlookItems, onenote: onenoteItems, account: accountItems };
  const fn = Object.prototype.hasOwnProperty.call(filters, cmd) ? filters[cmd] : null;
  if (!fn) return output([info(`Unknown command ${cmd}`, "", "error")]);
  try {
    const r = fn(query);
    return Array.isArray(r) ? output(r) : output(r.items, r.extra || {});
  } catch (e) {
    return output(errorItems(e));
  }
}
