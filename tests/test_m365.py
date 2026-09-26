#!/usr/bin/env python3
"""End-to-end tests: run the Script Filters and actions the way Alfred does, against a local mock
of the Microsoft identity platform (device code, token, refresh) and Microsoft Graph.
No real Microsoft calls: the base URLs point at a http.server in this process, and the Keychain
is replaced by a temporary directory (M365_TEST_KEYCHAIN_DIR)."""
import json, os, plistlib, shutil, stat, subprocess, sys, tempfile, threading, time, unittest
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse, unquote

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "src")
CLIENT = "11111111-2222-3333-4444-555555555555"
KC_FILE = f"oauth_{CLIENT}_common.secret"


# ---------------------------------------------------------------- mock server

class Mock:
    def __init__(self):
        self.reset()

    def reset(self):
        self.lock = threading.Lock()
        self.requests = []          # (method, path, query, headers, body)
        self.device_queue = []      # responses for device_code polls: "pending", "slow_down", "ok", "declined", ...
        self.device_polls = []      # (time, device_code)
        self.devicecode = {"device_code": "DEVCODE-secret-123", "user_code": "ABCD-EFGH",
                           "verification_uri": f"{self.base}/devicelogin", "expires_in": 900, "interval": 5,
                           "message": "To sign in…"}
        self.devicecode_error = None
        self.valid_access = set()
        self.valid_refresh = set()
        self.refresh_error = None
        self.issued = 0
        self.events = []
        self.presence = {"availability": "Available", "activity": "Available"}
        self.people = []
        self.mails = []
        self.pages = []
        self.sections = []
        self.too_many_sections = False
        self.overrides = {}         # (method, path) -> list of (status, headers, json)
        self.created = []
        self.foreign_next = False
        self.me = {"id": "user-1", "displayName": "Zoë Tester", "mail": "zoe@contoso.com"}
        self.device_scopes = []

    def issue(self):
        self.issued += 1
        at, rt = f"AT-{self.issued}", f"RT-{self.issued}"
        self.valid_access.add(at)
        self.valid_refresh.add(rt)
        return {"token_type": "Bearer", "access_token": at, "refresh_token": rt, "expires_in": 3599,
                "scope": "User.Read Calendars.Read"}

    def graph_requests(self, path):
        return [r for r in self.requests if r[1] == path]


MOCK = Mock.__new__(Mock)


class Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def log_message(self, *a):
        pass

    def send(self, status, body=None, headers=None):
        data = b"" if body is None else (body if isinstance(body, bytes) else json.dumps(body).encode())
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        for k, v in (headers or {}).items():
            self.send_header(k, v)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        self.route("GET")

    def do_POST(self):
        self.route("POST")

    def route(self, method):
        u = urlparse(self.path)
        n = int(self.headers.get("Content-Length") or 0)
        body = self.rfile.read(n).decode("utf-8") if n else ""
        q = parse_qs(u.query)
        with MOCK.lock:
            MOCK.requests.append((method, u.path, q, dict(self.headers), body))
            ov = MOCK.overrides.get((method, u.path))
            if ov:
                status, headers, js = ov.pop(0)
                if not ov:
                    del MOCK.overrides[(method, u.path)]
                return self.send(status, js, headers)
        if u.path.endswith("/oauth2/v2.0/devicecode"):
            return self.devicecode(parse_qs(body))
        if u.path.endswith("/oauth2/v2.0/token"):
            return self.token(parse_qs(body))
        if u.path.startswith("/v1.0/"):
            auth = self.headers.get("Authorization", "")
            if not auth.startswith("Bearer ") or auth[7:] not in MOCK.valid_access:
                return self.send(401, {"error": {"code": "InvalidAuthenticationToken", "message": "Access token has expired."}})
            return self.graph(method, u.path[len("/v1.0"):], q, body)
        self.send(404, {"error": {"code": "NotFound", "message": "nope"}})

    def devicecode(self, form):
        if MOCK.devicecode_error:
            return self.send(400, MOCK.devicecode_error)
        assert form["client_id"] == [CLIENT]
        assert "offline_access" in form["scope"][0]
        MOCK.device_scopes.append(form["scope"][0])
        self.send(200, MOCK.devicecode)

    def token(self, form):
        g = form.get("grant_type", [""])[0]
        if g == "urn:ietf:params:oauth:grant-type:device_code":
            MOCK.device_polls.append((time.time(), form.get("device_code", [""])[0]))
            nxt = MOCK.device_queue.pop(0) if MOCK.device_queue else "pending"
            if nxt == "ok":
                return self.send(200, MOCK.issue())
            err = {"pending": "authorization_pending", "slow_down": "slow_down", "declined": "authorization_declined",
                   "expired": "expired_token", "bad": "bad_verification_code"}.get(nxt, nxt)
            return self.send(400, {"error": err, "error_description": f"AADSTS70016: {err}", "error_codes": [70016]})
        if g == "refresh_token":
            if MOCK.refresh_error:
                return self.send(400, MOCK.refresh_error)
            rt = form.get("refresh_token", [""])[0]
            if rt not in MOCK.valid_refresh:
                return self.send(400, {"error": "invalid_grant", "error_description": "AADSTS70008: The refresh token has expired.", "error_codes": [70008]})
            assert "offline_access" in form.get("scope", [""])[0]
            return self.send(200, MOCK.issue())
        self.send(400, {"error": "unsupported_grant_type"})

    def page(self, items, size, q, path):
        start = int(q.get("skip", ["0"])[0])
        chunk = items[start:start + size]
        out = {"value": chunk}
        if start + size < len(items):
            host = "https://evil.example.com" if MOCK.foreign_next else MOCK.base
            out["@odata.nextLink"] = f"{host}/v1.0{path}?skip={start + size}"
        return out

    def graph(self, method, path, q, body):
        if path == "/me":
            return self.send(200, MOCK.me)
        if path == "/me/calendarView":
            return self.send(200, self.page(MOCK.events, 2, q, path))
        if path == "/me/presence":
            return self.send(200, MOCK.presence)
        if path in ("/me/presence/setUserPreferredPresence", "/me/presence/clearUserPreferredPresence",
                    "/me/presence/setStatusMessage"):
            return self.send(200, None)
        if path == "/me/people":
            return self.send(200, {"value": MOCK.people})
        if path == "/me/messages":
            return self.send(200, {"value": MOCK.mails})
        if path == "/me/onenote/pages" and method == "GET":
            if MOCK.too_many_sections:
                return self.send(400, {"error": {"code": "20266", "message": "The number of maximum sections is exceeded for this request."}})
            return self.send(200, self.page(MOCK.pages, 3, q, path))
        if path == "/me/onenote/sections":
            return self.send(200, self.page(MOCK.sections, 2, q, path))
        if path.startswith("/me/onenote/sections/") and path.endswith("/pages") and method == "GET":
            sid = unquote(path.split("/")[4])
            return self.send(200, {"value": [p for p in MOCK.pages if p.get("_section") == sid]})
        if path.endswith("/pages") and method == "POST":
            MOCK.created.append((path, body, self.headers.get("Content-Type")))
            return self.send(201, {"id": "new-page", "title": "created",
                                   "links": {"oneNoteClientUrl": {"href": "onenote:https://d.docs.live.net/new"},
                                             "oneNoteWebUrl": {"href": "https://onedrive.live.com/new"}},
                                   "parentSection": {"displayName": "Quick Notes"}})
        self.send(404, {"error": {"code": "NotFound", "message": path}})


SERVER = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
MOCK.base = f"http://127.0.0.1:{SERVER.server_address[1]}"
MOCK.reset()
threading.Thread(target=SERVER.serve_forever, daemon=True).start()

# curl wrapper that records its arguments, so we can prove secrets never reach argv
TOOLS = tempfile.mkdtemp(prefix="m365-tools-")
ARGV_LOG = os.path.join(TOOLS, "argv.log")
CURL_WRAP = os.path.join(TOOLS, "curl")
with open(CURL_WRAP, "w") as f:
    f.write('#!/bin/bash\nprintf "%s\\n" "$*" >> "$M365_ARGV_LOG"\nexec /usr/bin/curl "$@"\n')
os.chmod(CURL_WRAP, 0o755)


# ---------------------------------------------------------------- helpers

def validate(data):
    assert isinstance(data.get("items"), list)
    for it in data["items"]:
        assert isinstance(it.get("title"), str) and it["title"], it
        assert os.path.exists(os.path.join(SRC, it["icon"]["path"])), it["icon"]
        if it.get("valid", True) is not False:
            assert "arg" in it, it
        for m in (it.get("mods") or {}).values():
            assert "subtitle" in m and "arg" in m, m


class Base(unittest.TestCase):
    def setUp(self):
        MOCK.reset()
        self.dir = tempfile.mkdtemp(prefix="m365-test-")
        for d in ("cache", "data", "kc"):
            os.makedirs(os.path.join(self.dir, d))
        self.env = dict(os.environ, alfred_workflow_cache=os.path.join(self.dir, "cache"),
                        alfred_workflow_data=os.path.join(self.dir, "data"),
                        alfred_workflow_bundleid="io.github.x-o-r-r-o.microsoft-365",
                        M365_TEST_KEYCHAIN_DIR=os.path.join(self.dir, "kc"),
                        M365_LOGIN_BASE=MOCK.base, M365_GRAPH_BASE=MOCK.base + "/v1.0",
                        M365_TEST_CURL=CURL_WRAP, M365_ARGV_LOG=ARGV_LOG,
                        M365_TEST_OPEN_FILE=os.path.join(self.dir, "opened"),
                        M365_TEST_NOTIFY_FILE=os.path.join(self.dir, "notified"),
                        M365_TEST_CLIPBOARD_OUT=os.path.join(self.dir, "copied"),
                        M365_TEST_CLIPBOARD="", M365_TEST_BG_SYNC="1", M365_TEST_APPS="",
                        M365_TEST_NOW="2026-09-26T10:00:00Z", M365_LOCALE="en_GB", TZ="UTC",
                        client_id=CLIENT, tenant="common")
        if os.path.exists(ARGV_LOG):
            os.remove(ARGV_LOG)

    def tearDown(self):
        # Tokens, refresh tokens and device codes must never appear in any process's arguments.
        if os.path.exists(ARGV_LOG):
            with open(ARGV_LOG) as f:
                argv = f.read()
            for secret in ("AT-", "RT-", "DEVCODE", "Bearer", "refresh_token"):
                self.assertNotIn(secret, argv)
        shutil.rmtree(self.dir, ignore_errors=True)

    def run_js(self, *args, **env):
        e = dict(self.env, **env)
        out = subprocess.run(["osascript", "-l", "JavaScript", "./m365.js", *args], cwd=SRC, env=e,
                             capture_output=True, text=True, timeout=60)
        self.assertEqual(out.returncode, 0, out.stderr)
        return out.stdout.rstrip("\n")

    def sf(self, cmd, query="", **env):
        data = json.loads(self.run_js(cmd, query, **env))
        validate(data)
        self.last = data
        return data["items"]

    def act(self, action, arg="", **vars):
        return self.run_js("act", arg, m365_action=action, **vars)

    def item_act(self, item, mod=None, **env):
        src = item["mods"][mod] if mod else item
        vars = dict(src.get("variables") or {})
        return self.run_js("act", src["arg"], **vars, **env)

    def read(self, name):
        p = os.path.join(self.dir, name)
        if not os.path.exists(p):
            return ""
        with open(p) as f:
            return f.read()

    def tokens(self):
        p = os.path.join(self.dir, "kc", KC_FILE)
        if not os.path.exists(p):
            return None
        with open(p) as f:
            return json.load(f)

    def sign_in(self, access="AT-seed", refresh="RT-seed", expires_in=3600, valid=True):
        with open(os.path.join(self.dir, "kc", KC_FILE), "w") as f:
            json.dump({"access_token": access, "refresh_token": refresh,
                       "expires_at": int(time.time() * 1000) + expires_in * 1000}, f)
        if valid:
            MOCK.valid_access.add(access)
        MOCK.valid_refresh.add(refresh)
        with open(os.path.join(self.dir, "data", "account.json"), "w") as f:
            json.dump({"id": "user-1", "name": "Zoë Tester", "email": "zoe@contoso.com",
                       "client_id": CLIENT, "tenant": "common"}, f)

    def wait_for(self, fn, timeout=20):
        end = time.time() + timeout
        while time.time() < end:
            v = fn()
            if v:
                return v
            time.sleep(0.05)
        self.fail("timed out waiting")

    def titles(self, items):
        return [i["title"] for i in items]


def age(cache, seconds):
    with open(cache) as f:
        c = json.load(f)
    c["fetched_at"] -= seconds * 1000
    with open(cache, "w") as f:
        json.dump(c, f)


def ev(subject, start, end, join="https://teams.microsoft.com/l/meetup-join/19%3ameeting_x%40thread.v2/0?context=%7b%7d",
       all_day=False, provider="teamsForBusiness", **kw):
    e = {"id": subject, "subject": subject, "isAllDay": all_day, "isCancelled": False,
         "start": {"dateTime": start, "timeZone": "UTC"}, "end": {"dateTime": end, "timeZone": "UTC"},
         "onlineMeeting": {"joinUrl": join} if join else None, "onlineMeetingProvider": provider if join else "unknown",
         "webLink": f"https://outlook.office365.com/owa/?itemid={subject.replace(' ', '')}",
         "organizer": {"emailAddress": {"name": "Anna"}}, "responseStatus": {"response": "accepted"},
         "location": {"displayName": "Room 1"}}
    e.update(kw)
    return e


def page(title, section="Notes", notebook="Work", modified="2026-09-20T10:00:00Z", **kw):
    p = {"id": title, "title": title, "lastModifiedDateTime": modified,
         "links": {"oneNoteClientUrl": {"href": f"onenote:https://d.docs.live.net/{title}"},
                   "oneNoteWebUrl": {"href": f"https://onedrive.live.com/{title}"}},
         "parentSection": {"id": section, "displayName": section}, "parentNotebook": {"displayName": notebook}}
    p.update(kw)
    return p


# ---------------------------------------------------------------- configuration and sign-in

class ConfigTests(Base):
    def test_missing_and_invalid_config(self):
        it = self.sf("teams", client_id="")
        self.assertIn("client) ID", it[0]["title"])
        self.assertEqual(it[0]["variables"]["m365_action"], "open")
        self.assertIn("must look like", self.sf("outlook", client_id="not-a-guid")[0]["title"])
        self.assertIn("tenant", self.sf("onenote", tenant="../evil")[0]["title"])
        self.assertIn("tenant", self.act("login", tenant="a/b"))

    def test_not_signed_in(self):
        for cmd in ("teams", "onenote", "outlook"):
            it = self.sf(cmd, "anything")
            self.assertEqual(it[0]["title"], "Sign in to Microsoft 365")
            self.assertEqual(it[0]["variables"]["m365_action"], "login")
        self.assertEqual(self.sf("account")[0]["title"], "Sign in to Microsoft 365")

    def test_unicode_queries_produce_valid_json(self):
        self.sign_in()
        for q in ['"quoted"', "new\nline", "émoji 🎉", "back\\slash", "'; rm -rf /", "$(id)"]:
            for cmd in ("teams", "onenote", "outlook", "account"):
                self.sf(cmd, q)


class LoginTests(Base):
    def start_login(self, scale="0.02", **env):
        out = self.act("login", M365_TEST_INTERVAL_SCALE=scale, **env)
        return out

    def test_device_code_flow_with_slow_down(self):
        MOCK.device_queue = ["pending", "slow_down", "pending", "ok"]
        out = self.start_login()
        self.assertEqual(out, "Code ABCD-EFGH copied: paste it in the browser to sign in")
        self.assertEqual(self.read("copied"), "ABCD-EFGH")
        self.assertEqual(self.read("opened").strip(), f"{MOCK.base}/devicelogin")
        pending = json.loads(self.read("data/login-pending.json"))
        self.assertNotIn("DEVCODE", json.dumps(pending))
        # While waiting, the account keyword shows the code.
        items = self.sf("account")
        self.assertEqual(items[0]["title"], "Enter code ABCD-EFGH to sign in")
        self.assertEqual(self.item_act(items[0]), "Code ABCD-EFGH copied")
        self.wait_for(lambda: "Signed in as Zoë Tester" in self.read("notified"))
        t = self.tokens()
        self.assertTrue(t["access_token"].startswith("AT-") and t["refresh_token"].startswith("RT-"))
        self.assertFalse(os.path.exists(os.path.join(self.dir, "data", "login-pending.json")))
        acct = json.loads(self.read("data/account.json"))
        self.assertEqual(acct["email"], "zoe@contoso.com")
        polls = MOCK.device_polls
        self.assertEqual(len(polls), 4)
        self.assertTrue(all(p[1] == "DEVCODE-secret-123" for p in polls))
        gaps = [b[0] - a[0] for a, b in zip(polls, polls[1:])]
        # interval 5 s (scaled to 0.1 s); after slow_down it becomes 10 s (0.2 s) for later polls
        self.assertGreaterEqual(gaps[0], 0.09)
        self.assertGreaterEqual(gaps[1], 0.19)
        self.assertGreaterEqual(gaps[2], 0.19)
        self.assertIn("Signed in as Zoë Tester", self.titles(self.sf("account"))[0])

    def test_declined_expired_and_bad_code(self):
        for resp, msg in (("declined", "Sign-in was declined."), ("expired", "The sign-in code expired"),
                          ("bad", "didn't recognize the sign-in code")):
            MOCK.reset()
            MOCK.device_queue = ["pending", resp]
            if os.path.exists(os.path.join(self.dir, "notified")):
                os.remove(os.path.join(self.dir, "notified"))
            self.start_login()
            self.wait_for(lambda: msg in self.read("notified"))
            self.assertIsNone(self.tokens())
            self.wait_for(lambda: not os.path.exists(os.path.join(self.dir, "data", "login-pending.json")))

    def test_local_expiry_stops_polling(self):
        MOCK.devicecode = dict(MOCK.devicecode, expires_in=1, interval=1)
        self.start_login(scale="1.2")
        self.wait_for(lambda: "expired" in self.read("notified"))
        self.assertEqual(MOCK.device_polls, [])

    def test_cancel_stops_the_poller(self):
        self.start_login()
        self.wait_for(lambda: len(MOCK.device_polls) >= 2)
        self.assertEqual(self.act("cancel-login"), "Sign-in cancelled")
        time.sleep(0.4)
        n = len(MOCK.device_polls)
        time.sleep(0.5)
        self.assertEqual(len(MOCK.device_polls), n)
        self.assertEqual(self.sf("account")[0]["title"], "Sign in to Microsoft 365")

    def test_new_login_supersedes_old_poller(self):
        self.start_login()
        self.wait_for(lambda: len(MOCK.device_polls) >= 1)
        MOCK.devicecode = dict(MOCK.devicecode, device_code="DEVCODE-second", user_code="WXYZ-1234")
        self.start_login()
        time.sleep(0.6)
        cut = time.time() - 0.3
        late = [c for t, c in MOCK.device_polls if t > cut]
        self.assertTrue(late)
        self.assertEqual(set(late), {"DEVCODE-second"})
        self.act("cancel-login")

    def test_aadsts_errors_are_explained(self):
        cases = {7000218: "Allow public client flows", 700016: "client ID", 50194: "single-tenant",
                 65001: "consent", 90094: "admin", 9002346: "consumers", 90002: "Tenant not found"}
        for code, text in cases.items():
            MOCK.devicecode_error = {"error": "invalid_request", "error_description": f"AADSTS{code}: Something.\r\nTrace ID: x",
                                     "error_codes": [code]}
            self.assertIn(text, self.act("login"))
        MOCK.devicecode_error = {"error": "invalid_request", "error_description": "AADSTS12345: Weird thing.\r\nTrace ID: x"}
        self.assertEqual(self.act("login"), "Weird thing.")

    def test_login_offline(self):
        self.assertIn("Can't reach Microsoft", self.act("login", M365_LOGIN_BASE="http://127.0.0.1:9"))

    def test_logout(self):
        self.sign_in()
        MOCK.events = [ev("Standup", "2026-09-26T11:00:00.0000000", "2026-09-26T11:15:00.0000000")]
        self.sf("teams")
        self.assertTrue(os.path.isdir(os.path.join(self.dir, "cache", "user-1")))
        self.assertEqual(self.act("logout"), "Signed out of Microsoft 365")
        self.assertIsNone(self.tokens())
        self.assertFalse(os.path.exists(os.path.join(self.dir, "cache", "user-1")))


class TokenTests(Base):
    def setUp(self):
        super().setUp()
        MOCK.events = [ev("Standup", "2026-09-26T11:00:00.0000000", "2026-09-26T11:15:00.0000000")]

    def test_refresh_when_expired(self):
        self.sign_in(access="AT-old", expires_in=-10, valid=False)
        self.assertEqual(self.sf("teams")[0]["title"], "Standup")
        t = self.tokens()
        self.assertEqual(t["access_token"], "AT-1")
        self.assertEqual(t["refresh_token"], "RT-1")  # rotated

    def test_401_refreshes_once_and_retries(self):
        self.sign_in(access="AT-revoked", valid=False)
        self.assertEqual(self.sf("teams")[0]["title"], "Standup")
        self.assertEqual(self.tokens()["access_token"], "AT-1")

    def test_refresh_keeps_old_refresh_token_when_none_returned(self):
        self.sign_in(access="AT-old", expires_in=-10, valid=False)
        MOCK.overrides[("POST", "/common/oauth2/v2.0/token")] = [(200, {}, {"access_token": "AT-x", "expires_in": 3600})]
        MOCK.valid_access.add("AT-x")
        self.sf("teams")
        self.assertEqual(self.tokens()["refresh_token"], "RT-seed")

    def test_invalid_grant_asks_to_sign_in_again(self):
        self.sign_in(access="AT-old", refresh="RT-dead", expires_in=-10, valid=False)
        MOCK.valid_refresh.clear()
        it = self.sf("teams")
        self.assertEqual(it[0]["title"], "Sign in to Microsoft 365")
        self.assertIn("expired", it[0]["subtitle"])
        self.assertIsNone(self.tokens())

    def test_consent_required_on_refresh(self):
        self.sign_in(access="AT-old", expires_in=-10, valid=False)
        MOCK.refresh_error = {"error": "invalid_grant", "error_description": "AADSTS65001: The user or administrator has not consented.",
                              "error_codes": [65001]}
        it = self.sf("outlook")
        self.assertIn("consent", it[0]["title"])
        self.assertEqual(it[1]["title"], "Sign in again")

    def test_403_permission_denied(self):
        self.sign_in()
        MOCK.overrides[("GET", "/v1.0/me/messages")] = [(403, {}, {"error": {"code": "ErrorAccessDenied", "message": "Access is denied."}})]
        it = self.sf("outlook", "invoice")
        self.assertEqual(it[0]["title"], "Permission denied: Access is denied.")
        self.assertIn("admin consent", it[0]["subtitle"])

    def test_429_short_retry_after_is_retried(self):
        self.sign_in()
        MOCK.overrides[("GET", "/v1.0/me/messages")] = [(429, {"Retry-After": "1"}, {"error": {"code": "TooManyRequests", "message": "slow"}})]
        MOCK.mails = [{"subject": "Hi", "from": {"emailAddress": {"name": "A", "address": "a@x.com"}}, "receivedDateTime": "2026-09-26T09:00:00Z",
                       "webLink": "https://outlook.office365.com/owa/?ItemID=1", "isRead": True}]
        t = time.time()
        self.assertEqual(self.sf("outlook", "hi")[0]["title"], "Hi")
        self.assertGreaterEqual(time.time() - t, 0.9)

    def test_429_long_retry_after_backs_off(self):
        self.sign_in()
        MOCK.overrides[("GET", "/v1.0/me/messages")] = [(429, {"Retry-After": "120"}, {"error": {"code": "TooManyRequests", "message": "slow"}})]
        it = self.sf("outlook", "hi")
        self.assertEqual(it[0]["title"], "Microsoft 365 is busy")
        self.assertIn("120 seconds", it[0]["subtitle"])
        n = len(MOCK.requests)
        it = self.sf("outlook", "other")
        self.assertEqual(it[0]["title"], "Microsoft 365 is busy")
        self.assertEqual(len(MOCK.requests), n)  # didn't hammer the service

    def test_offline(self):
        self.sign_in()
        it = self.sf("teams", M365_GRAPH_BASE="http://127.0.0.1:9/v1.0")
        self.assertEqual(it[0]["title"], "You're offline")

    def test_offline_with_stale_cache_shows_cached_results(self):
        self.sign_in()
        self.sf("teams")
        cache = os.path.join(self.dir, "cache", "user-1", "events-2026-09-26.json")
        age(cache, 3600)
        it = self.sf("teams", M365_GRAPH_BASE="http://127.0.0.1:9/v1.0")
        self.assertTrue(it[0]["title"].startswith("Offline: showing results from 1 h ago"), it[0])
        self.assertEqual(it[1]["title"], "Standup")

    def test_foreign_next_link_is_refused(self):
        self.sign_in()
        MOCK.events = [ev(f"M{i}", "2026-09-26T11:00:00", "2026-09-26T11:15:00") for i in range(3)]
        MOCK.foreign_next = True
        it = self.sf("teams")
        self.assertIn("outside Microsoft Graph", it[0]["title"])


# ---------------------------------------------------------------- Teams

class TeamsTests(Base):
    def setUp(self):
        super().setUp()
        self.sign_in()

    def test_meetings_order_filter_and_pagination(self):
        MOCK.events = [
            ev("Early sync", "2026-09-26T08:00:00.0000000", "2026-09-26T08:30:00.0000000"),
            ev("Design review", "2026-09-26T09:30:00.1234567", "2026-09-26T10:30:00.0000000"),
            ev("Planning", "2026-09-26T13:00:00.0000000", "2026-09-26T14:00:00.0000000"),
            ev("In person", "2026-09-26T15:00:00.0000000", "2026-09-26T16:00:00.0000000", join=None),
            ev("Zoom call", "2026-09-26T15:00:00.0000000", "2026-09-26T16:00:00.0000000", join="https://zoom.us/j/1", provider="unknown"),
            ev("Declined", "2026-09-26T16:00:00.0000000", "2026-09-26T17:00:00.0000000", responseStatus={"response": "declined"}),
            ev("Cancelled", "2026-09-26T16:00:00.0000000", "2026-09-26T17:00:00.0000000", isCancelled=True),
        ]
        it = self.sf("teams")
        self.assertEqual(self.titles(it), ["Design review", "Planning", "Early sync", "Set your status…"])
        self.assertTrue(it[0]["subtitle"].startswith("Now · 09:30–10:30"), it[0]["subtitle"])
        self.assertTrue(it[1]["subtitle"].startswith("In 3 h · 13:00–14:00"))
        self.assertTrue(it[2]["subtitle"].startswith("Ended"))
        # every page was requested with the UTC Prefer header (4 pages of 2)
        reqs = MOCK.graph_requests("/v1.0/me/calendarView")
        self.assertEqual(len(reqs), 4)
        self.assertTrue(all(r[3].get("Prefer") == 'outlook.timezone="UTC"' for r in reqs))
        q = reqs[0][2]
        self.assertEqual(q["startDateTime"], ["2026-09-26T00:00:00.000Z"])
        self.assertEqual(q["endDateTime"], ["2026-09-27T00:00:00.000Z"])
        self.assertEqual(self.titles(self.sf("teams", "plan")), ["Planning"])

    def test_join_prefers_app_then_web(self):
        MOCK.events = [ev("Standup", "2026-09-26T11:00:00", "2026-09-26T11:15:00")]
        item = self.sf("teams")[0]
        self.item_act(item)  # Teams not installed -> https
        self.item_act(item, M365_TEST_APPS="msteams")
        self.item_act(item, M365_TEST_APPS="msteams", teams_open="web")
        opened = self.read("opened").splitlines()
        self.assertTrue(opened[0].startswith("https://teams.microsoft.com/l/meetup-join/"))
        self.assertTrue(opened[1].startswith("msteams:/l/meetup-join/19%3ameeting_x"))
        self.assertTrue(opened[2].startswith("https://"))
        self.assertEqual(self.item_act(item, "cmd"), "Copied https://teams.microsoft.com/l/meetup-join/19%3ameeting_x%40…")
        self.item_act(item, "alt")
        self.assertIn("outlook.office365.com", self.read("opened").splitlines()[-1])

    def test_time_zones(self):
        MOCK.events = [ev("Standup", "2026-09-26T03:30:00.0000000", "2026-09-26T04:00:00.0000000")]
        it = self.sf("teams", TZ="Asia/Kolkata", M365_TEST_NOW="2026-09-26T02:00:00Z")
        self.assertIn("09:00–09:30", it[0]["subtitle"])
        q = MOCK.graph_requests("/v1.0/me/calendarView")[0][2]
        self.assertEqual(q["startDateTime"], ["2026-09-25T18:30:00.000Z"])  # local midnight in UTC

    def test_dst_day_window(self):
        # 2026-11-01 is 25 hours long in New York
        self.sf("teams", TZ="America/New_York", M365_TEST_NOW="2026-11-01T15:00:00Z")
        q = MOCK.graph_requests("/v1.0/me/calendarView")[0][2]
        self.assertEqual(q["startDateTime"], ["2026-11-01T04:00:00.000Z"])
        self.assertEqual(q["endDateTime"], ["2026-11-02T05:00:00.000Z"])

    def test_all_day_events_use_calendar_dates(self):
        MOCK.events = [
            ev("Yesterday", "2026-09-25T00:00:00.0000000", "2026-09-26T00:00:00.0000000", join=None, all_day=True),
            ev("Today", "2026-09-26T00:00:00.0000000", "2026-09-27T00:00:00.0000000", join=None, all_day=True),
            ev("Tomorrow", "2026-09-27T00:00:00.0000000", "2026-09-28T00:00:00.0000000", join=None, all_day=True),
            ev("Trip", "2026-09-24T00:00:00.0000000", "2026-09-29T00:00:00.0000000", join=None, all_day=True),
        ]
        for tz, now in (("America/Los_Angeles", "2026-09-26T20:00:00Z"), ("Asia/Tokyo", "2026-09-26T01:00:00Z")):
            shutil.rmtree(os.path.join(self.dir, "cache", "user-1"), ignore_errors=True)
            it = self.sf("outlook", TZ=tz, M365_TEST_NOW=now)
            self.assertEqual(self.titles(it)[:2], ["Today", "Trip"], tz)
            self.assertTrue(it[0]["subtitle"].startswith("All day"))

    def test_no_meetings(self):
        it = self.sf("teams")
        self.assertEqual(self.titles(it), ["No Teams meetings today", "Set your status…"])
        self.assertEqual(it[1]["autocomplete"], "status ")

    def test_status_items_and_actions(self):
        MOCK.presence = {"availability": "DoNotDisturb", "activity": "Presenting"}
        it = self.sf("teams", "status")
        self.assertEqual(it[0]["title"], "Current status: Do not disturb")
        self.assertEqual(it[0]["subtitle"], "Presenting")
        self.assertEqual(len(it), 9)
        self.assertEqual(it[1]["autocomplete"], "message ")
        it = self.sf("teams", "status busy 2h")
        self.assertEqual(self.titles(it)[1:], ["Busy"])
        self.assertEqual(it[1]["subtitle"], "Set for 2 hours")
        self.assertEqual(self.item_act(it[1]), "Status set to Busy for 2 hours")
        body = json.loads(MOCK.graph_requests("/v1.0/me/presence/setUserPreferredPresence")[-1][4])
        self.assertEqual(body, {"availability": "Busy", "activity": "Busy", "expirationDuration": "PT2H"})
        it = self.sf("teams", "status dnd 1h 30m")
        self.assertEqual(self.titles(it)[1:], ["Do not disturb"])
        self.item_act(it[1])
        body = json.loads(MOCK.graph_requests("/v1.0/me/presence/setUserPreferredPresence")[-1][4])
        self.assertEqual(body["expirationDuration"], "PT1H30M")
        it = self.sf("teams", "status offline")
        self.assertEqual(self.item_act(it[1]), "Status set to Appear offline (expires in 7 days)")
        body = json.loads(MOCK.graph_requests("/v1.0/me/presence/setUserPreferredPresence")[-1][4])
        self.assertEqual(body, {"availability": "Offline", "activity": "OffWork"})
        it = self.sf("teams", "status reset")
        self.assertEqual(self.item_act(it[1]), "Status reset: Teams sets it automatically")
        self.assertEqual(len(MOCK.graph_requests("/v1.0/me/presence/clearUserPreferredPresence")), 1)
        self.assertEqual(self.sf("teams", "status 3d")[1]["subtitle"], "Set for 3 days")
        self.assertEqual(self.sf("teams", "status nope")[-1]["title"], "Unknown status")

    def test_presence_on_personal_account(self):
        MOCK.overrides[("GET", "/v1.0/me/presence")] = [(403, {}, {"error": {"code": "Forbidden", "message": "Not supported for MSA"}})]
        MOCK.overrides[("POST", "/v1.0/me/presence/setUserPreferredPresence")] = [(403, {}, {"error": {"code": "Forbidden", "message": "x"}})]
        it = self.sf("teams", "status busy")
        self.assertEqual(it[0]["title"], "Couldn't read your status")
        self.assertIn("work or school", it[0]["subtitle"])
        self.assertIn("work or school", self.item_act(it[1]))

    def test_people_search_and_chat(self):
        MOCK.people = [
            {"displayName": "José Núñez", "scoredEmailAddresses": [{"address": "jose+x@contoso.com"}], "jobTitle": "PM", "personType": {"class": "Person"}},
            {"displayName": "Team DL", "scoredEmailAddresses": [{"address": "dl@contoso.com"}], "personType": {"class": "Group"}},
            {"displayName": "Dup", "scoredEmailAddresses": [{"address": "JOSE+x@contoso.com"}], "personType": {"class": "Person"}},
        ]
        it = self.sf("teams", 'jo"sé')
        self.assertEqual(self.titles(it), ["José Núñez"])
        search = MOCK.graph_requests("/v1.0/me/people")[-1][2]["$search"][0]
        self.assertEqual(search, '"jo\\"sé"')
        self.item_act(it[0])
        self.assertEqual(self.read("opened").strip(), "https://teams.microsoft.com/l/chat/0/0?users=jose%2Bx%40contoso.com")
        self.item_act(it[0], M365_TEST_APPS="msteams")
        self.assertEqual(self.read("opened").splitlines()[-1], "msteams:/l/chat/0/0?users=jose%2Bx%40contoso.com")
        self.item_act(it[0], "alt", M365_TEST_APPS="msteams")
        self.assertEqual(self.read("opened").splitlines()[-1], "msteams:/l/call/0/0?users=jose%2Bx%40contoso.com&withVideo=true")
        self.assertEqual(self.item_act(it[0], "cmd"), "Copied jose+x@contoso.com")
        self.assertEqual(self.read("copied"), "jose+x@contoso.com")

    def test_people_search_error_keeps_meetings(self):
        MOCK.events = [ev("Josh sync", "2026-09-26T11:00:00", "2026-09-26T11:15:00")]
        MOCK.overrides[("GET", "/v1.0/me/people")] = [(403, {}, {"error": {"code": "Forbidden", "message": "People.Read missing"}})]
        it = self.sf("teams", "josh")
        self.assertEqual(it[0]["title"], "Josh sync")
        self.assertIn("People.Read missing", it[1]["title"])


# ---------------------------------------------------------------- Outlook

class OutlookTests(Base):
    def setUp(self):
        super().setUp()
        self.sign_in()

    def test_agenda(self):
        MOCK.events = [
            ev("Lunch", "2026-09-26T12:00:00", "2026-09-26T13:00:00", join=None),
            ev("Offsite", "2026-09-26T00:00:00", "2026-09-27T00:00:00", join=None, all_day=True),
            ev("Night shift", "2026-09-26T22:00:00", "2026-09-27T06:00:00", join=None),
            ev("Standup", "2026-09-26T11:00:00", "2026-09-26T11:15:00"),
        ]
        it = self.sf("outlook")
        self.assertEqual(self.titles(it), ["Offsite", "Standup", "Lunch", "Night shift", "Type to search your mail"])
        self.assertIn("Sep 27 06:00", it[3]["subtitle"].replace("27 Sept", "Sep 27").replace("27 Sep", "Sep 27"))
        self.assertEqual(it[2]["mods"]["cmd"]["valid"], False)
        self.item_act(it[1], "cmd")
        self.assertTrue(self.read("opened").startswith("https://teams.microsoft.com/l/meetup-join/"))
        self.item_act(it[2])
        self.assertIn("itemid=Lunch", self.read("opened").splitlines()[-1])

    def test_empty_agenda(self):
        self.assertEqual(self.sf("outlook")[0]["title"], "Nothing on your calendar today")

    def test_mail_search(self):
        MOCK.mails = [
            {"subject": "Invoice “Q3” ✓", "from": {"emailAddress": {"name": "Ana", "address": "ana@x.com"}},
             "receivedDateTime": "2026-09-26T08:15:00Z", "bodyPreview": "Hello\nthere", "isRead": False, "hasAttachments": True,
             "webLink": "https://outlook.office365.com/owa/?ItemID=AAA%2B&exvsurl=1"},
            {"subject": "", "from": None, "receivedDateTime": "2025-01-02T08:15:00Z", "isRead": True,
             "webLink": "javascript:alert(1)"},
        ]
        it = self.sf("outlook", 'invoice "q3"')
        self.assertEqual(it[0]["title"], "● Invoice “Q3” ✓ 📎")
        self.assertEqual(it[0]["subtitle"], "Ana · 08:15 · Hello there")
        self.assertEqual(it[0]["mods"]["cmd"]["arg"], "ana@x.com")
        self.assertEqual(it[1]["title"], "(No subject)")
        self.assertIs(it[1]["valid"], False)  # unsafe link isn't offered
        self.assertIn("2025", it[1]["subtitle"])
        q = MOCK.graph_requests("/v1.0/me/messages")[-1][2]
        self.assertEqual(q["$search"], ['"invoice \\"q3\\""'])
        self.item_act(it[0])
        self.assertEqual(self.read("opened").strip(), "https://outlook.office365.com/owa/?ItemID=AAA%2B&exvsurl=1")

    def test_no_mail(self):
        self.assertTrue(self.sf("outlook", "zzz")[0]["title"].startswith("No mail matches"))


# ---------------------------------------------------------------- OneNote

class OneNoteTests(Base):
    def setUp(self):
        super().setUp()
        self.sign_in()
        MOCK.pages = [
            page("Café notes", modified="2026-09-25T10:00:00Z"),
            page("Meeting notes", modified="2026-09-24T10:00:00Z"),
            page("Recipes", section="Home", notebook="Personal", modified="2026-09-23T10:00:00Z"),
            page("Weekly café review", modified="2026-09-22T10:00:00Z"),
            page("C++ tips (draft)", modified="2026-09-21T10:00:00Z"),
            page("", modified="2026-09-20T10:00:00Z"),
            page("日本語のメモ", modified="2026-09-19T10:00:00Z"),
        ]

    def test_index_search_and_open(self):
        it = self.sf("onenote")
        self.assertEqual(len(MOCK.graph_requests("/v1.0/me/onenote/pages")), 3)  # 7 pages, 3 per page
        self.assertEqual(it[0]["title"], "Café notes")
        self.assertRegex(it[0]["subtitle"], r"^Work › Notes · edited 25 Sept?$")
        self.assertIn("Untitled Page", self.titles(it))
        self.assertEqual(self.titles(it)[-2:], ["Create a new page…", "Rebuild the page index"])
        self.assertEqual(self.titles(self.sf("onenote", "CAFE"))[:2], ["Café notes", "Weekly café review"])
        self.assertEqual(self.titles(self.sf("onenote", "notes"))[:2], ["Café notes", "Meeting notes"])
        self.assertEqual(self.titles(self.sf("onenote", "personal"))[0], "Recipes")
        self.assertEqual(self.titles(self.sf("onenote", "c++ (dr"))[0], "C++ tips (draft)")
        self.assertEqual(self.titles(self.sf("onenote", "メモ"))[0], "日本語のメモ")
        self.assertTrue(self.sf("onenote", "zzz")[0]["title"].startswith("No page titles match"))
        item = self.sf("onenote", "recipes")[0]
        self.item_act(item)
        self.item_act(item, M365_TEST_APPS="onenote")
        self.item_act(item, M365_TEST_APPS="onenote", onenote_open="web")
        self.item_act(item, "cmd")
        self.assertEqual(self.read("opened").splitlines(), [
            "https://onedrive.live.com/Recipes", "onenote:https://d.docs.live.net/Recipes",
            "https://onedrive.live.com/Recipes", "https://onedrive.live.com/Recipes"])

    def test_index_is_cached(self):
        self.sf("onenote")
        n = len(MOCK.requests)
        self.sf("onenote", "cafe")
        self.assertEqual(len(MOCK.requests), n)

    def test_too_many_sections_falls_back_to_sections(self):
        MOCK.too_many_sections = True
        MOCK.sections = [{"id": "s1", "displayName": "Notes", "parentNotebook": {"displayName": "Work"}},
                         {"id": "s/2", "displayName": "Home", "parentNotebook": {"displayName": "Personal"}},
                         {"id": "s3", "displayName": "Empty", "parentNotebook": {"displayName": "Work"}}]
        for p in MOCK.pages:
            p["_section"] = "s/2" if p["title"] == "Recipes" else "s1"
            del p["parentSection"], p["parentNotebook"]
        it = self.sf("onenote", "recipes")
        self.assertEqual(it[0]["title"], "Recipes")
        self.assertTrue(it[0]["subtitle"].startswith("Personal › Home"))
        self.assertEqual(self.sf("onenote")[0]["title"], "Café notes")  # sorted by modified date
        self.assertTrue(MOCK.graph_requests("/v1.0/me/onenote/sections/s%2F2/pages"))

    def test_background_indexing_without_cache(self):
        env = {"M365_TEST_BG_SYNC": ""}
        it = self.sf("onenote", **env)
        self.assertEqual(it[0]["title"], "Indexing your OneNote pages…")
        self.assertEqual(self.last["rerun"], 1)
        self.wait_for(lambda: os.path.exists(os.path.join(self.dir, "cache", "user-1", "onenote-index.json")))
        self.wait_for(lambda: self.sf("onenote", "recipes", **env)[0]["title"] == "Recipes")

    def test_stale_index_refreshes_in_background(self):
        self.sf("onenote")
        cache = os.path.join(self.dir, "cache", "user-1", "onenote-index.json")
        age(cache, 7200)
        MOCK.pages.append(page("Brand new", modified="2026-09-26T09:00:00Z"))
        env = {"M365_TEST_BG_SYNC": ""}
        it = self.sf("onenote", "brand", **env)
        self.assertTrue(it[0]["title"].startswith("No page titles"))
        self.assertEqual(self.last["rerun"], 0.5)
        self.wait_for(lambda: self.sf("onenote", "brand", **env)[0]["title"] == "Brand new")

    def test_index_error(self):
        MOCK.overrides[("GET", "/v1.0/me/onenote/pages")] = [(500, {}, {"error": {"code": "19999", "message": "Boom"}})]
        it = self.sf("onenote")
        self.assertIn("Boom", it[0]["title"])

    def test_create_page_from_clipboard(self):
        self.sf("onenote")
        it = self.sf("onenote", "new Ideas & <plans>", M365_TEST_CLIPBOARD="Line 1 <b>\nLine 2 é\x01\n\nPara 2 🎉")
        self.assertEqual(it[0]["title"], "Create “Ideas & <plans>”")
        self.assertIn("clipboard (30 characters)", it[0]["subtitle"])
        out = self.item_act(it[0], M365_TEST_CLIPBOARD="Line 1 <b>\nLine 2 é\x01\n\nPara 2 🎉")
        self.assertEqual(out, "Created “Ideas & <plans>”")
        path, body, ctype = MOCK.created[-1]
        self.assertEqual(path, "/me/onenote/pages")
        self.assertEqual(ctype, "text/html; charset=utf-8")
        self.assertIn("<title>Ideas &amp; &lt;plans&gt;</title>", body)
        self.assertIn("<p>Line 1 &lt;b&gt;<br/>Line 2 é</p>\n<p>Para 2 🎉</p>", body)
        self.assertIn('<meta name="created" content="2026-09-26T10:00:00+00:00" />', body)
        # the new page is searchable immediately
        self.assertEqual(self.sf("onenote", "created")[0]["title"], "created")

    def test_create_page_inline_and_empty(self):
        it = self.sf("onenote", "new Todo :: buy milk")
        self.assertEqual(it[0]["title"], "Create “Todo”")
        self.item_act(it[0], M365_TEST_CLIPBOARD="ignored")
        self.assertIn("<p>buy milk</p>", MOCK.created[-1][1])
        self.item_act(it[0], "alt")
        self.assertIn("<body>\n\n</body>", MOCK.created[-1][1])
        self.item_act(it[0], "cmd", M365_TEST_APPS="onenote")
        self.assertEqual(self.read("opened").strip(), "onenote:https://d.docs.live.net/new")
        self.assertEqual(self.sf("onenote", "new")[0]["title"], "Type a title for the new page")

    def test_create_page_in_configured_section(self):
        MOCK.sections = [{"id": "sec-1", "displayName": "Inbox", "parentNotebook": {"displayName": "Work"}},
                         {"id": "sec-2", "displayName": "Inbox", "parentNotebook": {"displayName": "Home"}},
                         {"id": "sec-3", "displayName": "Ideas", "parentNotebook": {"displayName": "Home"}}]
        it = self.sf("onenote", "new T")
        self.item_act(it[0], onenote_section="home / inbox")
        self.assertEqual(MOCK.created[-1][0], "/me/onenote/sections/sec-2/pages")
        self.item_act(it[0], onenote_section="Ideas")
        self.assertEqual(MOCK.created[-1][0], "/me/onenote/sections/sec-3/pages")
        self.assertIn("Several sections", self.item_act(it[0], onenote_section="Inbox"))
        self.item_act(it[0], onenote_section="New Stuff")
        self.assertEqual(MOCK.requests[-1][2]["sectionName"], ["New Stuff"])
        self.assertIn("not found", self.item_act(it[0], onenote_section="Nope/Nope"))

    def test_create_page_errors(self):
        MOCK.overrides[("POST", "/v1.0/me/onenote/pages")] = [(507, {}, {"error": {"code": "19999", "message": "full"}})]
        it = self.sf("onenote", "new T")
        self.assertIn("section is full", self.item_act(it[0]))
        self.assertIn("Can't reach", self.item_act(it[0], M365_GRAPH_BASE="http://127.0.0.1:9/v1.0"))


class AccountTests(Base):
    def test_signed_in_rows(self):
        self.sign_in()
        it = self.sf("account")
        self.assertEqual(it[0]["title"], "Signed in as Zoë Tester (zoe@contoso.com)")
        self.assertEqual(self.titles(self.sf("account", "logout")), ["Sign out"])
        self.sf("onenote")
        self.assertEqual(self.act("clear-cache"), "Cleared cached data")


# ---------------------------------------------------------------- regression tests (audit pass 1)

class AuditOneTests(Base):
    def fake_curl(self, raw):
        path = os.path.join(self.dir, "fakecurl")
        with open(os.path.join(self.dir, "response"), "wb") as f:
            f.write(raw)
        with open(path, "w") as f:
            f.write(f'#!/bin/bash\ncat >/dev/null\ncat "{self.dir}/response"\n')
        os.chmod(path, 0o755)
        return path

    def test_proxy_and_interim_header_blocks(self):
        self.sign_in()
        body = json.dumps({"value": [{"subject": "Via proxy", "webLink": "https://outlook.office365.com/x", "receivedDateTime": "2026-09-26T08:00:00Z"}]})
        raw = ("HTTP/1.1 200 Connection established\r\n\r\nHTTP/1.1 100 Continue\r\n\r\n"
               "HTTP/2 200\r\ncontent-type: application/json\r\n\r\n" + body).encode()
        it = self.sf("outlook", "proxy", M365_TEST_CURL=self.fake_curl(raw))
        self.assertEqual(it[0]["title"], "Via proxy")

    def test_second_401_keeps_the_refresh_token(self):
        self.sign_in(access="AT-bad", valid=False)
        MOCK.overrides[("GET", "/v1.0/me/messages")] = [(401, {}, {"error": {"code": "InvalidAuthenticationToken", "message": "claims"}})] * 2
        it = self.sf("outlook", "x")
        self.assertEqual(it[0]["title"], "Sign in to Microsoft 365")
        self.assertIsNotNone(self.tokens())

    def test_stale_throttle_does_not_break_sign_in(self):
        with open(os.path.join(self.dir, "cache", "throttled-until"), "w") as f:
            f.write(str(int(time.time() * 1000) + 300000))
        MOCK.device_queue = ["ok"]
        self.act("login", M365_TEST_INTERVAL_SCALE="0.02")
        self.wait_for(lambda: "Signed in as Zoë Tester" in self.read("notified"))

    def test_expiry_is_noticed_before_the_next_interval(self):
        MOCK.devicecode = dict(MOCK.devicecode, expires_in=1, interval=30)
        t = time.time()
        self.act("login")
        self.wait_for(lambda: "expired" in self.read("notified"), timeout=10)
        self.assertLess(time.time() - t, 6)
        self.assertEqual(MOCK.device_polls, [])

    def test_keychain_failure_is_reported(self):
        MOCK.device_queue = ["ok"]
        os.chmod(os.path.join(self.dir, "kc"), 0o500)
        try:
            self.act("login", M365_TEST_INTERVAL_SCALE="0.02")
            self.wait_for(lambda: "Keychain" in self.read("notified"))
        finally:
            os.chmod(os.path.join(self.dir, "kc"), 0o700)

    def test_old_search_caches_are_pruned(self):
        self.sign_in()
        old = os.path.join(self.dir, "cache", "user-1", "mail-deadbeef.json")
        os.makedirs(os.path.dirname(old), exist_ok=True)
        with open(old, "w") as f:
            f.write("{}")
        os.utime(old, (time.time() - 2 * 86400, time.time() - 2 * 86400))
        self.sf("outlook", "fresh")
        self.assertFalse(os.path.exists(old))

    def test_onenote_10008(self):
        self.sign_in()
        MOCK.overrides[("GET", "/v1.0/me/onenote/pages")] = [(403, {}, {"error": {"code": "10008", "message": "too many items"}})]
        self.assertIn("5,000 OneNote items", self.sf("onenote")[0]["title"])

    def test_new_prefix_still_finds_pages(self):
        self.sign_in()
        MOCK.pages = [page("New ideas"), page("Other")]
        self.sf("onenote")
        it = self.sf("onenote", "new ideas")
        self.assertEqual(self.titles(it), ["Create “ideas”", "New ideas"])

    def test_tenant_case_does_not_lose_the_account(self):
        self.sign_in()
        self.assertIn("Zoë Tester", self.sf("account", tenant="Common")[0]["title"])

    def test_truncation_keeps_emoji_whole(self):
        self.sign_in()
        MOCK.events = [ev("a" * 119 + "🎉🎉", "2026-09-26T11:00:00", "2026-09-26T11:15:00")]
        title = self.sf("teams")[0]["title"]
        self.assertEqual(title, "a" * 119 + "…")
        MOCK.reset()
        MOCK.valid_access.add("AT-seed")
        shutil.rmtree(os.path.join(self.dir, "cache", "user-1"))
        MOCK.events = [ev("a" * 118 + "🎉🎉🎉", "2026-09-26T11:00:00", "2026-09-26T11:15:00")]
        self.assertEqual(self.sf("teams")[0]["title"], "a" * 118 + "🎉…")

    def test_one_letter_query_hint(self):
        self.sign_in()
        self.assertEqual(self.sf("teams", "j")[0]["title"], "Keep typing to find people")

    def test_logout_without_account_file_clears_cache(self):
        self.sign_in()
        os.remove(os.path.join(self.dir, "data", "account.json"))
        self.sf("teams")
        self.assertTrue(os.path.isdir(os.path.join(self.dir, "cache", "default")))
        self.act("logout")
        self.assertFalse(os.path.exists(os.path.join(self.dir, "cache", "default")))


# ---------------------------------------------------------------- regression tests (audit pass 2)

class AuditTwoTests(Base):
    def dead_lock(self, name):
        d = os.path.join(self.dir, "cache", "locks", f"{name}.lock")
        os.makedirs(d)
        with open(os.path.join(d, "pid"), "w") as f:
            f.write("999999")

    def test_lock_of_a_killed_process_is_ignored(self):
        self.sign_in()
        MOCK.pages = [page("Old")]
        self.sf("onenote")
        age(os.path.join(self.dir, "cache", "user-1", "onenote-index.json"), 7200)
        MOCK.pages = [page("Fresh")]
        self.dead_lock("refresh-onenote-index")
        it = self.sf("onenote")
        self.assertEqual(it[0]["title"], "Fresh")
        self.assertNotIn("rerun", self.last)

    def test_token_lock_of_a_killed_process_does_not_stall(self):
        self.sign_in(access="AT-old", expires_in=-10, valid=False)
        MOCK.events = [ev("Standup", "2026-09-26T11:00:00", "2026-09-26T11:15:00")]
        self.dead_lock("token")
        t = time.time()
        self.assertEqual(self.sf("teams")[0]["title"], "Standup")
        self.assertLess(time.time() - t, 5)

    def test_newlines_in_graph_fields_are_flattened(self):
        self.sign_in()
        MOCK.events = [ev("Two\nlines", "2026-09-26T11:00:00", "2026-09-26T11:15:00", join=None, webLink="",
                          location={"displayName": "Room\r\n2"})]
        it = self.sf("outlook")
        self.assertEqual(it[0]["title"], "Two lines")
        self.assertIn("Room 2", it[0]["subtitle"])

    def test_personal_accounts_skip_presence(self):
        self.act("login", tenant="consumers")
        self.assertNotIn("Presence", MOCK.device_scopes[-1])
        self.act("cancel-login", tenant="consumers")
        self.act("login")
        self.assertIn("Presence.ReadWrite", MOCK.device_scopes[-1])
        self.act("cancel-login")
        MOCK.valid_access.add("AT-seed")
        kc = os.path.join(self.dir, "kc", KC_FILE.replace("common", "consumers"))
        with open(kc, "w") as f:
            json.dump({"access_token": "AT-seed", "refresh_token": "RT-seed", "expires_at": int(time.time() * 1000) + 3600000}, f)
        it = self.sf("teams", "status busy", tenant="consumers")
        self.assertEqual(it[0]["title"], "Teams status needs a work or school account")

    def test_signing_in_as_someone_else_drops_their_cache(self):
        self.sign_in()
        self.sf("teams")
        self.assertTrue(os.path.isdir(os.path.join(self.dir, "cache", "user-1")))
        MOCK.me = {"id": "user-2", "displayName": "Other", "mail": "o@x.com"}
        MOCK.device_queue = ["ok"]
        self.act("login", M365_TEST_INTERVAL_SCALE="0.02")
        self.wait_for(lambda: "Signed in as Other" in self.read("notified"))
        self.assertFalse(os.path.exists(os.path.join(self.dir, "cache", "user-1")))


# ---------------------------------------------------------------- regression tests (audit pass 3)

class AuditThreeTests(Base):
    def portal_curl(self):
        path = os.path.join(self.dir, "portal")
        with open(path, "w") as f:
            f.write('#!/bin/bash\ncat >/dev/null\nprintf "HTTP/1.1 200 OK\\r\\ncontent-type: text/html\\r\\n\\r\\n<html>Hotel Wi-Fi login</html>"\n')
        os.chmod(path, 0o755)
        return path

    def test_captive_portal_is_not_an_empty_result(self):
        self.sign_in()
        it = self.sf("teams", M365_TEST_CURL=self.portal_curl())
        self.assertEqual(it[0]["title"], "You're offline")
        self.assertIn("captive portal", it[0]["subtitle"])
        self.assertFalse(os.path.exists(os.path.join(self.dir, "cache", "user-1", "events-2026-09-26.json")))
        self.assertIn("captive portal", self.act("login", M365_TEST_CURL=self.portal_curl()))

    def test_captive_portal_during_refresh_keeps_the_sign_in(self):
        self.sign_in(access="AT-old", expires_in=-10, valid=False)
        it = self.sf("outlook", M365_TEST_CURL=self.portal_curl())
        self.assertEqual(it[0]["title"], "You're offline")
        self.assertIsNotNone(self.tokens())

    def test_deleted_section_reloads_the_section_list(self):
        self.sign_in()
        MOCK.sections = [{"id": "gone", "displayName": "Inbox", "parentNotebook": {"displayName": "Work"}}]
        it = self.sf("onenote", "new T")
        self.item_act(it[0], onenote_section="Work/Inbox")
        MOCK.sections = [{"id": "moved", "displayName": "Inbox", "parentNotebook": {"displayName": "Work"}}]
        MOCK.overrides[("POST", "/v1.0/me/onenote/sections/gone/pages")] = [(404, {}, {"error": {"code": "20102", "message": "gone"}})]
        self.assertEqual(self.item_act(it[0], onenote_section="Work/Inbox"), "Created “T”")
        self.assertEqual(MOCK.created[-1][0], "/me/onenote/sections/moved/pages")


# ---------------------------------------------------------------- regression tests (audit pass 4)

class AuditFourTests(Base):
    def test_locked_keychain_is_not_signed_out(self):
        self.sign_in()
        for cmd in ("teams", "outlook", "onenote", "account"):
            it = self.sf(cmd, M365_TEST_KEYCHAIN_STATUS="-25308")
            self.assertEqual(it[0]["title"], "Your Keychain is locked", cmd)
            self.assertIn("-25308", it[0]["subtitle"])
        self.assertIn("Keychain", self.act("presence", "busy", m365_presence="busy", M365_TEST_KEYCHAIN_STATUS="-25308"))
        self.assertIsNotNone(self.tokens())

    def test_lock_held_by_a_reused_pid_is_ignored(self):
        # A live process that isn't this script (here: the test runner) owns the lock's PID.
        self.sign_in()
        MOCK.pages = [page("Old")]
        self.sf("onenote")
        age(os.path.join(self.dir, "cache", "user-1", "onenote-index.json"), 7200)
        MOCK.pages = [page("Fresh")]
        d = os.path.join(self.dir, "cache", "locks", "refresh-onenote-index.lock")
        os.makedirs(d)
        with open(os.path.join(d, "pid"), "w") as f:
            f.write(str(os.getpid()))
        it = self.sf("onenote")
        self.assertEqual(it[0]["title"], "Fresh")
        self.assertNotIn("rerun", self.last)

    def test_background_job_has_its_own_process_group(self):
        import socket
        black_hole = socket.socket()
        black_hole.bind(("127.0.0.1", 0))
        black_hole.listen(8)  # accepts connections, never answers
        port = black_hole.getsockname()[1]
        self.sign_in()
        try:
            it = self.sf("onenote", M365_TEST_BG_SYNC="", M365_GRAPH_BASE=f"http://127.0.0.1:{port}/v1.0")
            self.assertEqual(it[0]["title"], "Indexing your OneNote pages…")
            def job():
                out = subprocess.run(["ps", "-axo", "pid=,pgid=,command="], capture_output=True, text=True).stdout
                for line in out.splitlines():
                    if "m365.js refresh onenote-index" in line:
                        pid, pgid = line.split()[:2]
                        return pid, pgid
            pid, pgid = self.wait_for(job, timeout=10)
            self.assertEqual(pid, pgid)
            os.kill(int(pid), 15)
        finally:
            black_hole.close()

    def test_trace_ids_are_dropped_from_sign_in_errors(self):
        MOCK.devicecode_error = {"error": "invalid_request", "error_description":
                                 "AADSTS9002313: Invalid request. Request is malformed or invalid. Trace ID: 1 Correlation ID: 2 Timestamp: 3"}
        self.assertEqual(self.act("login"), "Invalid request. Request is malformed or invalid.")

    def test_teams_cloud_microsoft_links_open_in_the_app(self):
        self.sign_in()
        MOCK.events = [ev("New host", "2026-09-26T11:00:00", "2026-09-26T11:15:00", provider="unknown",
                          join="https://teams.cloud.microsoft/l/meetup-join/19%3ameeting_y%40thread.v2/0")]
        item = self.sf("teams")[0]
        self.assertEqual(item["title"], "New host")
        self.item_act(item, M365_TEST_APPS="msteams")
        self.assertEqual(self.read("opened").strip(), "msteams:/l/meetup-join/19%3ameeting_y%40thread.v2/0")

    def test_missing_permissions_are_shown(self):
        self.sign_in()
        p = os.path.join(self.dir, "kc", KC_FILE)
        with open(p) as f:
            t = json.load(f)
        t["scope"] = "User.Read Calendars.Read Mail.ReadWrite Notes.ReadWrite.All People.Read profile openid"
        with open(p, "w") as f:
            json.dump(t, f)
        it = self.sf("account")
        self.assertEqual(it[1]["title"], "Missing permission: Presence.ReadWrite")
        t["scope"] += " https://graph.microsoft.com/Presence.ReadWrite"
        with open(p, "w") as f:
            json.dump(t, f)
        self.assertNotIn("Missing", " ".join(self.titles(self.sf("account"))))

    def test_offline_presence_explains_teams_session(self):
        self.sign_in()
        MOCK.presence = {"availability": "Offline", "activity": "OffWork"}
        it = self.sf("teams", "status")
        self.assertEqual(it[0]["title"], "Current status: Offline")
        self.assertIn("signed in to Teams", it[0]["subtitle"])

    def test_section_name_onenote_cannot_create(self):
        self.sign_in()
        it = self.sf("onenote", "new T")
        self.assertIn("can't create a section", self.item_act(it[0], onenote_section="Q&A"))
        self.assertEqual(MOCK.created, [])

    def test_curl_config_quoting_round_trips(self):
        # Tokens and bodies travel in curl's "-K -" config: quotes, backslashes, tabs and a
        # literal backslash-n must arrive unchanged.
        self.sign_in(access='AT-"q\\z\\n', refresh='RT-x\\"y')
        MOCK.events = [ev("Standup", "2026-09-26T11:00:00", "2026-09-26T11:15:00")]
        self.assertEqual(self.sf("teams")[0]["title"], "Standup")
        self.assertEqual(self.tokens()["access_token"], 'AT-"q\\z\\n')  # accepted as sent, no refresh
        text = 'back\\slash "quote" \\n literal\ttab\r\nnext'
        it = self.sf("onenote", "new T :: " + text)
        self.assertEqual(self.item_act(it[0]), "Created “T”")
        self.assertIn('<p>back\\slash &quot;quote&quot; \\n literal\ttab<br/>next</p>', MOCK.created[-1][1])

    def test_hostile_queries(self):
        self.sign_in()
        MOCK.pages = [page("Ideas")]
        for q in ["   ", "-K -", "--help", "\u202eevil", "e\u0301\u0301\u0301", "a" * 5000, "status \u0000", "new ::", "new :: ::", "status 99999d"]:
            for cmd in ("teams", "onenote", "outlook", "account"):
                self.sf(cmd, q.replace("\u0000", ""))
        self.assertEqual(self.sf("teams", "status busy 99999d")[1]["subtitle"], "Set for 7 days")

    def test_huge_clipboard_is_refused(self):
        self.sign_in()
        it = self.sf("onenote", "new T", M365_TEST_CLIPBOARD="x" * 50, M365_TEST_CLIPBOARD_MAX="10")
        self.assertIn("too large", it[0]["title"])
        it = self.sf("onenote", "new T")
        self.assertIn("too large", self.item_act(it[0], M365_TEST_CLIPBOARD="x" * 50, M365_TEST_CLIPBOARD_MAX="10"))
        self.assertEqual(MOCK.created, [])


class FinalReviewTests(Base):
    def test_bidi_control_and_lone_surrogates_are_cleaned(self):
        self.sign_in()
        MOCK.mails = [{"subject": "Pay‮FDP.exe\u0007 \ud83d", "isRead": True,
                       "from": {"emailAddress": {"name": "E⁦vil", "address": "x‮@y.com"}},
                       "receivedDateTime": "2026-09-26T08:15:00Z", "webLink": "https://outlook.office365.com/owa/?1"}]
        out = self.run_js("outlook", "pay")
        self.assertNotIn("\\ud83d", out.lower())
        it = json.loads(out)["items"][0]
        self.assertEqual(it["title"], "PayFDP.exe �")
        self.assertTrue(it["subtitle"].startswith("Evil · "))
        self.assertEqual(it["mods"]["cmd"]["subtitle"], "Copy x@y.com")
        self.assertEqual(it["text"]["largetype"], "PayFDP.exe �")

    def test_corrupt_cache_shapes_are_refetched(self):
        self.sign_in()
        MOCK.events = [ev("Standup", "2026-09-26T11:00:00", "2026-09-26T11:15:00")]
        MOCK.pages = [page("Plan")]
        d = os.path.join(self.dir, "cache", "user-1")
        os.makedirs(d, exist_ok=True)
        for body in ("[]", "null", "5", '{"fetched_at": "x", "data": []}', '{"fetched_at": 1, "data": "str"}',
                     '{"fetched_at": 99999999999999, "data": []}'):
            for n in ("events-2026-09-26", "onenote-index", "presence"):
                with open(os.path.join(d, n + ".json"), "w") as f:
                    f.write(body)
            with open(os.path.join(d, "onenote-index.error.json"), "w") as f:
                f.write("[1]")
            self.assertEqual(self.sf("teams")[0]["title"], "Standup", body)
            self.assertEqual(self.sf("onenote")[0]["title"], "Plan", body)
            self.assertIn("Current status", self.sf("teams", "status")[0]["title"], body)

    def test_query_cache_is_capped(self):
        self.sign_in()
        for q in ("alpha", "bravo", "charlie", "delta"):
            self.sf("outlook", q, M365_TEST_QUERY_CACHE_MAX="2")
            time.sleep(0.02)
        files = [n for n in os.listdir(os.path.join(self.dir, "cache", "user-1")) if n.startswith("mail-")]
        self.assertEqual(len(files), 2)

    def test_people_named_like_object_properties(self):
        self.sign_in()
        MOCK.people = [{"displayName": "C", "scoredEmailAddresses": [{"address": "constructor"}]},
                       {"displayName": "P", "scoredEmailAddresses": [{"address": "__proto__"}]}]
        self.assertEqual(self.titles(self.sf("teams", "co"))[-2:], ["C", "P"])

    def test_missing_test_override_never_reaches_real_services(self):
        self.sign_in()
        for var in ("M365_TEST_KEYCHAIN_DIR", "M365_LOGIN_BASE", "M365_GRAPH_BASE", "M365_TEST_OPEN_FILE",
                    "M365_TEST_CLIPBOARD_OUT"):
            e = dict(self.env)
            del e[var]
            out = subprocess.run(["osascript", "-l", "JavaScript", "./m365.js", "act", "https://example.com/x"],
                                 cwd=SRC, env=dict(e, m365_action="copy" if "CLIPBOARD" in var else
                                                   "open" if "OPEN" in var else "logout" if "KEYCHAIN" in var else "login"),
                                 capture_output=True, text=True, timeout=60).stdout
            self.assertIn(f"Test mode: {var} is not set", out, var)
        # Test mode only ever talks to the local mock, even when an override names a real host.
        # (The mock never sees the request; an http:// URL on a closed port proves nothing left.)
        it = self.sf("teams", M365_GRAPH_BASE="http://graph.invalid/v1.0")
        self.assertIn("refusing", it[0]["title"])

    def test_refresh_job_stops_at_its_deadline(self):
        self.sign_in()
        MOCK.pages = [page("Plan")]
        self.run_js("refresh", "onenote-index", M365_TEST_JOB_DEADLINE="-1")
        with open(os.path.join(self.dir, "cache", "user-1", "onenote-index.error.json")) as f:
            self.assertEqual(json.load(f)["kind"], "timeout")
        self.assertEqual(MOCK.graph_requests("/v1.0/me/onenote/pages"), [])


class RoundFourTests(Base):
    """Round 4: Alfred's real runtime (minimal environment, paths with spaces, fresh install) and v1.1 features."""

    def alfred_env(self, root, **extra):
        # What Alfred passes: no LANG/LC_*, no Homebrew on PATH, its own variables, paths with spaces.
        b = "io.github.x-o-r-r-o.microsoft-365"
        e = {"HOME": os.environ.get("HOME", "/tmp"), "USER": os.environ.get("USER", ""), "TMPDIR": tempfile.gettempdir(),
             "PATH": "/usr/bin:/bin:/usr/sbin:/sbin",
             "alfred_workflow_data": os.path.join(root, "Application Support", "Alfred", "Workflow Data", b),
             "alfred_workflow_cache": os.path.join(root, "Caches", "com.runningwithcrayons.Alfred", "Workflow Data", b),
             "alfred_preferences": os.path.join(root, "Mobile Documents", "Alfred.alfredpreferences"),
             "alfred_version": "5.6", "alfred_version_build": "2290", "alfred_theme_subtext": "3",
             "alfred_workflow_bundleid": b, "alfred_workflow_name": "Microsoft 365",
             "alfred_workflow_uid": "user.workflow.1A2B", "alfred_workflow_version": "1.0.0", "alfred_debug": "1",
             "client_id": " " + CLIENT + " ", "tenant": " common ", "keyword_teams": "Téams", "keyword_onenote": "onenote",
             "keyword_outlook": "outlook", "keyword_account": "", "teams_open": "app", "onenote_open": "web", "onenote_section": "",
             "M365_LOGIN_BASE": MOCK.base, "M365_GRAPH_BASE": MOCK.base + "/v1.0",
             "M365_TEST_KEYCHAIN_DIR": os.path.join(self.dir, "kc"), "M365_TEST_CURL": CURL_WRAP, "M365_ARGV_LOG": ARGV_LOG,
             "M365_TEST_OPEN_FILE": os.path.join(self.dir, "opened"), "M365_TEST_NOTIFY_FILE": os.path.join(self.dir, "notified"),
             "M365_TEST_CLIPBOARD_OUT": os.path.join(self.dir, "copied"), "M365_TEST_CLIPBOARD": "", "M365_TEST_APPS": "",
             "M365_TEST_INTERVAL_SCALE": "0.01"}
        e.update(extra)
        return e

    def plist_scripts(self):
        with open(os.path.join(SRC, "info.plist"), "rb") as f:
            p = plistlib.load(f)
        out = {}
        for o in p["objects"]:
            c = o["config"]
            if "script" in c:
                m = __import__("re").search(r"m365\.js (\w+)", c["script"])
                out[m.group(1)] = c
        return out

    def run_alfred(self, cwd, env, name, arg):
        # Exactly like Alfred: /bin/bash -c <script> with the query as $1 (scriptargtype 1).
        script = self.plist_scripts()[name]["script"]
        r = subprocess.run(["/bin/bash", "-c", script, "bash", arg], cwd=cwd, env=env, capture_output=True, text=True, timeout=60)
        self.assertEqual(r.returncode, 0, r.stderr)
        return r.stdout

    def test_fresh_install_in_alfreds_environment(self):
        root = os.path.join(self.dir, "Library with spaces")
        wf = os.path.join(root, "Mobile Documents", "Alfred.alfredpreferences", "workflows", "user.workflow.1A2B")
        shutil.copytree(SRC, wf)
        env = self.alfred_env(root)
        self.assertFalse(os.path.exists(env["alfred_workflow_data"]))
        # First keystroke, not signed in yet: the empty keyword setting still gives a sensible hint.
        data = json.loads(self.run_alfred(wf, env, "teams", ""))
        validate(data)
        self.assertEqual(data["items"][0]["title"], "Sign in to Microsoft 365")
        # Sign in through the real detached poller (no synchronous test shortcut).
        MOCK.device_queue = ["pending", "ok"]
        MOCK.devicecode = dict(MOCK.devicecode, interval=1)
        out = self.run_alfred(wf, dict(env, m365_action="login"), "act", "login")
        self.assertEqual(out, "Code ABCD-EFGH copied: paste it in the browser to sign in")
        self.wait_for(lambda: "Signed in as Zoë Tester" in self.read("notified"), timeout=30)
        with open(os.path.join(env["alfred_workflow_data"], "account.json")) as f:
            self.assertEqual(json.load(f)["id"], "user-1")
        # An action that opens a link prints nothing at all, so "only show if populated" hides the notification.
        self.assertEqual(self.run_alfred(wf, dict(env, m365_action="open"), "act", "https://example.com/x"), "")
        MOCK.people = [{"displayName": "Anna", "scoredEmailAddresses": [{"address": "anna@contoso.com"}]}]
        data = json.loads(self.run_alfred(wf, env, "teams", "anna"))
        validate(data)
        self.assertEqual(data["items"][-1]["title"], "Anna")
        # The empty account keyword falls back to the default in messages.
        MOCK.device_queue = ["expired"]
        self.run_alfred(wf, dict(env, m365_action="login"), "act", "login")
        self.wait_for(lambda: "Use the m365 keyword" in self.read("notified"), timeout=30)
        # Locks and caches live under the paths with spaces.
        self.assertTrue(os.path.isdir(os.path.join(env["alfred_workflow_cache"], "user-1")))

    def test_upgrade_from_v1_0_0_caches(self):
        # v1.0.0 presence and event caches (no status message, no location links) still work.
        self.sign_in()
        MOCK.events = [ev("Standup", "2026-09-26T10:30:00", "2026-09-26T10:45:00")]
        self.sf("teams")
        cache = os.path.join(self.dir, "cache", "user-1")
        with open(os.path.join(cache, "presence.json"), "w") as f:
            json.dump({"fetched_at": int(time.time() * 1000), "data": {"availability": "Busy", "activity": "InACall"}}, f)
        self.assertEqual(self.sf("teams", "status")[0]["title"], "Current status: Busy")
        self.assertEqual(self.titles(self.sf("teams", "message"))[-1], "Type your status message")
        self.assertEqual(self.sf("teams")[0]["title"], "Standup")

    def test_network_filters_terminate_the_previous_run(self):
        subprocess.run([sys.executable, "tools/build.py"], cwd=ROOT, check=True, capture_output=True)
        s = self.plist_scripts()
        self.assertEqual(s["teams"]["queuemode"], 2)
        self.assertEqual(s["outlook"]["queuemode"], 2)
        self.assertEqual(s["onenote"]["queuemode"], 1)
        self.assertEqual(s["account"]["queuemode"], 1)

    def test_test_mode_is_loopback_only(self):
        self.sign_in()
        port = MOCK.base.rsplit(":", 1)[1]
        for base in (f"http://localhost:{port}/v1.0", f"http://127.0.0.1.nip.io:{port}/v1.0", f"http://127.0.0.1@graph.invalid:{port}/v1.0"):
            it = self.sf("outlook", "x", M365_GRAPH_BASE=base)
            self.assertIn("refusing", it[0]["title"], base)
        # A proxy from the environment never sees the mock's traffic.
        MOCK.mails = []
        it = self.sf("outlook", "x", http_proxy="http://127.0.0.1:9", HTTP_PROXY="http://127.0.0.1:9", ALL_PROXY="http://127.0.0.1:9")
        self.assertEqual(it[0]["title"], "No mail matches “x”")

    def test_status_message(self):
        self.sign_in()
        MOCK.presence = {"availability": "Available", "activity": "Available",
                         "statusMessage": {"message": {"content": "<p>Out &amp; about</p>", "contentType": "html"}}}
        it = self.sf("teams", "status")
        self.assertEqual(it[1]["title"], "Status message: Out & about")
        it = self.sf("teams", "message")
        self.assertEqual(self.titles(it), ["Status message: Out & about", "Type your status message", "Clear the status message"])
        self.assertEqual(self.item_act(it[2]), "Status message cleared")
        body = json.loads(MOCK.graph_requests("/v1.0/me/presence/setStatusMessage")[-1][4])
        self.assertEqual(body, {"statusMessage": {"message": {"content": "", "contentType": "text"}}})
        it = self.sf("teams", 'message Back at 3 "sharp" :: 1h 30m')
        self.assertEqual(it[0]["title"], 'Set status message “Back at 3 "sharp"”')
        self.assertEqual(it[0]["subtitle"], "Clears after 1 h 30 min")
        self.assertEqual(self.item_act(it[0]), "Status message set for 1 h 30 min")
        body = json.loads(MOCK.graph_requests("/v1.0/me/presence/setStatusMessage")[-1][4])["statusMessage"]
        self.assertEqual(body["message"], {"content": 'Back at 3 "sharp"', "contentType": "text"})
        self.assertEqual(body["expiryDateTime"]["timeZone"], "UTC")
        self.assertRegex(body["expiryDateTime"]["dateTime"], r"^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}$")
        self.assertEqual(self.sf("teams", "message lunch :: soon")[0]["title"], "Unknown duration")
        self.assertIn("too long", self.sf("teams", "message " + "x" * 281)[0]["title"])
        it = self.sf("teams", "message lunch")
        MOCK.overrides[("POST", "/v1.0/me/presence/setStatusMessage")] = [(403, {}, {"error": {"code": "Forbidden", "message": "x"}})]
        self.assertIn("work or school", self.item_act(it[0]))
        shutil.copy(os.path.join(self.dir, "kc", KC_FILE), os.path.join(self.dir, "kc", f"oauth_{CLIENT}_consumers.secret"))
        self.assertIn("work or school", self.sf("teams", "message lunch", tenant="consumers")[0]["title"])

    def test_actions_without_a_message_print_nothing(self):
        # Straight from osascript (no shell wrapper): not even a newline.
        self.sign_in()
        for action, arg in (("open", "https://example.com/x"), ("join", "https://teams.microsoft.com/l/meetup-join/x"), ("cancel-nothing", "")):
            r = subprocess.run(["osascript", "-l", "JavaScript", "./m365.js", "act", arg], cwd=SRC,
                               env=dict(self.env, m365_action=action), capture_output=True, timeout=60)
            self.assertEqual(r.stdout, b"", action)
        r = subprocess.run(["osascript", "-l", "JavaScript", "./m365.js", "act", "x"], cwd=SRC,
                           env=dict(self.env, m365_action="copy"), capture_output=True, text=True, timeout=60)
        self.assertEqual(r.stdout, "Copied x\n")

    def test_ctrl_copies_the_chat_link(self):
        self.sign_in()
        MOCK.people = [{"displayName": "Anna", "scoredEmailAddresses": [{"address": "anna+x@contoso.com"}]}]
        it = self.sf("teams", "anna")
        self.assertEqual(it[-1]["mods"]["ctrl"]["subtitle"], "Copy the Teams chat link")
        self.item_act(it[-1], "ctrl")
        self.assertEqual(self.read("copied"), "https://teams.microsoft.com/l/chat/0/0?users=anna%2Bx%40contoso.com")
        with open(os.path.join(SRC, "info.plist"), "rb") as f:
            p = plistlib.load(f)
        teams = next(o["uid"] for o in p["objects"] if o["config"].get("keyword") == "{var:keyword_teams}")
        self.assertIn(262144, [c["modifiers"] for c in p["connections"][teams]])

    def test_join_links_from_location_and_body(self):
        self.sign_in()
        MOCK.events = [
            ev("Zoom call", "2026-09-26T11:00:00", "2026-09-26T11:30:00", join=None,
               location={"displayName": "https://us02web.zoom.us/j/123?pwd=abc)."}),
            ev("External Teams", "2026-09-26T12:00:00", "2026-09-26T12:30:00", join=None,
               bodyPreview="Join: https://teams.microsoft.com/l/meetup-join/19%3aext/0 Meeting ID 1"),
            ev("Lookalike", "2026-09-26T13:00:00", "2026-09-26T13:30:00", join=None,
               location={"displayName": "https://zoom.us.evil.example/j/1 https://evil.example/?u=https://zoom.us/j/2"}),
        ]
        it = self.sf("outlook")
        by = {i["title"]: i for i in it}
        self.assertEqual(by["Zoom call"]["mods"]["cmd"]["arg"], "https://us02web.zoom.us/j/123?pwd=abc")
        self.assertEqual(by["External Teams"]["mods"]["cmd"]["arg"], "https://teams.microsoft.com/l/meetup-join/19%3aext/0")
        self.assertIs(by["Lookalike"]["mods"]["cmd"]["valid"], False)
        # Only Teams meetings are listed under the teams keyword, including the one from another organization.
        self.assertEqual(self.titles(self.sf("teams"))[:1], ["External Teams"])
        self.item_act(self.sf("teams")[0], M365_TEST_APPS="msteams")
        self.assertEqual(self.read("opened").splitlines()[-1], "msteams:/l/meetup-join/19%3aext/0")


class PlistTests(unittest.TestCase):
    def test_build_and_plist(self):
        subprocess.run([sys.executable, "tools/build.py"], cwd=ROOT, check=True, capture_output=True)
        with open(os.path.join(SRC, "info.plist"), "rb") as f:
            p = plistlib.load(f)
        uids = [o["uid"] for o in p["objects"]]
        self.assertEqual(len(uids), len(set(uids)))
        for src, conns in p["connections"].items():
            self.assertIn(src, uids)
            for c in conns:
                self.assertIn(c["destinationuid"], uids)
        kws = [o["config"]["keyword"] for o in p["objects"] if o["config"].get("keyword")]
        self.assertEqual(len(kws), 4)
        for kw in kws:
            self.assertRegex(kw, r"^\{var:keyword_\w+\}$")
        self.assertTrue(p["readme"].startswith("## Usage"))
        self.assertTrue(any(o["type"] == "alfred.workflow.trigger.external" and o["config"]["triggerid"] == "notify" for o in p["objects"]))
        out = subprocess.run(["sips", "-g", "pixelWidth", os.path.join(SRC, "icon.png")], capture_output=True, text=True).stdout
        self.assertGreaterEqual(int(out.split()[-1]), 256)

    def test_icons_exist(self):
        with open(os.path.join(ROOT, "tools", "icons.json")) as f:
            names = json.load(f)
        for name in names:
            if name != "icon":
                self.assertTrue(os.path.exists(os.path.join(SRC, "icons", f"{name}.png")), name)


if __name__ == "__main__":
    unittest.main(verbosity=1)
