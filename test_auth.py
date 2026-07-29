"""Executable spec for the auth subsystem, ported from the DMR Reconciler
handoff §10: login/throttling, stateless sessions, setup/recovery, the
session-gate middleware, invites, password reset, email confirmation, and
team-page authority checks. No test touches the network — the mailer is
monkeypatched at module level."""

from __future__ import annotations

import re
import urllib.parse

import pytest

import auth
import mailer
from app import create_app

CODE = "root-setup-code"
PW = "hunter2-strong"


def make_app(tmp_path, **kw):
    kw.setdefault("setup_code", CODE)
    kw.setdefault("secret", "s" * 64)
    kw.setdefault("cookie_secure", False)
    return create_app(data_dir=tmp_path / "data", **kw)


def bootstrap(app, username="boss", password=PW, **extra):
    """Create the admin via /setup; returns a logged-in client."""
    client = app.test_client()
    resp = client.post("/setup", data={"code": CODE, "username": username,
                                       "password": password, **extra})
    assert resp.status_code == 303, resp.data
    return client


def login(app, username, password):
    client = app.test_client()
    resp = client.post("/login", data={"username": username,
                                       "password": password})
    return client, resp


@pytest.fixture
def outbox(monkeypatch):
    """Email 'provider on': capture every message instead of sending."""
    sent = []
    monkeypatch.setattr(mailer, "enabled", lambda: True)
    monkeypatch.setattr(
        mailer, "send",
        lambda to, subject, html, text: sent.append(
            {"to": to, "subject": subject, "html": html, "text": text})
        or True)
    return sent


def link_from(message, kind):
    m = re.search(rf"https?://[^\s\"<>]+/{kind}/[A-Za-z0-9_-]+",
                  message["text"])
    assert m, f"no /{kind}/ link in email"
    return urllib.parse.urlparse(m.group(0)).path


# --- login ------------------------------------------------------------------

def test_login_success_wrong_and_unknown(tmp_path):
    app = make_app(tmp_path)
    bootstrap(app)
    ok, resp = login(app, "boss", PW)
    assert resp.status_code == 303
    assert ok.get("/").status_code == 200

    _, wrong = login(app, "boss", "not-the-password")
    assert wrong.status_code == 401
    _, unknown = login(app, "nobody", "not-the-password")
    assert unknown.status_code == 401
    # one generic message: never distinguish unknown-user from wrong-password
    # (the page reflects the TYPED username back into the form; normalize
    # that away — it reveals the input, not whether the account exists)
    assert wrong.data.replace(b"boss", b"NAME") \
        == unknown.data.replace(b"nobody", b"NAME")


def test_login_unicode_password(tmp_path):
    app = make_app(tmp_path)
    bootstrap(app, password="Zürich2026！")
    _, resp = login(app, "boss", "Zürich2026！")
    assert resp.status_code == 303
    _, resp = login(app, "boss", "pässwörd-wrong")
    assert resp.status_code == 401


def test_login_user_bucket_throttles_and_clears(tmp_path):
    app = make_app(tmp_path)
    bootstrap(app)
    client = app.test_client()
    for _ in range(5):
        assert client.post("/login", data={"username": "boss",
                                           "password": "bad"}
                           ).status_code == 401
    resp = client.post("/login", data={"username": "boss", "password": PW})
    assert resp.status_code == 429          # even the right password waits
    assert b"429" in str(resp.status).encode() or resp.status_code == 429

    # a successful login clears the user bucket but only releases ONE ip slot
    ctx = app.extensions["auth"]
    ctx.throttle.clear("user", "boss")
    ip_key = ("ip", "127.0.0.1")
    before = len(ctx.throttle._buckets.get(ip_key, []))
    _, resp = login(app, "boss", PW)
    assert resp.status_code == 303
    assert ("user", "boss") not in ctx.throttle._buckets
    after = len(ctx.throttle._buckets.get(ip_key, []))
    assert after == before                   # reserved slot released, rest kept


def test_login_throttle_window_slides(tmp_path, monkeypatch):
    app = make_app(tmp_path)
    bootstrap(app)
    client = app.test_client()
    for _ in range(5):
        client.post("/login", data={"username": "boss", "password": "bad"})
    assert client.post("/login", data={"username": "boss", "password": PW}
                       ).status_code == 429
    real = auth.now()
    monkeypatch.setattr(auth, "now", lambda: real + 301)  # window passes
    assert client.post("/login", data={"username": "boss", "password": PW}
                       ).status_code == 303


# --- sessions ---------------------------------------------------------------

def test_session_tamper_expiry_and_flags(tmp_path, monkeypatch):
    app = make_app(tmp_path)
    client = bootstrap(app)
    ctx = app.extensions["auth"]
    user = ctx.db.get_user("boss")

    good = auth.make_session_cookie("boss", user["password_hash"], ctx.secret)
    tampered = good[:-2] + ("aa" if not good.endswith("aa") else "bb")
    anon = app.test_client()
    anon.set_cookie(auth.SESSION_COOKIE, tampered)
    assert anon.get("/").status_code == 303   # rejected -> redirect

    # server-side expiry check (not just the cookie's max-age)
    real = auth.now()
    monkeypatch.setattr(auth, "now", lambda: real + auth.SESSION_TTL + 60)
    stale = app.test_client()
    stale.set_cookie(auth.SESSION_COOKIE, good)
    assert stale.get("/").status_code == 303
    monkeypatch.setattr(auth, "now", lambda: real)

    # cookie attributes
    resp = app.test_client().post(
        "/login", data={"username": "boss", "password": PW})
    cookie = resp.headers["Set-Cookie"]
    assert "HttpOnly" in cookie and "SameSite=Lax" in cookie
    secure_app = make_app(tmp_path / "sec", cookie_secure=True)
    bootstrap(secure_app)
    resp = secure_app.test_client().post(
        "/login", data={"username": "boss", "password": PW},
        base_url="https://localhost")
    assert "Secure" in resp.headers["Set-Cookie"]


def test_password_change_invalidates_sessions(tmp_path):
    app = make_app(tmp_path)
    admin = bootstrap(app)
    assert admin.get("/").status_code == 200
    # self password change: even the session that made it dies
    resp = admin.post("/team/password", data={"username": "boss",
                                              "password": "new-password9"})
    assert resp.status_code == 303
    assert admin.get("/").status_code == 303  # logged out everywhere
    _, resp = login(app, "boss", "new-password9")
    assert resp.status_code == 303


# --- setup ------------------------------------------------------------------

def test_setup_wrong_code_and_throttle(tmp_path):
    app = make_app(tmp_path)
    client = app.test_client()
    for _ in range(5):
        assert client.post("/setup", data={"code": "nope", "username": "x",
                                           "password": PW}
                           ).status_code == 401
    # blocked now, even with the correct code
    assert client.post("/setup", data={"code": CODE, "username": "x",
                                       "password": PW}).status_code == 429


def test_setup_correct_code_clears_bucket(tmp_path):
    app = make_app(tmp_path)
    client = app.test_client()
    for _ in range(4):
        client.post("/setup", data={"code": "nope", "username": "x",
                                    "password": PW})
    assert client.post("/setup", data={"code": CODE, "username": "boss",
                                       "password": PW}).status_code == 303
    ctx = app.extensions["auth"]
    assert ("setup", "127.0.0.1") not in ctx.throttle._buckets


def test_setup_validation(tmp_path):
    app = make_app(tmp_path)
    client = app.test_client()
    bad = [({"username": "X!bad", "password": PW}, b""),
           ({"username": "ok-user", "password": "short"}, b""),
           ({"username": "ok-user", "password": PW,
             "email": "not-an-email"}, b"")]
    for extra, _ in bad:
        assert client.post("/setup", data={"code": CODE, **extra}
                           ).status_code == 400


def test_setup_rerun_resets_admin_and_keeps_confirmed_email(tmp_path, outbox):
    app = make_app(tmp_path)
    bootstrap(app, email="boss@example.com")
    ctx = app.extensions["auth"]
    assert ctx.db.get_user("boss")["email_verified"] == 0
    assert len(outbox) == 1                       # confirmation sent
    app.test_client().get(link_from(outbox[0], "verify"))
    assert ctx.db.get_user("boss")["email_verified"] == 1

    # re-run with a BLANK email: address and confirmation preserved
    client = app.test_client()
    resp = client.post("/setup", data={"code": CODE, "username": "boss",
                                       "password": "brand-new-pass"})
    assert resp.status_code == 303
    user = ctx.db.get_user("boss")
    assert user["email"] == "boss@example.com"
    assert user["email_verified"] == 1
    _, resp = login(app, "boss", "brand-new-pass")
    assert resp.status_code == 303
    _, resp = login(app, "boss", PW)
    assert resp.status_code == 401                # old password is gone

    # re-run with a NEW email: stored unconfirmed + confirmation sent
    client.post("/setup", data={"code": CODE, "username": "boss",
                                "password": "brand-new-pass",
                                "email": "new@example.com"})
    user = ctx.db.get_user("boss")
    assert user["email"] == "new@example.com"
    assert user["email_verified"] == 0
    assert len(outbox) == 2


# --- middleware -------------------------------------------------------------

def test_gate_redirects_and_public_paths(tmp_path):
    app = make_app(tmp_path)
    client = app.test_client()
    # zero users -> first-run bootstrap
    r = client.get("/")
    assert r.status_code == 303 and r.headers["Location"].endswith("/setup")
    bootstrap(app)
    r = client.get("/")
    assert r.status_code == 303 and r.headers["Location"].endswith("/login")
    # public paths reachable logged-out (token pages 404 on junk, not 303)
    assert client.get("/login").status_code == 200
    assert client.get("/forgot").status_code == 200
    assert client.get("/healthz").status_code == 200
    assert client.get("/invite/junk").status_code == 404
    assert client.get("/reset/junk").status_code == 404
    assert client.get("/verify/junk").status_code == 404


def test_gate_fails_closed_when_unconfigured(tmp_path):
    app = make_app(tmp_path, setup_code="", open_access=False)
    client = app.test_client()
    assert client.get("/").status_code == 503
    assert client.get("/healthz").status_code == 200   # public passes first
    assert client.get("/login").status_code == 200
    open_app = make_app(tmp_path / "open", setup_code="", open_access=True)
    assert open_app.test_client().get("/").status_code == 200


def test_gate_xhr_gets_json_401(tmp_path):
    app = make_app(tmp_path)
    bootstrap(app)
    r = app.test_client().post(
        "/process", headers={"X-Requested-With": "XMLHttpRequest"})
    assert r.status_code == 401
    assert "error" in r.get_json()


# --- invites ----------------------------------------------------------------

def _invite_coworker(admin, outbox, username="ada", email="ada@example.com",
                     **extra):
    resp = admin.post("/team/add", data={"username": username,
                                         "email": email, **extra})
    assert "msg_invite_sent" in resp.headers["Location"]
    return link_from(outbox[-1], "invite")


def test_invite_full_lifecycle(tmp_path, outbox):
    app = make_app(tmp_path)
    admin = bootstrap(app)
    path = _invite_coworker(admin, outbox)
    ctx = app.extensions["auth"]

    # pending account is inert: no password authenticates
    assert ctx.db.get_user("ada")["password_hash"] == ""
    _, resp = login(app, "ada", "")
    assert resp.status_code == 401
    _, resp = login(app, "ada", "anything-at-all")
    assert resp.status_code == 401

    visitor = app.test_client()
    assert visitor.get(path).status_code == 200        # peek, not consumed
    assert visitor.get(path).status_code == 200        # still alive
    # a too-short password fails validation WITHOUT burning the link
    assert visitor.post(path, data={"password": "short"}).status_code == 400
    assert visitor.get(path).status_code == 200
    # accepting sets the password, confirms the email, signs in
    resp = visitor.post(path, data={"password": "ada-password9"})
    assert resp.status_code == 303
    assert visitor.get("/").status_code == 200
    user = ctx.db.get_user("ada")
    assert user["password_hash"] != "" and user["email_verified"] == 1
    _, resp = login(app, "ada", "ada-password9")
    assert resp.status_code == 303
    # single-use: the link is dead now
    assert app.test_client().get(path).status_code == 404
    assert app.test_client().post(path, data={"password": "x" * 10}
                                  ).status_code == 404


def test_invite_expires(tmp_path, outbox, monkeypatch):
    app = make_app(tmp_path)
    admin = bootstrap(app)
    path = _invite_coworker(admin, outbox)
    real = auth.now()
    monkeypatch.setattr(auth, "now", lambda: real + 73 * 3600)
    assert app.test_client().get(path).status_code == 404


def test_tokens_are_purpose_scoped(tmp_path, outbox):
    app = make_app(tmp_path)
    admin = bootstrap(app)
    invite_path = _invite_coworker(admin, outbox)
    ctx = app.extensions["auth"]
    raw_reset = ctx.db.issue_token("ada", "reset", "ada@example.com", 3600)
    client = app.test_client()
    # a reset token cannot activate an invite, and vice versa
    invite_raw = invite_path.rsplit("/", 1)[-1]
    assert client.get(f"/invite/{raw_reset}").status_code == 404
    assert client.get(f"/reset/{invite_raw}").status_code == 404
    assert client.get(f"/verify/{invite_raw}").status_code == 404


def test_invite_send_failure_keeps_account(tmp_path, outbox, monkeypatch):
    app = make_app(tmp_path)
    admin = bootstrap(app)
    monkeypatch.setattr(mailer, "send", lambda *a, **k: False)
    resp = admin.post("/team/add", data={"username": "eve",
                                         "email": "eve@example.com"})
    assert "msg_invite_send_failed" in resp.headers["Location"]
    ctx = app.extensions["auth"]
    assert ctx.db.get_user("eve") is not None          # account survives
    # resend works once the provider recovers
    monkeypatch.setattr(
        mailer, "send",
        lambda to, subject, html, text: outbox.append(
            {"to": to, "subject": subject, "html": html, "text": text})
        or True)
    resp = admin.post("/team/resend-invite", data={"username": "eve"})
    assert "msg_invite_resent" in resp.headers["Location"]
    path = link_from(outbox[-1], "invite")
    assert app.test_client().get(path).status_code == 200


def test_manual_add_without_email(tmp_path):
    app = make_app(tmp_path)
    admin = bootstrap(app)
    # no email and no password -> refused
    resp = admin.post("/team/add", data={"username": "manu"})
    assert "err_password_required" in resp.headers["Location"]
    resp = admin.post("/team/add", data={"username": "manu",
                                         "password": "manu-pass99"})
    assert "msg_user_added" in resp.headers["Location"]
    _, resp = login(app, "manu", "manu-pass99")
    assert resp.status_code == 303


# --- password reset ---------------------------------------------------------

def _confirmed_user(app, outbox, email="boss@example.com"):
    client = bootstrap(app, email=email)
    app.test_client().get(link_from(outbox[-1], "verify"))
    return client


def test_forgot_only_mails_confirmed_addresses(tmp_path, outbox):
    app = make_app(tmp_path)
    _confirmed_user(app, outbox)
    admin = login(app, "boss", PW)[0]
    admin.post("/team/add", data={"username": "raw",
                                  "password": "raw-pass-99"})
    raw = login(app, "raw", "raw-pass-99")[0]
    raw.post("/team/email", data={"username": "raw",
                                  "email": "raw@example.com"})
    base = len(outbox)

    client = app.test_client()
    # confirmed address -> reset link sent
    r1 = client.post("/forgot", data={"email": "boss@example.com"})
    assert len(outbox) == base + 1
    # unconfirmed address -> nothing sent, identical page
    r2 = client.post("/forgot", data={"email": "raw@example.com"})
    # unknown address -> nothing sent, identical page
    r3 = client.post("/forgot", data={"email": "ghost@example.com"})
    assert len(outbox) == base + 1
    assert r1.data == r2.data == r3.data
    assert r1.status_code == r2.status_code == r3.status_code == 200


def test_forgot_throttles_silently_and_shares_setup_bucket(tmp_path, outbox):
    app = make_app(tmp_path)
    _confirmed_user(app, outbox)
    client = app.test_client()
    base = len(outbox)
    for _ in range(5):
        assert client.post("/forgot", data={"email": "boss@example.com"}
                           ).status_code == 200
    assert len(outbox) == base + 5
    # 6th: same success page, but nothing is sent
    assert client.post("/forgot", data={"email": "boss@example.com"}
                       ).status_code == 200
    assert len(outbox) == base + 5
    # /forgot and /setup share the per-IP bucket
    assert client.post("/setup", data={"code": CODE, "username": "x",
                                       "password": PW}).status_code == 429


def test_reset_lifecycle(tmp_path, outbox):
    app = make_app(tmp_path)
    old_session = _confirmed_user(app, outbox)
    client = app.test_client()
    client.post("/forgot", data={"email": "boss@example.com"})
    first = link_from(outbox[-1], "reset")
    client.post("/forgot", data={"email": "boss@example.com"})
    second = link_from(outbox[-1], "reset")
    assert first != second

    visitor = app.test_client()
    assert visitor.get(second).status_code == 200      # peek keeps it alive
    resp = visitor.post(second, data={"password": "reset-pass-77"})
    assert resp.status_code == 303
    assert visitor.get("/").status_code == 200         # signed straight in

    ctx = app.extensions["auth"]
    assert ctx.db.get_user("boss")["email_verified"] == 1
    _, resp = login(app, "boss", "reset-pass-77")
    assert resp.status_code == 303
    _, resp = login(app, "boss", PW)
    assert resp.status_code == 401                     # old password dead
    assert old_session.get("/").status_code == 303     # old sessions dead
    assert app.test_client().get(second).status_code == 404  # used
    assert app.test_client().get(first).status_code == 404   # sibling killed


# --- email confirmation -----------------------------------------------------

def test_verify_confirms_once(tmp_path, outbox):
    app = make_app(tmp_path)
    bootstrap(app, email="boss@example.com")
    path = link_from(outbox[-1], "verify")
    ctx = app.extensions["auth"]
    assert app.test_client().get(path).status_code == 200
    assert ctx.db.get_user("boss")["email_verified"] == 1
    assert app.test_client().get(path).status_code == 404  # second click


def test_verify_dead_after_address_change(tmp_path, outbox):
    app = make_app(tmp_path)
    admin = bootstrap(app, email="boss@example.com")
    path = link_from(outbox[-1], "verify")
    # the account moves to a different address before the click
    admin.post("/team/email", data={"username": "boss",
                                    "email": "other@example.com"})
    assert app.test_client().get(path).status_code == 404
    ctx = app.extensions["auth"]
    assert ctx.db.get_user("boss")["email_verified"] == 0


def test_email_change_always_drops_confirmation(tmp_path, outbox):
    app = make_app(tmp_path)
    admin = _confirmed_user(app, outbox)
    ctx = app.extensions["auth"]
    assert ctx.db.get_user("boss")["email_verified"] == 1
    # even re-entering the SAME address drops confirmation
    admin.post("/team/email", data={"username": "boss",
                                    "email": "boss@example.com"})
    assert ctx.db.get_user("boss")["email_verified"] == 0
    # empty address unbinds
    admin.post("/team/email", data={"username": "boss", "email": ""})
    user = ctx.db.get_user("boss")
    assert user["email"] is None and user["email_verified"] == 0


# --- team authority ---------------------------------------------------------

def _member(app, name="memb", password="member-pass9"):
    admin = login(app, "boss", PW)[0]
    admin.post("/team/add", data={"username": name, "password": password})
    return login(app, name, password)[0]


def test_non_admin_cannot_manage_accounts(tmp_path):
    app = make_app(tmp_path)
    bootstrap(app)
    member = _member(app)
    ctx = app.extensions["auth"]

    resp = member.post("/team/add", data={"username": "smuggled",
                                          "password": "x" * 10})
    assert "err_not_admin" in resp.headers["Location"]
    assert ctx.db.get_user("smuggled") is None

    resp = member.post("/team/delete", data={"username": "boss"})
    assert "err_not_admin" in resp.headers["Location"]
    assert ctx.db.get_user("boss") is not None

    # the username field is advisory: a member naming someone else is refused
    resp = member.post("/team/password", data={"username": "boss",
                                               "password": "x" * 10})
    assert "err_not_admin" in resp.headers["Location"]
    _, ok = login(app, "boss", PW)
    assert ok.status_code == 303                      # boss unchanged
    resp = member.post("/team/email", data={"username": "boss",
                                            "email": "evil@example.com"})
    assert "err_not_admin" in resp.headers["Location"]
    assert ctx.db.get_user("boss")["email"] is None

    # but members manage THEMSELVES fine
    resp = member.post("/team/email", data={"username": "memb",
                                            "email": "me@example.com"})
    assert "msg_email_saved" in resp.headers["Location"] \
        or "msg_email_saved_noconfirm" in resp.headers["Location"]


def test_admin_resets_member_password(tmp_path):
    app = make_app(tmp_path)
    bootstrap(app)
    member = _member(app)
    assert member.get("/").status_code == 200
    admin = login(app, "boss", PW)[0]
    resp = admin.post("/team/password", data={"username": "memb",
                                              "password": "rotated-99"})
    assert "msg_pw_changed" in resp.headers["Location"]
    assert member.get("/").status_code == 303          # member logged out
    _, resp = login(app, "memb", "rotated-99")
    assert resp.status_code == 303


def test_duplicate_email_rejected_everywhere(tmp_path, outbox):
    app = make_app(tmp_path)
    admin = bootstrap(app, email="boss@example.com")
    # /team/add
    resp = admin.post("/team/add", data={"username": "dup",
                                         "email": "BOSS@example.com"})
    assert "auth_email_taken" in resp.headers["Location"]
    # /team/email
    admin.post("/team/add", data={"username": "memb",
                                  "password": "member-pass9"})
    resp = admin.post("/team/email", data={"username": "memb",
                                           "email": "Boss@Example.com"})
    assert "auth_email_taken" in resp.headers["Location"]
    # /setup for a different username
    resp = app.test_client().post(
        "/setup", data={"code": CODE, "username": "boss2", "password": PW,
                        "email": "boss@EXAMPLE.com"})
    assert resp.status_code == 400


def test_delete_guards(tmp_path):
    app = make_app(tmp_path)
    admin = bootstrap(app)
    member = _member(app)

    resp = admin.post("/team/delete", data={"username": "ghost"})
    assert "err_no_such_user" in resp.headers["Location"]
    resp = admin.post("/team/delete", data={"username": "boss"})
    assert "err_delete_self" in resp.headers["Location"]

    # deleting a member works and kills their session
    resp = admin.post("/team/delete", data={"username": "memb"})
    assert "msg_user_deleted" in resp.headers["Location"]
    assert member.get("/").status_code == 303

    # a second admin can be deleted while two exist; the survivor can
    # never delete themself, so the instance cannot orphan itself
    admin.post("/team/add", data={"username": "boss2", "is_admin": "1",
                                  "password": "second-admin9"})
    resp = admin.post("/team/delete", data={"username": "boss2"})
    assert "msg_user_deleted" in resp.headers["Location"]
    resp = admin.post("/team/delete", data={"username": "boss"})
    assert "err_delete_self" in resp.headers["Location"]


def test_last_admin_guard_at_db_level(tmp_path):
    """The route guard for 'cannot delete the last admin' (defense in
    depth: normally unreachable because the actor is always an admin and
    self-deletion is blocked first)."""
    app = make_app(tmp_path)
    admin = bootstrap(app)
    ctx = app.extensions["auth"]
    admin.post("/team/add", data={"username": "boss2", "is_admin": "1",
                                  "password": "second-admin9"})
    other = login(app, "boss2", "second-admin9")[0]
    ctx.db.update_user("boss2", is_admin=0)  # demoted mid-session
    resp = other.post("/team/delete", data={"username": "boss"})
    assert "err_not_admin" in resp.headers["Location"]
    assert ctx.db.get_user("boss") is not None


# --- hardening regressions (adversarial-review findings) --------------------

def test_non_ascii_cookie_is_anonymous_not_500(tmp_path):
    """A garbage session cookie with non-ASCII bytes in the signature must
    read as anonymous — compare_digest on str would raise TypeError and
    500 every request until the browser cookie was cleared."""
    app = make_app(tmp_path)
    bootstrap(app)
    for value in ("a|1|b|café", "boss|999|zz|éé", "ÿ"):
        client = app.test_client()
        client.set_cookie(auth.SESSION_COOKIE, value)
        assert client.get("/").status_code == 303       # anonymous redirect
        assert client.get("/login").status_code == 200  # public still fine
        assert client.get("/healthz").status_code == 200


def test_team_email_reports_send_failure(tmp_path, outbox, monkeypatch):
    """The address is saved either way, but a failed confirmation send must
    not be reported as 'a confirmation link was sent'."""
    app = make_app(tmp_path)
    admin = bootstrap(app)
    monkeypatch.setattr(mailer, "send", lambda *a, **k: False)
    resp = admin.post("/team/email", data={"username": "boss",
                                           "email": "boss@example.com"})
    assert "msg_email_saved_sendfail" in resp.headers["Location"]
    user = app.extensions["auth"].db.get_user("boss")
    assert user["email"] == "boss@example.com"
    assert user["email_verified"] == 0


def test_startup_validation_unconditional(tmp_path, monkeypatch):
    """An environment-driven boot with no setup code and no explicit
    open-access opt-out must fail at startup — on any host, not only Fly."""
    monkeypatch.delenv("APP_PASSWORD", raising=False)
    monkeypatch.delenv("ALLOW_OPEN_ACCESS", raising=False)
    with pytest.raises(RuntimeError):
        create_app(data_dir=tmp_path / "boom")
    monkeypatch.setenv("ALLOW_OPEN_ACCESS", "1")
    assert create_app(data_dir=tmp_path / "ok") is not None
    monkeypatch.delenv("ALLOW_OPEN_ACCESS", raising=False)
    monkeypatch.setenv("APP_PASSWORD", "some-setup-code")
    assert create_app(data_dir=tmp_path / "ok2") is not None


def test_mailer_send_never_raises(monkeypatch):
    """urlopen can raise beyond URLError (e.g. BadStatusLine); send() must
    contain anything and return False."""
    import http.client
    import urllib.request

    monkeypatch.setenv("RESEND_API_KEY", "re_test")
    for exc in (http.client.BadStatusLine("garbage"),
                ConnectionResetError(),
                ValueError("bad")):
        def boom(*a, _exc=exc, **k):
            raise _exc
        monkeypatch.setattr(urllib.request, "urlopen", boom)
        assert mailer.send("a@example.com", "s", "<p>h</p>", "t") is False


# --- primitives -------------------------------------------------------------

def test_password_hash_roundtrip_and_malformed():
    h = auth.hash_password("secret-password")
    assert auth.verify_password("secret-password", h)
    assert not auth.verify_password("wrong", h)
    for bad in ("", "no-dollar", "$", "zz$zz", None):
        assert not auth.verify_password("anything", bad)


def test_token_consume_is_single_use(tmp_path):
    db = auth.AuthDB(tmp_path / "t.db")
    db.create_user("u", auth.hash_password("x" * 10))
    raw = db.issue_token("u", "invite", "u@example.com", 3600)
    assert db.peek_token(raw, "invite") is not None
    assert db.peek_token(raw, "reset") is None          # purpose-scoped
    assert db.consume_token(raw, "invite") is not None
    assert db.consume_token(raw, "invite") is None      # exactly once
    assert db.peek_token(raw, "invite") is None
    # only the hash is stored — the raw token never touches the database
    import sqlite3
    with sqlite3.connect(tmp_path / "t.db") as conn:
        stored = conn.execute("SELECT token_hash FROM auth_tokens").fetchall()
    assert all(raw not in row[0] for row in stored)
