"""Login / setup / team / token routes for the closed-team auth system.

Route behavior follows the DMR Reconciler auth handoff spec §6. All form
POSTs redirect 303; /team/* feedback travels as ?msg=/?error= translation
keys on the redirect back to /team. The session gate itself lives in
app.py (create_app); everything here assumes it already ran.
"""

from __future__ import annotations

import hmac
import logging

from flask import (
    Blueprint, current_app, g, redirect, render_template, request, url_for,
)

import auth
import mailer
from translations import STRINGS, translate

log = logging.getLogger(__name__)

bp = Blueprint("auth", __name__)


def _tr(key: str, **fmt) -> str:
    return translate(getattr(g, "lang", "en"), key, **fmt)


def _ctx() -> auth.AuthContext:
    return current_app.extensions["auth"]


def _client_ip() -> str:
    # Fly sets (and strips from inbound traffic) Fly-Client-IP at the edge.
    # Never trust client-supplied X-Forwarded-For.
    return request.headers.get("Fly-Client-IP") or request.remote_addr or "?"


def _base_url() -> str:
    """Absolute origin for emailed links. PUBLIC_BASE_URL in production —
    never the Host header, which is attacker-controllable and would let a
    spoofed request plant a foreign domain inside a real reset email. The
    request origin is the fallback for local dev only."""
    ctx = _ctx()
    return ctx.public_base_url or request.url_root.rstrip("/")


def _signed_in(user, target: str = "/"):
    ctx = _ctx()
    resp = redirect(target, code=303)
    resp.set_cookie(
        auth.SESSION_COOKIE,
        auth.make_session_cookie(user["username"], user["password_hash"],
                                 ctx.secret),
        max_age=auth.SESSION_TTL, httponly=True, samesite="Lax",
        secure=ctx.cookie_secure)
    return resp


def _team_redirect(**params):
    return redirect(url_for("auth.team_page", **params), code=303)


def _send_confirmation(username: str, email: str) -> bool:
    ctx = _ctx()
    hours = ctx.invite_ttl_hours
    raw = ctx.db.issue_token(username, "verify", email, hours * 3600)
    return mailer.send_confirm(email, f"{_base_url()}/verify/{raw}", hours)


def _dead_link_page():
    return render_template(
        "auth/notice.html", title=_tr("auth_token_dead_title"),
        message=_tr("auth_token_dead"),
        link_href=url_for("auth.login_page"),
        link_label=_tr("auth_go_login")), 404


# --- login / logout ---------------------------------------------------------

@bp.route("/login", methods=["GET", "POST"])
def login_page():
    ctx = _ctx()
    if request.method == "GET":
        return render_template("auth/login.html")

    username = auth.normalize_username(request.form.get("username"))
    password = request.form.get("password") or ""
    wait = ctx.throttle.reserve([("user", username), ("ip", _client_ip())])
    if wait:
        return render_template("auth/login.html", username=username,
                               error=_tr("auth_throttled", n=wait)), 429

    user = ctx.db.get_user(username)
    # Unknown users verify against a dummy hash: same work, same timing.
    stored = user["password_hash"] if user else auth.DUMMY_HASH
    if auth.verify_password(password, stored) and user is not None:
        ctx.throttle.clear("user", username)
        ctx.throttle.release("ip", _client_ip())
        log.info("login ok: %s", username)
        return _signed_in(user)
    # One generic message — never distinguish unknown-user from wrong-password.
    return render_template("auth/login.html", username=username,
                           error=_tr("auth_login_failed")), 401


@bp.get("/logout")
def logout():
    # Sessions are stateless; deleting the cookie is the whole logout.
    resp = redirect(url_for("auth.login_page"), code=303)
    resp.delete_cookie(auth.SESSION_COOKIE)
    return resp


# --- setup: bootstrap & admin recovery --------------------------------------

@bp.route("/setup", methods=["GET", "POST"])
def setup_page():
    ctx = _ctx()
    if not ctx.setup_code:
        return render_template(
            "auth/notice.html", title=_tr("auth_setup_title"),
            message=_tr("auth_setup_disabled"),
            link_href=url_for("auth.login_page"),
            link_label=_tr("auth_go_login")), 503
    if request.method == "GET":
        return render_template("auth/setup.html")

    def fail(key, status=400):
        return render_template("auth/setup.html", error=_tr(key),
                               form=request.form), status

    wait = ctx.throttle.reserve([("setup", _client_ip())])
    if wait:
        return render_template("auth/setup.html",
                               error=_tr("auth_throttled", n=wait),
                               form=request.form), 429
    code = request.form.get("code") or ""
    if not hmac.compare_digest(code.encode("utf-8"),
                               ctx.setup_code.encode("utf-8")):
        return fail("auth_setup_wrong_code", 401)  # the failure slot stands
    ctx.throttle.clear("setup", _client_ip())

    username = auth.normalize_username(request.form.get("username"))
    if not auth.valid_username(username):
        return fail("auth_bad_username")
    password = request.form.get("password") or ""
    if not auth.valid_password(password):
        return fail("auth_bad_password")
    display = (request.form.get("display") or "").strip()
    email = auth.normalize_email(request.form.get("email"))
    if email:
        if not auth.valid_email(email):
            return fail("auth_bad_email")
        if ctx.db.email_taken(email, except_username=username):
            return fail("auth_email_taken")

    password_hash = auth.hash_password(password)
    existing = ctx.db.get_user(username)
    send_confirm_to = None
    if existing:
        # /setup doubles as admin recovery: blank fields keep what is on
        # file (never silently unbind); same email keeps its confirmation.
        ctx.db.update_user(username, password_hash=password_hash, is_admin=1,
                           display=display or existing["display"])
        if email and email != (existing["email"] or ""):
            ctx.db.update_user(username, email=email, email_verified=0)
            send_confirm_to = email
    else:
        ctx.db.create_user(username, password_hash, display=display or None,
                           is_admin=True, email=email or None,
                           email_verified=False)
        send_confirm_to = email or None
    if send_confirm_to and mailer.enabled():
        if not _send_confirmation(username, send_confirm_to):
            # Degraded path: the account works either way; confirmation can
            # be re-sent from /team once the provider recovers.
            log.warning("setup: confirmation email to %r failed to send",
                        send_confirm_to)
    log.info("setup: admin account %r %s", username,
             "reset" if existing else "created")
    user = ctx.db.get_user(username)
    return _signed_in(user)


# --- forgot password --------------------------------------------------------

@bp.route("/forgot", methods=["GET", "POST"])
def forgot_page():
    ctx = _ctx()
    if request.method == "GET":
        return render_template("auth/forgot.html",
                               email_on=mailer.enabled())

    # Every request counts (reset-spam + enumeration probing); when blocked,
    # render the success page anyway — silently.
    blocked = ctx.throttle.reserve([("setup", _client_ip())])
    email = auth.normalize_email(request.form.get("email"))
    if not blocked and mailer.enabled() and auth.valid_email(email):
        user = ctx.db.get_user_by_email(email)
        if user is not None and user["email_verified"]:
            # Reset links go ONLY to confirmed addresses — an unconfirmed
            # address may be a typo pointing at a stranger's inbox.
            hours = ctx.reset_ttl_hours
            raw = ctx.db.issue_token(user["username"], "reset", email,
                                     hours * 3600)
            mailer.send_reset(email, f"{_base_url()}/reset/{raw}", hours)
    # Identical response for every branch: no user enumeration.
    return render_template("auth/forgot.html", email_on=mailer.enabled(),
                           sent=True)


# --- emailed-token pages ----------------------------------------------------

def _password_form(purpose: str, raw: str, signin_extra=None):
    """Shared GET/POST handling for /invite/<t> and /reset/<t>."""
    ctx = _ctx()
    title_key = ("auth_pw_invite_title" if purpose == "invite"
                 else "auth_pw_reset_title")
    if request.method == "GET":
        # Non-destructive peek: rendering the form must not burn the link.
        if ctx.db.peek_token(raw, purpose) is None:
            return _dead_link_page()
        return render_template("auth/password_form.html",
                               title=_tr(title_key))

    password = request.form.get("password") or ""
    # Validation that can fail happens BEFORE consumption, so a typo does
    # not kill the link.
    if not auth.valid_password(password):
        if ctx.db.peek_token(raw, purpose) is None:
            return _dead_link_page()
        return render_template("auth/password_form.html",
                               title=_tr(title_key),
                               error=_tr("auth_bad_password")), 400
    row = ctx.db.consume_token(raw, purpose)
    if row is None:
        return _dead_link_page()
    username = row["username"]
    ctx.db.update_user(username, password_hash=auth.hash_password(password))
    # Clicking an emailed link proves ownership of that inbox — confirm the
    # address, scoped to the one the link was actually sent to.
    ctx.db.confirm_email(username, row["email"])
    if signin_extra:
        signin_extra(username)
    user = ctx.db.get_user(username)
    if user is None:  # account deleted while the link was in flight
        return _dead_link_page()
    log.info("%s accepted for %s", purpose, username)
    return _signed_in(user)


@bp.route("/invite/<token>", methods=["GET", "POST"])
def invite_page(token: str):
    return _password_form("invite", token)


@bp.route("/reset/<token>", methods=["GET", "POST"])
def reset_page(token: str):
    ctx = _ctx()
    # A successful reset also invalidates the user's other outstanding
    # reset links, so an older email in the same inbox cannot overwrite
    # the new password.
    return _password_form(
        "reset", token,
        signin_extra=lambda u: ctx.db.invalidate_tokens(u, "reset"))


@bp.get("/verify/<token>")
def verify_page(token: str):
    ctx = _ctx()
    row = ctx.db.consume_token(token, "verify")
    if row is None:
        return _dead_link_page()
    if not ctx.db.confirm_email(row["username"], row["email"]):
        # The account moved to a different address after this link was
        # sent; the link only proves ownership of the inbox it landed in.
        return _dead_link_page()
    return render_template(
        "auth/notice.html", title=_tr("auth_verified_title"),
        message=_tr("auth_verified_msg"),
        link_href=url_for("auth.login_page"),
        link_label=_tr("auth_go_login"))


# --- team management --------------------------------------------------------

@bp.get("/team")
def team_page():
    ctx = _ctx()
    if g.account is None:  # open-access mode has no accounts to manage
        return redirect(url_for("auth.login_page"), code=303)
    msg_key = request.args.get("msg", "")
    err_key = request.args.get("error", "")
    valid = STRINGS["en"]
    return render_template(
        "auth/team.html",
        users=ctx.db.list_users(),
        email_on=mailer.enabled(),
        msg=_tr(msg_key) if msg_key in valid else None,
        error=_tr(err_key) if err_key in valid else None,
    )


def _require_account():
    return g.account


def _target_username():
    """Resolve the account a /team form acts on. The form's username field
    is advisory — the server enforces authority: admins may name anyone,
    members only themselves. Returns None when the claim is not allowed."""
    requested = auth.normalize_username(request.form.get("username"))
    actor = g.account
    if actor["is_admin"]:
        return requested or actor["username"]
    if requested and requested != actor["username"]:
        return None
    return actor["username"]


@bp.post("/team/add")
def team_add():
    ctx = _ctx()
    actor = _require_account()
    if actor is None or not actor["is_admin"]:
        return _team_redirect(error="err_not_admin")
    username = auth.normalize_username(request.form.get("username"))
    if not auth.valid_username(username):
        return _team_redirect(error="auth_bad_username")
    if ctx.db.get_user(username) is not None:
        return _team_redirect(error="auth_username_taken")
    display = (request.form.get("display") or "").strip() or None
    is_admin = bool(request.form.get("is_admin"))
    email = auth.normalize_email(request.form.get("email"))
    if email:
        if not auth.valid_email(email):
            return _team_redirect(error="auth_bad_email")
        if ctx.db.email_taken(email):
            return _team_redirect(error="auth_email_taken")
    password = request.form.get("password") or ""

    if email and mailer.enabled():
        # Invite path: passwordless (inert) account; the coworker chooses
        # their own password via the emailed link — the admin never knows it.
        ctx.db.create_user(username, "", display=display, is_admin=is_admin,
                           email=email, email_verified=False)
        hours = ctx.invite_ttl_hours
        raw = ctx.db.issue_token(username, "invite", email, hours * 3600)
        sent = mailer.send_invite(email, f"{_base_url()}/invite/{raw}",
                                  inviter=g.user, hours=hours)
        if not sent:
            # The account still exists — the admin can resend or set a
            # password by hand.
            return _team_redirect(error="msg_invite_send_failed")
        return _team_redirect(msg="msg_invite_sent")

    # Manual path: no email, or provider off — an initial password is
    # required and the admin shares it privately.
    if not password:
        return _team_redirect(error="err_password_required")
    if not auth.valid_password(password):
        return _team_redirect(error="auth_bad_password")
    ctx.db.create_user(username, auth.hash_password(password),
                       display=display, is_admin=is_admin,
                       email=email or None, email_verified=False)
    return _team_redirect(msg="msg_user_added")


@bp.post("/team/resend-invite")
def team_resend_invite():
    ctx = _ctx()
    actor = _require_account()
    if actor is None or not actor["is_admin"]:
        return _team_redirect(error="err_not_admin")
    username = auth.normalize_username(request.form.get("username"))
    user = ctx.db.get_user(username)
    if user is None:
        return _team_redirect(error="err_no_such_user")
    if user["password_hash"] != "":
        return _team_redirect(error="err_pending_only")
    if not user["email"]:
        return _team_redirect(error="err_no_email_on_file")
    if not mailer.enabled():
        return _team_redirect(error="err_email_off")
    hours = ctx.invite_ttl_hours
    raw = ctx.db.issue_token(username, "invite", user["email"], hours * 3600)
    sent = mailer.send_invite(user["email"], f"{_base_url()}/invite/{raw}",
                              inviter=g.user, hours=hours)
    return (_team_redirect(msg="msg_invite_resent") if sent
            else _team_redirect(error="msg_send_failed"))


@bp.post("/team/delete")
def team_delete():
    ctx = _ctx()
    actor = _require_account()
    if actor is None or not actor["is_admin"]:
        return _team_redirect(error="err_not_admin")
    username = auth.normalize_username(request.form.get("username"))
    target = ctx.db.get_user(username)
    if target is None:
        return _team_redirect(error="err_no_such_user")
    if username == actor["username"]:
        return _team_redirect(error="err_delete_self")
    if target["is_admin"] and ctx.db.admin_count() <= 1:
        return _team_redirect(error="err_delete_last_admin")
    ctx.db.delete_user(username)
    log.info("account %r deleted by %r", username, actor["username"])
    return _team_redirect(msg="msg_user_deleted")


@bp.post("/team/password")
def team_password():
    ctx = _ctx()
    if _require_account() is None:
        return redirect(url_for("auth.login_page"), code=303)
    username = _target_username()
    if username is None:
        return _team_redirect(error="err_not_admin")
    if ctx.db.get_user(username) is None:
        return _team_redirect(error="err_no_such_user")
    password = request.form.get("password") or ""
    if not auth.valid_password(password):
        return _team_redirect(error="auth_bad_password")
    # Rotating the hash rotates the credential fingerprint: every session
    # for this account dies — including, for a self-change, the one making
    # the change. Intended (it is the revocation mechanism).
    ctx.db.update_user(username, password_hash=auth.hash_password(password))
    log.info("password changed for %r by %r", username,
             g.account["username"])
    return _team_redirect(msg="msg_pw_changed")


@bp.post("/team/email")
def team_email():
    ctx = _ctx()
    if _require_account() is None:
        return redirect(url_for("auth.login_page"), code=303)
    username = _target_username()
    if username is None:
        return _team_redirect(error="err_not_admin")
    if ctx.db.get_user(username) is None:
        return _team_redirect(error="err_no_such_user")
    email = auth.normalize_email(request.form.get("email"))
    if not email:
        ctx.db.update_user(username, email=None, email_verified=0)
        return _team_redirect(msg="msg_email_removed")
    if not auth.valid_email(email):
        return _team_redirect(error="auth_bad_email")
    if ctx.db.email_taken(email, except_username=username):
        return _team_redirect(error="auth_email_taken")
    # ALWAYS unconfirmed on change — confirmation belongs to an address,
    # not to an account, even when the same address is re-entered.
    ctx.db.update_user(username, email=email, email_verified=0)
    if mailer.enabled():
        if _send_confirmation(username, email):
            return _team_redirect(msg="msg_email_saved")
        # The address IS saved; only the confirmation email failed.
        return _team_redirect(error="msg_email_saved_sendfail")
    return _team_redirect(msg="msg_email_saved_noconfirm")


@bp.post("/team/verify-email")
def team_verify_email():
    ctx = _ctx()
    if _require_account() is None:
        return redirect(url_for("auth.login_page"), code=303)
    username = _target_username()
    if username is None:
        return _team_redirect(error="err_not_admin")
    user = ctx.db.get_user(username)
    if user is None:
        return _team_redirect(error="err_no_such_user")
    if not user["email"]:
        return _team_redirect(error="err_no_email_on_file")
    if not mailer.enabled():
        return _team_redirect(error="err_email_off")
    sent = _send_confirmation(username, user["email"])
    return (_team_redirect(msg="msg_verify_sent") if sent
            else _team_redirect(error="msg_send_failed"))
