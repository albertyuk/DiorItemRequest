"""Account, session, token, and throttling primitives for the login system.

Ported from the DMR Reconciler auth handoff: closed-team accounts, the
APP_PASSWORD secret acting as a *setup code* (creates/recovers the admin
account at /setup — nobody logs in with it), stateless signed session
cookies versioned against the password hash, single-use email tokens
stored only as hashes, and in-process failure throttling (correct for
this single-process gunicorn deployment).
"""

from __future__ import annotations

import hashlib
import hmac
import os
import re
import secrets
import sqlite3
import threading
import time
from contextlib import closing
from dataclasses import dataclass
from pathlib import Path


def now() -> float:
    """Single time source so tests can time-travel."""
    return time.time()


# --- password hashing -------------------------------------------------------

PBKDF2_ITERATIONS = 200_000


def hash_password(password: str) -> str:
    salt = secrets.token_bytes(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt,
                             PBKDF2_ITERATIONS)
    return f"{salt.hex()}${dk.hex()}"


def verify_password(password: str, stored: str) -> bool:
    """Malformed `stored` (including '') verifies False without raising, so
    an invited-but-not-activated account is inert rather than special-cased."""
    try:
        salt_hex, dk_hex = stored.split("$")
        salt = bytes.fromhex(salt_hex)
        expected = bytes.fromhex(dk_hex)
    except (AttributeError, ValueError):
        return False
    if not salt or not expected:
        return False
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt,
                             PBKDF2_ITERATIONS)
    return hmac.compare_digest(dk, expected)


# Unknown usernames verify against this, so response timing does not reveal
# whether an account exists.
DUMMY_HASH = hash_password(secrets.token_hex(16))


# --- session cookies (stateless, signed) ------------------------------------

SESSION_COOKIE = "dior_session"
SESSION_TTL = 7 * 24 * 3600


def credential_fingerprint(password_hash: str) -> str:
    """The session is versioned against the user's current password hash:
    changing the password rotates this and instantly kills all sessions."""
    return hashlib.sha256(password_hash.encode("utf-8")).hexdigest()[:16]


def make_session_cookie(username: str, password_hash: str, secret: str) -> str:
    expiry = int(now()) + SESSION_TTL
    payload = f"{username}|{expiry}|{credential_fingerprint(password_hash)}"
    sig = hmac.new(secret.encode("utf-8"), payload.encode("utf-8"),
                   hashlib.sha256).hexdigest()
    return f"{payload}|{sig}"


def read_session_cookie(value: str, secret: str) -> tuple[str, str] | None:
    """(username, credential_fp) when signature and expiry hold, else None.
    The caller must still compare the fp against the user's CURRENT hash.
    Usernames cannot contain '|' (see USERNAME_RE), so the split is
    unambiguous."""
    parts = (value or "").split("|")
    if len(parts) != 4:
        return None
    username, expiry_s, fp, sig = parts
    payload = f"{username}|{expiry_s}|{fp}"
    good = hmac.new(secret.encode("utf-8"), payload.encode("utf-8"),
                    hashlib.sha256).hexdigest()
    # Compare as bytes: compare_digest on str raises TypeError for
    # non-ASCII input, and a garbage cookie must read as anonymous — never
    # turn into a 500 on every request until the cookie is cleared.
    if not hmac.compare_digest(sig.encode("utf-8"), good.encode("utf-8")):
        return None
    try:
        expiry = int(expiry_s)
    except ValueError:
        return None
    if now() > expiry:
        return None
    return username, fp


def resolve_session(ctx: "AuthContext", cookie_value: str | None):
    """Cookie -> current user row, or None. Validation order: parse ->
    signature -> expiry -> re-fetch user and constant-time compare the
    CURRENT credential fingerprint."""
    parsed = read_session_cookie(cookie_value or "", ctx.secret)
    if parsed is None:
        return None
    username, fp = parsed
    user = ctx.db.get_user(username)
    if user is None:
        return None
    current = credential_fingerprint(user["password_hash"])
    if not hmac.compare_digest(fp.encode(), current.encode()):
        return None
    return user


# --- input validation -------------------------------------------------------

USERNAME_RE = re.compile(r"^[a-z0-9][a-z0-9._-]{1,31}$")
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

# Hard ceilings on untrusted credential input. Usernames become throttle
# bucket keys and passwords are fed to PBKDF2, so neither may be
# unbounded — truncate before either is used, well above any legitimate
# value (valid_username caps real names at 32).
MAX_USERNAME_LEN = 64
MAX_PASSWORD_LEN = 1024


def normalize_username(raw) -> str:
    return (raw or "")[:MAX_USERNAME_LEN].strip().casefold()


def clamp_password(raw) -> str:
    """Bound a submitted password before it reaches the hasher."""
    return (raw or "")[:MAX_PASSWORD_LEN]


def valid_username(username: str) -> bool:
    return bool(USERNAME_RE.match(username))


def normalize_email(raw) -> str:
    return (raw or "").strip().lower()


def valid_email(email: str) -> bool:
    return len(email) <= 254 and bool(EMAIL_RE.match(email))


def valid_password(password) -> bool:
    return MAX_PASSWORD_LEN >= len(password or "") >= 8


# --- storage ----------------------------------------------------------------

_SCHEMA = """
CREATE TABLE IF NOT EXISTS users (
    username       TEXT PRIMARY KEY,
    display        TEXT,
    password_hash  TEXT NOT NULL,
    is_admin       INTEGER DEFAULT 0,
    created_at     REAL NOT NULL,
    email          TEXT,
    email_verified INTEGER DEFAULT 0
);
CREATE TABLE IF NOT EXISTS auth_tokens (
    token_hash TEXT PRIMARY KEY,
    username   TEXT NOT NULL,
    purpose    TEXT NOT NULL,
    email      TEXT,
    created_at REAL NOT NULL,
    expires_at REAL NOT NULL,
    used_at    REAL
);
CREATE INDEX IF NOT EXISTS idx_auth_tokens_user ON auth_tokens(username);
"""

_USER_FIELDS = {"display", "password_hash", "is_admin", "email",
                "email_verified"}


class AuthDB:
    """SQLite-backed users + tokens. A connection per call keeps this safe
    across gunicorn's threads; the atomic conditional UPDATE in
    consume_token is the single-use guarantee."""

    def __init__(self, path):
        self.path = str(path)
        with closing(self._conn()) as conn, conn:
            conn.executescript(_SCHEMA)
        # Password hashes and token hashes live here: keep it owner-only
        # (defense in depth — the Fly volume is single-tenant, but a
        # stray backup or a shared host should not expose it).
        for suffix in ("", "-wal", "-shm"):
            try:
                Path(self.path + suffix).chmod(0o600)
            except OSError:
                pass

    def _conn(self):
        conn = sqlite3.connect(self.path, timeout=10)
        conn.row_factory = sqlite3.Row
        return conn

    # -- users --

    def user_count(self) -> int:
        with closing(self._conn()) as conn:
            return conn.execute("SELECT COUNT(*) FROM users").fetchone()[0]

    def admin_count(self) -> int:
        with closing(self._conn()) as conn:
            return conn.execute(
                "SELECT COUNT(*) FROM users WHERE is_admin = 1").fetchone()[0]

    def get_user(self, username: str):
        with closing(self._conn()) as conn:
            return conn.execute("SELECT * FROM users WHERE username = ?",
                                (username,)).fetchone()

    def get_user_by_email(self, email: str):
        with closing(self._conn()) as conn:
            return conn.execute(
                "SELECT * FROM users WHERE email = ? COLLATE NOCASE",
                (email,)).fetchone()

    def email_taken(self, email: str, except_username: str = "") -> bool:
        with closing(self._conn()) as conn:
            row = conn.execute(
                "SELECT username FROM users WHERE email = ? COLLATE NOCASE "
                "AND username != ?", (email, except_username)).fetchone()
            return row is not None

    def list_users(self) -> list:
        with closing(self._conn()) as conn:
            return conn.execute(
                "SELECT * FROM users ORDER BY created_at").fetchall()

    def create_user(self, username: str, password_hash: str, display=None,
                    is_admin: bool = False, email=None,
                    email_verified: bool = False) -> None:
        with closing(self._conn()) as conn, conn:
            conn.execute(
                "INSERT INTO users (username, display, password_hash, "
                "is_admin, created_at, email, email_verified) "
                "VALUES (?, ?, ?, ?, ?, ?, ?)",
                (username, display, password_hash, int(is_admin), now(),
                 email, int(email_verified)))

    def update_user(self, username: str, **fields) -> None:
        bad = set(fields) - _USER_FIELDS
        if bad:
            raise ValueError(f"unknown user fields: {bad}")
        if not fields:
            return
        assign = ", ".join(f"{k} = ?" for k in fields)
        with closing(self._conn()) as conn, conn:
            conn.execute(f"UPDATE users SET {assign} WHERE username = ?",
                         (*fields.values(), username))

    def delete_user(self, username: str) -> None:
        with closing(self._conn()) as conn, conn:
            conn.execute("DELETE FROM auth_tokens WHERE username = ?",
                         (username,))
            conn.execute("DELETE FROM users WHERE username = ?", (username,))

    def confirm_email(self, username: str, sent_to) -> bool:
        """Mark the email confirmed, scoped to the address the link was sent
        to — a stale link must never bless a newer, different address."""
        if not sent_to:
            return False
        with closing(self._conn()) as conn, conn:
            cur = conn.execute(
                "UPDATE users SET email_verified = 1 "
                "WHERE username = ? AND email = ? COLLATE NOCASE",
                (username, sent_to))
            return cur.rowcount == 1

    # -- single-use tokens --

    def issue_token(self, username: str, purpose: str, email,
                    ttl_seconds: float) -> str:
        """Returns the raw token; only its SHA-256 is stored, so a database
        read cannot mint a working link."""
        raw = secrets.token_urlsafe(32)
        token_hash = hashlib.sha256(raw.encode("utf-8")).hexdigest()
        t = now()
        with closing(self._conn()) as conn, conn:
            conn.execute("DELETE FROM auth_tokens WHERE expires_at < ?",
                         (t - 30 * 86400,))  # housekeeping
            conn.execute(
                "INSERT INTO auth_tokens (token_hash, username, purpose, "
                "email, created_at, expires_at) VALUES (?, ?, ?, ?, ?, ?)",
                (token_hash, username, purpose, email, t, t + ttl_seconds))
        return raw

    def peek_token(self, raw: str, purpose: str):
        """Non-destructive validity check (rendering a form must not burn
        the link). Same WHERE clause as consumption, read-only."""
        token_hash = hashlib.sha256((raw or "").encode("utf-8")).hexdigest()
        with closing(self._conn()) as conn:
            return conn.execute(
                "SELECT * FROM auth_tokens WHERE token_hash = ? "
                "AND purpose = ? AND used_at IS NULL AND expires_at > ?",
                (token_hash, purpose, now())).fetchone()

    def consume_token(self, raw: str, purpose: str):
        """Single atomic conditional update — exactly one concurrent click
        wins. Never SELECT-then-UPDATE."""
        token_hash = hashlib.sha256((raw or "").encode("utf-8")).hexdigest()
        t = now()
        with closing(self._conn()) as conn, conn:
            cur = conn.execute(
                "UPDATE auth_tokens SET used_at = ? WHERE token_hash = ? "
                "AND purpose = ? AND used_at IS NULL AND expires_at > ?",
                (t, token_hash, purpose, t))
            if cur.rowcount != 1:
                return None
            return conn.execute(
                "SELECT * FROM auth_tokens WHERE token_hash = ?",
                (token_hash,)).fetchone()

    def invalidate_tokens(self, username: str, purpose: str) -> None:
        """Bulk-expire a user's outstanding tokens of one purpose (e.g. all
        other reset links after a successful reset)."""
        with closing(self._conn()) as conn, conn:
            conn.execute(
                "UPDATE auth_tokens SET used_at = ? WHERE username = ? "
                "AND purpose = ? AND used_at IS NULL", (now(), username,
                                                        purpose))


# --- failure throttling -----------------------------------------------------

THROTTLE_WINDOW = 300  # sliding 5 minutes
THROTTLE_LIMITS = {"user": 5, "ip": 20, "setup": 5}
THROTTLE_MAX_KEYS = 10_000


class Throttle:
    """In-process sliding-window failure counter. reserve() atomically
    checks AND records, closing the race where concurrent guesses all pass
    a stale count while the slow hash runs."""

    def __init__(self):
        self._lock = threading.Lock()
        self._buckets: dict[tuple[str, str], list[float]] = {}

    def reserve(self, pairs) -> int:
        """Reserve one failure slot in every (scope, key) bucket. Returns 0
        on success, else the seconds to wait (nothing is recorded when
        blocked)."""
        t = now()
        with self._lock:
            self._prune(t)
            for scope, key in pairs:
                bucket = self._buckets.get((scope, key), [])
                if len(bucket) >= THROTTLE_LIMITS[scope]:
                    return max(1, int(bucket[0] + THROTTLE_WINDOW - t) + 1)
            for scope, key in pairs:
                self._buckets.setdefault((scope, key), []).append(t)
            self._evict()
            return 0

    def clear(self, scope: str, key: str) -> None:
        """Wipe one bucket (e.g. the user bucket after a correct login)."""
        with self._lock:
            self._buckets.pop((scope, key), None)

    def release(self, scope: str, key: str) -> None:
        """Pop one reserved slot (a correct login must not count as an IP
        failure — but clearing the whole IP bucket would erase other users'
        strikes from a shared NAT)."""
        with self._lock:
            bucket = self._buckets.get((scope, key))
            if bucket:
                bucket.pop()
                if not bucket:
                    del self._buckets[(scope, key)]

    def _prune(self, t: float) -> None:
        for key in list(self._buckets):
            kept = [x for x in self._buckets[key] if x > t - THROTTLE_WINDOW]
            if kept:
                self._buckets[key] = kept
            else:
                del self._buckets[key]

    def _evict(self) -> None:
        while len(self._buckets) > THROTTLE_MAX_KEYS:
            oldest = min(self._buckets, key=lambda k: self._buckets[k][-1])
            del self._buckets[oldest]


# --- app-level context ------------------------------------------------------

@dataclass
class AuthContext:
    db: AuthDB
    throttle: Throttle
    setup_code: str          # APP_PASSWORD: creates/recovers the admin only
    open_access: bool        # explicit local no-auth opt-out
    secret: str              # session signing key (never the setup code)
    cookie_secure: bool
    invite_ttl_hours: int = 72
    reset_ttl_hours: int = 2
    public_base_url: str = ""
    # Only an edge that SETS and STRIPS the client-IP header may be
    # believed. Off that edge the header is client-controlled, and
    # trusting it would let one attacker mint a fresh throttle bucket per
    # request — defeating per-IP rate limiting entirely.
    trust_ip_header: bool = False


def load_secret(data_dir, configured: str | None = None) -> str:
    """APP_SECRET, or a generated secret persisted to DATA_DIR (0600) so
    sessions survive restarts. Deliberately NOT derived from the setup
    code — that secret is shared with teammates and must not let its
    holders forge cookies."""
    if configured:
        return configured
    env = os.environ.get("APP_SECRET", "")
    if env:
        return env
    path = Path(data_dir) / "session_secret"
    try:
        stored = path.read_text().strip()
        if stored:
            return stored
    except OSError:
        pass
    secret = secrets.token_hex(32)
    path.write_text(secret)
    path.chmod(0o600)
    return secret
