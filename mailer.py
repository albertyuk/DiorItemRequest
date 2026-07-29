"""Transactional email via the Resend HTTP API (stdlib only).

Three messages — invite, confirm-address, password-reset — built from one
bilingual layout (the team is mixed EN/ZH, and the sender cannot know the
reader's preference, so every body carries both languages). send() never
raises: it returns a bool and every caller has a degraded path for False.
"""

from __future__ import annotations

import html
import json
import logging
import os
import urllib.error
import urllib.request

log = logging.getLogger(__name__)

RESEND_URL = "https://api.resend.com/emails"
TIMEOUT = 12.0
DEFAULT_FROM = "Dior item request <onboarding@resend.dev>"


def enabled() -> bool:
    """Every email feature (invites, confirmation, reset) gates on this."""
    return bool(os.environ.get("RESEND_API_KEY"))


def send(to: str, subject: str, html_body: str, text_body: str) -> bool:
    """One transactional message. Returns False on ANY failure — callers
    flash an error or fall back to manual flows, never crash."""
    key = os.environ.get("RESEND_API_KEY", "")
    if not key:
        return False
    payload = json.dumps({
        "from": os.environ.get("EMAIL_FROM", DEFAULT_FROM),
        "to": [to],
        "subject": subject,
        "html": html_body,
        "text": text_body,
    }).encode("utf-8")
    req = urllib.request.Request(
        RESEND_URL, data=payload, method="POST",
        headers={"Authorization": f"Bearer {key}",
                 "Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=TIMEOUT) as resp:
            ok = 200 <= resp.status < 300
    except Exception as exc:
        # Deliberately broad: send() must NEVER raise — urlopen can throw
        # beyond URLError/OSError (e.g. http.client.BadStatusLine), and any
        # escape here would turn a caller's degraded path into a 500.
        log.warning("email send to %r failed: %s", to, exc)
        return False
    if not ok:
        log.warning("email send to %r failed: HTTP %s", to, resp.status)
    return ok


def _layout(title_en: str, title_zh: str, body_en: str, body_zh: str,
            button: str, link: str, hours: int) -> tuple[str, str]:
    """(html, text). All interpolated values must be pre-escaped by the
    caller except the link, which is escaped here."""
    safe_link = html.escape(link, quote=True)
    note_en = f"The link works once and expires in {hours} hours."
    note_zh = f"链接仅可使用一次，{hours} 小时后失效。"
    html_body = f"""\
<div style="font-family:system-ui,sans-serif;max-width:32rem;margin:0 auto;
            padding:1.5rem;color:#1b1f24">
  <p style="font-weight:700;font-size:1.05rem">Dior item request</p>
  <p><strong>{title_en}</strong><br>{body_en}</p>
  <p><strong>{title_zh}</strong><br>{body_zh}</p>
  <p style="margin:1.4rem 0">
    <a href="{safe_link}" style="background:#2c7a44;color:#fff;padding:.6rem
       1.4rem;border-radius:8px;text-decoration:none">{button}</a></p>
  <p style="color:#69737d;font-size:.9rem">{note_en}<br>{note_zh}</p>
  <p style="color:#69737d;font-size:.85rem;word-break:break-all">{safe_link}</p>
</div>"""
    text_body = (f"Dior item request\n\n{title_en}\n{body_en}\n\n{title_zh}\n"
                 f"{body_zh}\n\n{link}\n\n{note_en}\n{note_zh}\n")
    return html_body, text_body


def send_invite(to: str, link: str, inviter: str, hours: int) -> bool:
    who = html.escape(inviter or "An admin")
    html_body, text_body = _layout(
        "You have been invited",
        "邀请加入",
        f"{who} invited you to the Dior item request tool. Click the button "
        "to choose your password and activate your account.",
        f"{who} 邀请你使用 Dior 补货申请工具。点击下方按钮设置密码并激活账号。",
        "Choose password / 设置密码", link, hours)
    return send(to, "You're invited — Dior item request / 邀请加入",
                html_body, text_body)


def send_confirm(to: str, link: str, hours: int) -> bool:
    html_body, text_body = _layout(
        "Confirm your email address",
        "确认邮箱地址",
        "Confirm this address to enable password reset for your account.",
        "确认此邮箱后，你的账号才能使用“忘记密码”功能。",
        "Confirm / 确认", link, hours)
    return send(to, "Confirm your email — Dior item request / 确认邮箱",
                html_body, text_body)


def send_reset(to: str, link: str, hours: int) -> bool:
    html_body, text_body = _layout(
        "Reset your password",
        "重置密码",
        "Click the button to choose a new password. If you did not ask for "
        "this, you can ignore this email.",
        "点击下方按钮设置新密码。如果这不是你本人的操作，请忽略本邮件。",
        "Choose new password / 设置新密码", link, hours)
    return send(to, "Reset your password — Dior item request / 重置密码",
                html_body, text_body)
