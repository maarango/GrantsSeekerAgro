"""
Sending the weekly digest through any email provider (SMTP).

The SMTP password is stored in Windows Credential Manager under
GrantsSeekerAgro/smtp:<username>, never in config files. The environment
variable GRANTSSEEKER_SMTP_PASSWORD takes precedence when set.

Most big providers require an *app password* (not your normal password) for
SMTP when two-step verification is on; see APP_PASSWORD_HELP.
"""
from __future__ import annotations

import os
import smtplib
import ssl
from email.message import EmailMessage
from email.utils import formatdate, make_msgid
from pathlib import Path

import ai_providers as ai

PASSWORD_ENV = "GRANTSSEEKER_SMTP_PASSWORD"

# domain -> (host, port, security)
PRESETS = {
    "gmail.com": ("smtp.gmail.com", 587, "starttls"),
    "googlemail.com": ("smtp.gmail.com", 587, "starttls"),
    "outlook.com": ("smtp-mail.outlook.com", 587, "starttls"),
    "hotmail.com": ("smtp-mail.outlook.com", 587, "starttls"),
    "live.com": ("smtp-mail.outlook.com", 587, "starttls"),
    "msn.com": ("smtp-mail.outlook.com", 587, "starttls"),
    "yahoo.com": ("smtp.mail.yahoo.com", 465, "ssl"),
    "ymail.com": ("smtp.mail.yahoo.com", 465, "ssl"),
    "icloud.com": ("smtp.mail.me.com", 587, "starttls"),
    "me.com": ("smtp.mail.me.com", 587, "starttls"),
    "mac.com": ("smtp.mail.me.com", 587, "starttls"),
    "aol.com": ("smtp.aol.com", 465, "ssl"),
    "zoho.com": ("smtp.zoho.com", 465, "ssl"),
    "zohomail.com": ("smtp.zoho.com", 465, "ssl"),
    "gmx.com": ("mail.gmx.com", 587, "starttls"),
    "gmx.net": ("mail.gmx.net", 587, "starttls"),
    "proton.me": ("smtp.protonmail.ch", 587, "starttls"),
    "protonmail.com": ("smtp.protonmail.ch", 587, "starttls"),
}

APP_PASSWORD_HELP = {
    "smtp.gmail.com": "Gmail needs an app password: https://myaccount.google.com/apppasswords",
    "smtp.mail.yahoo.com": "Yahoo needs an app password: Account security > Generate app password",
    "smtp.mail.me.com": "iCloud needs an app-specific password: https://account.apple.com > Sign-In and Security",
    "smtp.aol.com": "AOL needs an app password: Account security > Generate app password",
    "smtp-mail.outlook.com": "Microsoft accounts may block SMTP passwords; if login fails, use the Outlook method",
    "smtp.office365.com": "Microsoft 365 often disables SMTP AUTH; if login fails, use the Outlook method",
    "smtp.protonmail.ch": "Proton needs a paid plan and an SMTP token (Settings > IMAP/SMTP)",
}

SECURITY = ("starttls", "ssl", "none")


def preset_for(address: str) -> tuple[str, int, str] | None:
    return PRESETS.get(address.rsplit("@", 1)[-1].lower())


def _cred_name(username: str) -> str:
    return f"smtp:{username.lower()}"


def get_password(username: str) -> str | None:
    return os.environ.get(PASSWORD_ENV) or ai.get_secret(_cred_name(username))


def set_password(username: str, password: str) -> None:
    ai.set_secret(_cred_name(username), password, "SMTP password used by GrantsSeekerAgro")


def _connect(smtp_cfg: dict, password: str | None) -> smtplib.SMTP:
    host, port = smtp_cfg["host"], int(smtp_cfg["port"])
    security = smtp_cfg.get("security", "starttls")
    ctx = ssl.create_default_context()
    if security == "ssl":
        server: smtplib.SMTP = smtplib.SMTP_SSL(host, port, context=ctx, timeout=60)
    else:
        server = smtplib.SMTP(host, port, timeout=60)
        server.ehlo()
        if security == "starttls":
            server.starttls(context=ctx)
            server.ehlo()
    if smtp_cfg.get("username"):
        if not password:
            server.quit()
            raise RuntimeError(f"no SMTP password stored for {smtp_cfg['username']}; run setup again "
                               f"or set {PASSWORD_ENV}")
        server.login(smtp_cfg["username"], password)
    return server


def test_login(smtp_cfg: dict, password: str | None = None) -> tuple[bool, str]:
    """Connect and authenticate without sending anything."""
    try:
        server = _connect(smtp_cfg, password or get_password(smtp_cfg.get("username", "")))
        server.quit()
        return True, "ok"
    except Exception as e:  # noqa: BLE001
        return False, str(e)


def build_message(sender: str, to: list[str], cc: list[str], subject: str, html_body: str,
                  attachment: Path | None = None, sender_name: str = "") -> EmailMessage:
    msg = EmailMessage()
    msg["From"] = f"{sender_name} <{sender}>" if sender_name else sender
    msg["To"] = ", ".join(to)
    if cc:
        msg["Cc"] = ", ".join(cc)
    msg["Subject"] = subject
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain=sender.rsplit("@", 1)[-1])
    msg.set_content("Este resumen está en formato HTML. Ábralo en un cliente de correo que muestre HTML.")
    msg.add_alternative(html_body, subtype="html")
    if attachment and attachment.exists():
        msg.add_attachment(attachment.read_bytes(), maintype="application",
                           subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           filename=attachment.name)
    return msg


def send(smtp_cfg: dict, msg: EmailMessage) -> None:
    server = _connect(smtp_cfg, get_password(smtp_cfg.get("username", "")))
    try:
        server.send_message(msg)
    finally:
        server.quit()
