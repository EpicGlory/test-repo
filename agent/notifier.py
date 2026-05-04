"""SMS via Verizon email gateway + daily summary email via Gmail SMTP."""
from __future__ import annotations

import smtplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from .classifier import Classification
from .config import AgentConfig
from .gmail_client import Message


def _smtp_send(cfg: AgentConfig, *, to: str, subject: str, body: str, html: str | None = None) -> None:
    msg = MIMEMultipart("alternative") if html else MIMEMultipart()
    msg["From"] = cfg.smtp_user
    msg["To"] = to
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain", "utf-8"))
    if html:
        msg.attach(MIMEText(html, "html", "utf-8"))
    # App passwords accept whitespace-stripped form.
    pw = cfg.smtp_pass.replace(" ", "")
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=30) as s:
        s.login(cfg.smtp_user, pw)
        s.sendmail(cfg.smtp_user, [to], msg.as_string())


def send_sms_test(cfg: AgentConfig, body: str = "gmail-auto-agent test") -> None:
    body = body[: cfg.sms_max_bytes]
    _smtp_send(cfg, to=cfg.sms_to, subject="", body=body)


def send_sms_digest(
    cfg: AgentConfig,
    items: list[tuple[Message, Classification]],
    *,
    urgent: bool = False,
) -> None:
    """Build a single 160-char digest in priority order."""
    if not items:
        return
    # Priority order: security > urgent > interview > job_app > recruiter > bill
    priority = {
        "security_alert": 0,
        "urgent": 1,
        "interview_request": 2,
        "job_application_response": 3,
        "recruiter_cold_outreach": 4,
        "bill_due": 5,
    }
    items_sorted = sorted(items, key=lambda mc: priority.get(mc[1].primary, 99))
    counts: dict[str, int] = {}
    for _, c in items_sorted:
        counts[c.primary] = counts.get(c.primary, 0) + 1

    parts: list[str] = []
    if urgent:
        parts.append("URGENT:")
    if cfg.dry_run:
        parts.append("[DRYRUN]")
    label_short = {
        "security_alert": "sec",
        "urgent": "urg",
        "interview_request": "intvw",
        "job_application_response": "job",
        "recruiter_cold_outreach": "rcrtr",
        "bill_due": "bill",
    }
    for cat, n in counts.items():
        parts.append(f"{n} {label_short.get(cat, cat)}")
    parts.append("- check Gmail")

    body = " ".join(parts)
    body = body[: cfg.sms_max_bytes]
    _smtp_send(cfg, to=cfg.sms_to, subject="", body=body)


def send_morning_summary(
    cfg: AgentConfig,
    *,
    yesterday_counts: dict[str, int],
    bills: list[tuple[Message, Classification]],
    reply_needed: list[tuple[Message, Classification]],
    urgent: list[tuple[Message, Classification]],
    invites: list[tuple[Message, Classification]],
    spend_yesterday_usd: float,
    spend_month_to_date_usd: float,
    mistakes: list[Message],
) -> None:
    when = datetime.now().strftime("%A, %B %d, %Y")
    lines = [
        f"Morning summary — {when}",
        "",
        ("=== ACTIVITY (yesterday + overnight) ==="),
    ]
    if not yesterday_counts:
        lines.append("  (nothing to report)")
    else:
        for k, v in yesterday_counts.items():
            lines.append(f"  {k}: {v}")
    lines.append("")

    def render_section(title: str, items: list[tuple[Message, Classification]], extra=None):
        lines.append(f"=== {title} ({len(items)}) ===")
        if not items:
            lines.append("  (none)")
        else:
            for m, c in items[:20]:
                tail = ""
                if extra == "bill_date" and c.bill_due_date:
                    tail = f" — due {c.bill_due_date}"
                lines.append(f"  • {m.sender_name or m.sender_email}: {m.subject}{tail}")
        lines.append("")

    render_section("BILLS DUE", bills, extra="bill_date")
    render_section("REPLY NEEDED", reply_needed)
    render_section("URGENT / TIME-SENSITIVE", urgent)
    render_section("CALENDAR INVITES (non-interview)", invites)

    lines.append("=== MISTAKE-TAGGED LAST 24H ===")
    if not mistakes:
        lines.append("  (none — agent calls all looked correct)")
    else:
        for m in mistakes[:20]:
            lines.append(f"  • {m.sender_name or m.sender_email}: {m.subject}")
    lines.append("")

    lines.append("=== SPEND ===")
    lines.append(f"  Anthropic spend yesterday: ${spend_yesterday_usd:.4f}")
    lines.append(f"  Month-to-date estimate:    ${spend_month_to_date_usd:.4f}")
    lines.append(f"  Daily cap:                 ${cfg.daily_spend_cap_usd:.2f}")
    lines.append("")
    if cfg.dry_run:
        lines.append("MODE: DRY RUN — no real trash/star/draft/SMS were issued.")
    else:
        lines.append("MODE: LIVE")

    body = "\n".join(lines)
    subject = f"[gmail-agent] Morning summary — {when}"
    _smtp_send(cfg, to=cfg.owner_email, subject=subject, body=body)
