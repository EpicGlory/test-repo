"""Heuristic-only email classification (no LLM, no API cost).

Trades off some accuracy on edge cases for $0/month operation. The hardest
class is `interview_request` vs `job_application_response` — we use the
presence of scheduling-ish language to disambiguate.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Literal

from .config import AgentConfig
from .gmail_client import Message

Category = Literal[
    "advertisement",
    "job_application_response",
    "recruiter_cold_outreach",
    "interview_request",
    "security_alert",
    "bill_due",
    "calendar_invite",
    "urgent",
    "other",
]


INTERVIEW_KEYWORDS = {
    "interview",
    "phone screen", "phone call", "screening call",
    "zoom call", "zoom meeting", "teams meeting", "video call",
    "schedule a call", "set up a call", "set up a time",
    "available to chat", "good time to chat",
    "schedule a meeting", "schedule a chat",
    "30 minutes", "30-minute", "half hour",
    "calendly", "book a time",
}

JOB_KEYWORDS = {
    "position", "opportunity", "the role", "this role",
    "candidate", "applicant", "application",
    "you applied", "your application",
    "recruiter", "recruiting", "talent acquisition",
    "hiring", "hiring manager",
    "next steps", "next step in the process",
    "your resume", "your background", "your experience",
}

RECRUITER_COLD_PHRASES = {
    "came across your",
    "your profile caught", "your background caught",
    "your linkedin",
    "would you be open", "would you be interested",
    "reaching out because", "i'm a recruiter", "talent partner",
    "looking for someone", "looking to fill",
}

SECURITY_KEYWORDS = {
    "security alert",
    "new sign-in", "new sign in", "new device",
    "verify it's you", "verify it was you",
    "verification code", "your code is",
    "2-step verification", "two-factor", "two factor",
    "password reset", "reset your password",
    "unusual activity", "suspicious activity",
    "login attempt", "sign-in attempt",
}

BILL_KEYWORDS = {
    "amount due", "payment due", "balance due",
    "invoice", "your statement", "monthly statement",
    "your bill", "autopay",
    "minimum payment",
}

URGENT_KEYWORDS = {
    "urgent", "asap", "as soon as possible",
    "by eod", "end of day", "end of business",
    "deadline", "time-sensitive", "time sensitive",
    "by tomorrow", "by today",
}

CAL_INVITE_KEYWORDS = {
    "you're invited",
    "calendar invite", "you have been invited",
    "view event", "rsvp",
}

INTERVIEW_FORMAT_PATTERNS = [
    ("zoom", re.compile(r"\bzoom\b", re.I)),
    ("teams", re.compile(r"\b(microsoft\s+)?teams\b", re.I)),
    ("in-person", re.compile(r"\b(in[- ]person|on[- ]site|in our office)\b", re.I)),
    ("phone", re.compile(r"\b(phone call|by phone|over the phone|call you)\b", re.I)),
]


@dataclass
class Classification:
    message_id: str
    primary: Category
    secondary: list[Category]
    confidence: float
    reasoning: str
    interview_proposed_time: str | None
    interview_format_hint: str | None
    bill_due_date: str | None
    is_advertisement: bool
    is_security: bool
    notify_sms: bool


def _hits(text_lower: str, keywords: set[str]) -> list[str]:
    return [k for k in keywords if k in text_lower]


def _detect_format(text: str) -> str | None:
    for label, pattern in INTERVIEW_FORMAT_PATTERNS:
        if pattern.search(text):
            return label
    return None


def _extract_due_date(text: str) -> str | None:
    pattern = re.compile(
        r"due\s+(?:by\s+|on\s+)?([A-Za-z]+\s+\d{1,2}(?:,\s*\d{4})?)",
        re.IGNORECASE,
    )
    m = pattern.search(text)
    return m.group(1) if m else None


def _due_within_3_days(due_str: str | None) -> bool:
    if not due_str:
        return False
    for fmt in ("%B %d, %Y", "%B %d", "%b %d, %Y", "%b %d"):
        try:
            dt = datetime.strptime(due_str, fmt)
            if dt.year == 1900:
                dt = dt.replace(year=datetime.now(timezone.utc).year)
            days = (dt - datetime.now()).days
            return 0 <= days <= 3
        except ValueError:
            continue
    return False


def _classify_one(m: Message) -> Classification:
    text = f"{m.subject}\n{m.snippet}\n{m.body_text}"
    text_l = text.lower()

    # 1. Advertisement — pure logic, very high precision
    if m.has_list_unsubscribe and m.in_promotions:
        return Classification(
            message_id=m.id, primary="advertisement", secondary=[],
            confidence=0.99,
            reasoning="List-Unsubscribe header + Promotions category",
            interview_proposed_time=None, interview_format_hint=None,
            bill_due_date=None,
            is_advertisement=True, is_security=False, notify_sms=False,
        )

    # 2. Security alert
    sec_hits = _hits(text_l, SECURITY_KEYWORDS)
    if sec_hits:
        return Classification(
            message_id=m.id, primary="security_alert", secondary=[],
            confidence=0.85,
            reasoning=f"Security keywords: {', '.join(sec_hits[:3])}",
            interview_proposed_time=None, interview_format_hint=None,
            bill_due_date=None,
            is_advertisement=False, is_security=True, notify_sms=True,
        )

    interview_hits = _hits(text_l, INTERVIEW_KEYWORDS)
    job_hits = _hits(text_l, JOB_KEYWORDS)
    cold_hits = _hits(text_l, RECRUITER_COLD_PHRASES)

    # 3. Interview request: scheduling-ish language + job context
    if interview_hits and (job_hits or cold_hits):
        return Classification(
            message_id=m.id, primary="interview_request", secondary=[],
            confidence=0.8,
            reasoning=f"Interview+context keywords: {', '.join((interview_hits + job_hits)[:3])}",
            interview_proposed_time=None,
            interview_format_hint=_detect_format(text),
            bill_due_date=None,
            is_advertisement=False, is_security=False, notify_sms=True,
        )

    # 4. Cold recruiter outreach
    if cold_hits and (job_hits or "linkedin" in text_l):
        return Classification(
            message_id=m.id, primary="recruiter_cold_outreach", secondary=[],
            confidence=0.7,
            reasoning=f"Cold-outreach phrases: {', '.join(cold_hits[:2])}",
            interview_proposed_time=None, interview_format_hint=None,
            bill_due_date=None,
            is_advertisement=False, is_security=False, notify_sms=True,
        )

    # 5. Generic job-application response
    if job_hits:
        return Classification(
            message_id=m.id, primary="job_application_response", secondary=[],
            confidence=0.6,
            reasoning=f"Job keywords: {', '.join(job_hits[:3])}",
            interview_proposed_time=None, interview_format_hint=None,
            bill_due_date=None,
            is_advertisement=False, is_security=False, notify_sms=True,
        )

    # 6. Bills
    bill_hits = _hits(text_l, BILL_KEYWORDS)
    if bill_hits:
        due = _extract_due_date(text)
        return Classification(
            message_id=m.id, primary="bill_due", secondary=[],
            confidence=0.7,
            reasoning=f"Bill keywords: {', '.join(bill_hits[:2])}",
            interview_proposed_time=None, interview_format_hint=None,
            bill_due_date=due,
            is_advertisement=False, is_security=False,
            notify_sms=_due_within_3_days(due),
        )

    # 7. Calendar invite (non-interview)
    if _hits(text_l, CAL_INVITE_KEYWORDS) or \
            "text/calendar" in (m.raw_headers.get("content-type") or "").lower():
        return Classification(
            message_id=m.id, primary="calendar_invite", secondary=[],
            confidence=0.8,
            reasoning="Calendar invite indicators",
            interview_proposed_time=None, interview_format_hint=None,
            bill_due_date=None,
            is_advertisement=False, is_security=False, notify_sms=False,
        )

    # 8. Urgent
    urgent_hits = _hits(text_l, URGENT_KEYWORDS)
    if urgent_hits:
        return Classification(
            message_id=m.id, primary="urgent", secondary=[],
            confidence=0.6,
            reasoning=f"Urgent keywords: {', '.join(urgent_hits[:2])}",
            interview_proposed_time=None, interview_format_hint=None,
            bill_due_date=None,
            is_advertisement=False, is_security=False, notify_sms=True,
        )

    # 9. Default
    return Classification(
        message_id=m.id, primary="other", secondary=[],
        confidence=0.5,
        reasoning="No matching heuristics",
        interview_proposed_time=None, interview_format_hint=None,
        bill_due_date=None,
        is_advertisement=False, is_security=False, notify_sms=False,
    )


def classify(messages: list[Message], cfg: AgentConfig) -> list[Classification]:
    """Heuristic classification — no API calls, no cost."""
    return [_classify_one(m) for m in messages]
