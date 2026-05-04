"""Claude-based email classification.

One LLM call per run, batched. Returns a Classification per message.
"""
from __future__ import annotations

import json
from dataclasses import dataclass
from typing import Literal

from anthropic import Anthropic

from .config import AgentConfig
from .gmail_client import Message

# Approximate Sonnet 4.6 pricing per million tokens (input/output).
# Used for the daily spend cap. Adjust if pricing changes.
SONNET_INPUT_USD_PER_MTOK = 3.0
SONNET_OUTPUT_USD_PER_MTOK = 15.0

Category = Literal[
    "advertisement",
    "job_application_response",
    "recruiter_cold_outreach",
    "interview_request",
    "security_alert",
    "bill_due",
    "calendar_invite",
    "personal_reply_needed",
    "urgent",
    "other",
]


@dataclass
class Classification:
    message_id: str
    primary: Category
    secondary: list[Category]
    confidence: float
    reasoning: str
    interview_proposed_time: str | None  # human-readable, parsed downstream
    interview_format_hint: str | None    # zoom/teams/in-person/phone/None
    bill_due_date: str | None
    is_advertisement: bool
    is_security: bool
    notify_sms: bool


SYSTEM_PROMPT = """You are an email triage assistant. For each email, classify it into ONE primary category and zero or more secondary categories.

Primary categories (pick exactly one):
- advertisement: marketing/promotional. Has List-Unsubscribe header AND in Promotions category.
- job_application_response: a company or recruiter responding to an application the user submitted (interview invitation counts here only if no specific time is proposed).
- recruiter_cold_outreach: a recruiter or headhunter reaching out about a role the user did NOT apply to.
- interview_request: a request to schedule a specific interview (zoom/teams/in-person/phone), typically proposing a time or asking for availability.
- security_alert: account login alerts, password reset codes, 2FA notices, "new device" notifications from real services.
- bill_due: invoice or bill with a payment due date.
- calendar_invite: an .ics calendar invitation that is NOT an interview.
- personal_reply_needed: a real human (not marketing) is directly asking the user something and expects a reply.
- urgent: contains explicit time pressure language (deadline, EOD, urgent) outside the above categories.
- other: anything else.

For interview_request, also extract:
- interview_proposed_time: a date/time mentioned in the email (e.g., "Thursday May 8 at 2pm"), or null.
- interview_format_hint: one of "zoom", "teams", "in-person", "phone", or null.

For bill_due, also extract:
- bill_due_date: ISO date if present, else null.

Set is_advertisement=true ONLY when primary=advertisement.
Set is_security=true ONLY when primary=security_alert.
Set notify_sms=true when the message warrants an SMS digest entry: job_application_response, recruiter_cold_outreach, interview_request, security_alert, urgent, or bill_due (within 3 days).

Return ONLY a JSON array of objects, no prose, one object per input email, in the same order. Each object has these fields:
{ "message_id": str, "primary": str, "secondary": [str], "confidence": float (0-1), "reasoning": str (one sentence), "interview_proposed_time": str|null, "interview_format_hint": str|null, "bill_due_date": str|null, "is_advertisement": bool, "is_security": bool, "notify_sms": bool }
"""


def _format_message(m: Message) -> dict:
    return {
        "message_id": m.id,
        "from": f"{m.sender_name} <{m.sender_email}>",
        "subject": m.subject,
        "received_utc": m.received_utc.isoformat(),
        "has_list_unsubscribe": m.has_list_unsubscribe,
        "in_promotions_category": m.in_promotions,
        "snippet": m.snippet,
        "body": m.body_text[:4000],
    }


def classify(
    messages: list[Message], cfg: AgentConfig
) -> tuple[list[Classification], float]:
    """Classify a batch. Returns (classifications, usd_spent)."""
    if not messages:
        return [], 0.0

    client = Anthropic(api_key=cfg.anthropic_api_key)
    payload = [_format_message(m) for m in messages]
    user_msg = "Classify these emails:\n\n" + json.dumps(payload, indent=2)

    resp = client.messages.create(
        model=cfg.anthropic_model,
        max_tokens=4096,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_msg}],
    )
    text = "".join(b.text for b in resp.content if hasattr(b, "text"))
    usd = (
        resp.usage.input_tokens * SONNET_INPUT_USD_PER_MTOK / 1_000_000
        + resp.usage.output_tokens * SONNET_OUTPUT_USD_PER_MTOK / 1_000_000
    )
    parsed = _safe_json_parse(text)
    classifications: list[Classification] = []
    by_id = {p.get("message_id"): p for p in parsed}
    for m in messages:
        p = by_id.get(m.id) or _fallback(m.id)
        classifications.append(_to_classification(p))
    return classifications, usd


def _safe_json_parse(text: str) -> list[dict]:
    text = text.strip()
    # Strip optional ```json fences
    if text.startswith("```"):
        text = text.split("```", 2)[1]
        if text.startswith("json"):
            text = text[4:]
        text = text.rsplit("```", 1)[0]
    try:
        data = json.loads(text)
        return data if isinstance(data, list) else []
    except json.JSONDecodeError:
        return []


def _fallback(mid: str) -> dict:
    return {
        "message_id": mid,
        "primary": "other",
        "secondary": [],
        "confidence": 0.0,
        "reasoning": "Classifier did not return a result; defaulting to 'other'.",
        "interview_proposed_time": None,
        "interview_format_hint": None,
        "bill_due_date": None,
        "is_advertisement": False,
        "is_security": False,
        "notify_sms": False,
    }


def _to_classification(p: dict) -> Classification:
    return Classification(
        message_id=p.get("message_id", ""),
        primary=p.get("primary", "other"),
        secondary=p.get("secondary", []) or [],
        confidence=float(p.get("confidence", 0.0)),
        reasoning=p.get("reasoning", ""),
        interview_proposed_time=p.get("interview_proposed_time"),
        interview_format_hint=p.get("interview_format_hint"),
        bill_due_date=p.get("bill_due_date"),
        is_advertisement=bool(p.get("is_advertisement", False)),
        is_security=bool(p.get("is_security", False)),
        notify_sms=bool(p.get("notify_sms", False)),
    )
