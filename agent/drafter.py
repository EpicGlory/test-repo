"""Generate interview reply drafts using calendar availability + tone sample."""
from __future__ import annotations

from datetime import datetime
from zoneinfo import ZoneInfo

from anthropic import Anthropic

from .calendar_client import CalendarClient, TimeSlot
from .classifier import Classification, SONNET_INPUT_USD_PER_MTOK, SONNET_OUTPUT_USD_PER_MTOK
from .config import AgentConfig
from .gmail_client import Message


DRAFTER_SYSTEM = """You write professional, cordial email replies to interview invitations.

Match the tone of the provided sample. Keep it concise (3–5 short paragraphs max). Always:
- Thank them and express interest in the role.
- If the calendar shows the proposed time is FREE, accept it directly.
- If the proposed time conflicts OR no specific time was offered, propose the three given alternative slots.
- Confirm the format (zoom/teams/phone/in-person) if mentioned, or ask which they prefer if not.
- Sign off using the same signature pattern as the sample.

Return ONLY the email body text — no subject line, no headers, no markdown. The first line of your output must be a one-line reasoning header in this exact format:
[AGENT NOTE: drafted because <reason>; calendar checked; tone matched to sample. Delete this line before sending.]

Then a blank line, then the email body.
"""


def _try_parse_proposed_time(text: str | None, tz_name: str) -> datetime | None:
    if not text:
        return None
    # Best-effort: look for ISO-like substrings; if none, return None.
    # We deliberately don't use a heavy NLP date parser here; the LLM
    # downstream can still propose alternatives.
    try:
        return datetime.fromisoformat(text)
    except (ValueError, TypeError):
        return None


def draft_interview_reply(
    *,
    msg: Message,
    classification: Classification,
    cfg: AgentConfig,
    calendar: CalendarClient,
) -> tuple[str, float]:
    """Return (reply_body_text, usd_spent)."""
    tz_name = cfg.working_hours.timezone
    proposed_dt = _try_parse_proposed_time(classification.interview_proposed_time, tz_name)

    slots = calendar.find_free_slots(
        tz_name=tz_name,
        earliest_hour=cfg.working_hours.interview_earliest_hour,
        latest_hour=cfg.working_hours.interview_latest_hour,
        slot_minutes=cfg.interview_slot_minutes,
        buffer_minutes=cfg.interview_buffer_minutes,
        days_ahead=cfg.interview_search_days_ahead,
        slots_wanted=cfg.interview_slot_options,
        weekdays_only=cfg.working_hours.interview_weekdays_only,
        prefer_around=proposed_dt,
    )
    slot_strings = [s.format_for_email(tz_name) for s in slots]

    proposed_is_free = _is_free(proposed_dt, slots) if proposed_dt else False

    user_msg = f"""Tone sample (match this style and signature):
---
{cfg.tone_sample}
---

Incoming email:
From: {msg.sender_name} <{msg.sender_email}>
Subject: {msg.subject}
Body:
{msg.body_text[:3000]}

Classifier extracted:
- Proposed time: {classification.interview_proposed_time or "(none)"}
- Format hint: {classification.interview_format_hint or "(none)"}
- Reasoning: {classification.reasoning}

Calendar status:
- Proposed time is {"FREE" if proposed_is_free else "NOT FREE or NOT GIVEN"}.
- Three alternative free slots: {slot_strings}

Write the reply now."""

    client = Anthropic(api_key=cfg.anthropic_api_key)
    resp = client.messages.create(
        model=cfg.anthropic_model,
        max_tokens=1024,
        system=DRAFTER_SYSTEM,
        messages=[{"role": "user", "content": user_msg}],
    )
    body = "".join(b.text for b in resp.content if hasattr(b, "text")).strip()
    usd = (
        resp.usage.input_tokens * SONNET_INPUT_USD_PER_MTOK / 1_000_000
        + resp.usage.output_tokens * SONNET_OUTPUT_USD_PER_MTOK / 1_000_000
    )
    return body, usd


def _is_free(dt: datetime | None, slots: list[TimeSlot]) -> bool:
    if not dt or not slots:
        return False
    # Treat "free" as: dt falls within ±15 min of any returned free slot start.
    target = dt
    if target.tzinfo is None:
        target = target.replace(tzinfo=ZoneInfo("UTC"))
    for s in slots:
        s_start = s.start_local.astimezone(target.tzinfo)
        delta = abs((s_start - target).total_seconds())
        if delta <= 15 * 60:
            return True
    return False
