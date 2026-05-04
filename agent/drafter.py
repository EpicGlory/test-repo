"""Template-based interview reply drafts (no LLM, no API cost).

The draft is intentionally generic — it doesn't reference the specific role
or company because we have no way to extract those reliably without an LLM.
The three calendar-aware time slots are the real value here; the rest is
keystroke-saving boilerplate that you'll customize before sending.
"""
from __future__ import annotations

from .calendar_client import CalendarClient
from .classifier import Classification
from .config import AgentConfig
from .gmail_client import Message


TEMPLATE = """[AGENT NOTE: drafted from template; calendar checked. Delete this line and customize before sending.]

Hi {first_name},

Thank you for reaching out — I'd love to learn more about this opportunity.

Here are three times that work for me:
{slots}

If none of those work, just let me know what's good for you{format_line}.

Best,
Ryan Tucker
(801) 358-7820
rptucker@gmail.com
"""


def _first_name(sender_name: str, sender_email: str) -> str:
    if sender_name:
        first = sender_name.strip().split()[0].strip(",.\"'")
        if first:
            return first
    local = sender_email.split("@")[0]
    return local.split(".")[0].capitalize() or "there"


def draft_interview_reply(
    *,
    msg: Message,
    classification: Classification,
    cfg: AgentConfig,
    calendar: CalendarClient,
) -> str:
    slots = calendar.find_free_slots(
        tz_name=cfg.working_hours.timezone,
        earliest_hour=cfg.working_hours.interview_earliest_hour,
        latest_hour=cfg.working_hours.interview_latest_hour,
        slot_minutes=cfg.interview_slot_minutes,
        buffer_minutes=cfg.interview_buffer_minutes,
        days_ahead=cfg.interview_search_days_ahead,
        slots_wanted=cfg.interview_slot_options,
        weekdays_only=cfg.working_hours.interview_weekdays_only,
        prefer_around=None,
    )
    if slots:
        slot_lines = "\n".join(
            f"  • {s.format_for_email(cfg.working_hours.timezone)}" for s in slots
        )
    else:
        slot_lines = (
            "  • (no calendar availability found in the next 14 days — "
            "please suggest a few times that work for you)"
        )

    fmt = classification.interview_format_hint
    format_line = f" — happy to do {fmt}" if fmt else ""

    return TEMPLATE.format(
        first_name=_first_name(msg.sender_name, msg.sender_email),
        slots=slot_lines,
        format_line=format_line,
    )
