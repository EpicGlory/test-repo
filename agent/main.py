"""Entry point for each scheduled run.

Decides what to do based on the current Mountain Time clock:
  - 7am MT every day: full processing + morning summary email
  - M-F 9am/12pm/3pm/6pm MT: full processing + SMS digest if applicable
  - Every other tick: security-only sweep (cheap)

The kill switch (a `PAUSE` file in the repo root) short-circuits everything.
"""
from __future__ import annotations

import logging
import sys
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo

from .calendar_client import CalendarClient
from .classifier import Classification, classify
from .config import AgentConfig, PAUSE_FILE, load_config
from .drafter import draft_interview_reply
from .gmail_client import GmailClient, Message
from .notifier import send_morning_summary, send_sms_digest, send_sms_test
from .state import State, load_learned_whitelist, save_learned_whitelist

log = logging.getLogger("agent")
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")


def _now_local(tz_name: str) -> datetime:
    return datetime.now(ZoneInfo(tz_name))


def _is_morning_run(now_local: datetime, cfg: AgentConfig) -> bool:
    return now_local.hour == cfg.working_hours.morning_email_hour


def _is_sms_run(now_local: datetime, cfg: AgentConfig) -> bool:
    weekday = now_local.isoweekday()  # 1=Mon
    return (
        weekday in cfg.working_hours.sms_days
        and now_local.hour in cfg.working_hours.sms_hours
        and now_local.minute < 30  # only the top-of-hour tick
    )


def _is_full_processing_window(now_local: datetime, cfg: AgentConfig) -> bool:
    return _is_morning_run(now_local) or _is_sms_run(now_local)


def main() -> int:
    if PAUSE_FILE.exists():
        log.info("PAUSE file present — exiting immediately.")
        return 0

    cfg = load_config()
    state = State.load()

    now_local = _now_local(cfg.working_hours.timezone)
    now_utc = datetime.now(timezone.utc)
    log.info("Run start. Local time: %s. Mode: %s. Dry run: %s",
             now_local.isoformat(), cfg.manual_mode or "auto", cfg.dry_run)

    # Manual override modes via workflow_dispatch
    if cfg.manual_mode == "sms_test":
        send_sms_test(cfg)
        log.info("SMS test sent.")
        return 0

    full = (
        cfg.manual_mode in {"full", "morning_summary"}
        or _is_full_processing_window(now_local, cfg)
    )
    morning = cfg.manual_mode == "morning_summary" or _is_morning_run(now_local, cfg)

    gmail = GmailClient(cfg.google_client_id, cfg.google_client_secret, cfg.google_refresh_token)
    gmail.ensure_labels()

    # Decide lookback window
    if state.last_run_utc is None:
        since_utc = now_utc - timedelta(hours=cfg.first_run_lookback_hours)
        log.info("First run — looking back %dh.", cfg.first_run_lookback_hours)
    else:
        since_utc = state.last_run_utc

    messages = gmail.fetch_unread_since(since_utc, max_results=cfg.max_messages_per_run)
    new_msgs = [m for m in messages if m.id not in set(state.processed_message_ids)]
    log.info("Fetched %d unread, %d new since last run.", len(messages), len(new_msgs))

    classifications: list[Classification] = []
    if new_msgs:
        classifications = classify(new_msgs, cfg)
        log.info("Classified %d msgs (heuristic).", len(classifications))

    # Apply actions
    actions = _apply_actions(gmail, cfg, new_msgs, classifications)

    # Generate interview drafts (only on full-processing ticks)
    if full and not cfg.dry_run:
        for m, c in zip(new_msgs, classifications):
            if c.primary == "interview_request":
                _generate_draft(gmail, cfg, m, c, state)

    # Mark processed
    for m in new_msgs:
        try:
            gmail.add_labels(m.id, ["auto/processed"])
        except Exception as e:
            log.warning("Failed to add auto/processed to %s: %s", m.id, e)
        state.processed_message_ids.append(m.id)

    # Notifications
    if full:
        _send_notifications(gmail, cfg, state, new_msgs, classifications, morning, now_local)

    # Self-correcting whitelist (Sunday morning only)
    if morning and now_local.isoweekday() == 7:
        _learn_from_mistakes(gmail)

    state.last_run_utc = now_utc
    state.save(cache_size=cfg.processed_id_cache_size)
    log.info("Run complete. Actions: %s", actions)
    return 0


def _apply_actions(
    gmail: GmailClient,
    cfg: AgentConfig,
    msgs: list[Message],
    classifications: list[Classification],
) -> dict[str, int]:
    counts = {"trashed": 0, "starred": 0, "labeled_jobs": 0, "labeled_bills": 0,
              "labeled_reply_needed": 0, "labeled_urgent": 0, "skipped_whitelist": 0}
    for m, c in zip(msgs, classifications):
        whitelisted = cfg.whitelist.matches(m.sender_email)
        # Trash advertisements (with strict guard rails)
        if c.primary == "advertisement" and m.has_list_unsubscribe and m.in_promotions:
            if whitelisted:
                counts["skipped_whitelist"] += 1
                continue
            if cfg.dry_run:
                gmail.add_labels(m.id, ["auto/would-trash"])
            else:
                gmail.trash(m.id)
            counts["trashed"] += 1
            continue

        # Star + (later) text on job-related
        if c.primary in {"job_application_response", "recruiter_cold_outreach", "interview_request"}:
            if cfg.dry_run:
                gmail.add_labels(m.id, ["auto/would-star", "Jobs 2026"])
            else:
                gmail.star(m.id)
                gmail.add_labels(m.id, ["Jobs 2026"])
            counts["starred"] += 1
            counts["labeled_jobs"] += 1

        # Category labels
        if c.primary == "bill_due":
            gmail.add_labels(m.id, ["Bills/Due"])
            counts["labeled_bills"] += 1
        if c.primary == "personal_reply_needed":
            gmail.add_labels(m.id, ["Reply Needed"])
            counts["labeled_reply_needed"] += 1
        if c.primary == "urgent":
            gmail.add_labels(m.id, ["Urgent"])
            counts["labeled_urgent"] += 1
    return counts


def _generate_draft(
    gmail: GmailClient, cfg: AgentConfig, msg: Message, c: Classification, _state: State,
) -> None:
    try:
        calendar = CalendarClient(cfg.google_client_id, cfg.google_client_secret, cfg.google_refresh_token)
        body = draft_interview_reply(msg=msg, classification=c, cfg=cfg, calendar=calendar)
        in_reply_to = msg.raw_headers.get("message-id")
        gmail.create_draft_reply(
            thread_id=msg.thread_id,
            to_address=msg.sender_email,
            subject=msg.subject,
            body_text=body,
            in_reply_to=in_reply_to,
        )
        gmail.add_labels(msg.id, ["AutoDraft Interviews"])
        log.info("Draft created for thread %s", msg.thread_id)
    except Exception as e:
        log.exception("Draft generation failed for %s: %s", msg.id, e)


def _send_notifications(
    gmail: GmailClient,
    cfg: AgentConfig,
    state: State,
    msgs: list[Message],
    classifications: list[Classification],
    is_morning: bool,
    now_local: datetime,
) -> None:
    paired = list(zip(msgs, classifications))

    # Security alerts -> immediate SMS at any tick
    sec = [(m, c) for m, c in paired if c.is_security]
    if sec:
        try:
            send_sms_digest(cfg, sec, urgent=True)
        except Exception as e:
            log.exception("Security SMS send failed: %s", e)

    if is_morning and state.last_morning_email_date != now_local.date().isoformat():
        # Build summary buckets from this run + (optionally) prior runs in last 24h.
        # For simplicity v1: only this run's messages contribute.
        bills = [(m, c) for m, c in paired if c.primary == "bill_due"]
        reply_needed = [(m, c) for m, c in paired if c.primary == "personal_reply_needed"]
        urgent = [(m, c) for m, c in paired if c.primary == "urgent"]
        invites = [(m, c) for m, c in paired if c.primary == "calendar_invite"]
        counts: dict[str, int] = {}
        for _, c in paired:
            counts[c.primary] = counts.get(c.primary, 0) + 1

        # Mistake learning lookahead (just count for the email)
        mistake_ids = gmail.list_messages_in_label("auto/mistake", max_results=20)
        mistakes: list[Message] = []
        for mid in mistake_ids[:20]:
            try:
                mistakes.append(gmail._get_full(mid))  # type: ignore[arg-type]
            except Exception:
                continue

        try:
            send_morning_summary(
                cfg,
                yesterday_counts=counts,
                bills=bills,
                reply_needed=reply_needed,
                urgent=urgent,
                invites=invites,
                mistakes=[m for m in mistakes if m],
            )
            state.last_morning_email_date = now_local.date().isoformat()
        except Exception as e:
            log.exception("Morning summary email failed: %s", e)

    if _is_sms_run(now_local, cfg):
        important = [(m, c) for m, c in paired if c.notify_sms and not c.is_security]
        if important:
            try:
                send_sms_digest(cfg, important)
            except Exception as e:
                log.exception("SMS digest send failed: %s", e)


def _learn_from_mistakes(gmail: GmailClient) -> None:
    """If the user labeled emails `auto/mistake`, add their senders to the
    learned whitelist so they don't get auto-trashed in the future."""
    try:
        ids = gmail.list_messages_in_label("auto/mistake", max_results=50)
        addresses = load_learned_whitelist()
        for mid in ids:
            m = gmail._get_full(mid)  # type: ignore[arg-type]
            if m and m.sender_email and m.sender_email not in addresses:
                addresses.append(m.sender_email)
        save_learned_whitelist(addresses)
    except Exception as e:
        log.exception("Learned-whitelist update failed: %s", e)


if __name__ == "__main__":
    sys.exit(main())
