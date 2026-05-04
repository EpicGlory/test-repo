"""Load YAML config files and required environment variables."""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml

REPO_ROOT = Path(__file__).resolve().parent.parent
CONFIG_DIR = REPO_ROOT / "config"
STATE_DIR = REPO_ROOT / "state"
PAUSE_FILE = REPO_ROOT / "PAUSE"


@dataclass
class WorkingHours:
    timezone: str
    sms_days: list[int]
    sms_hours: list[int]
    morning_email_hour: int
    interview_earliest_hour: int
    interview_latest_hour: int
    interview_weekdays_only: bool


@dataclass
class Whitelist:
    domains: list[str]
    addresses: list[str]
    learned: list[str] = field(default_factory=list)

    def matches(self, sender_email: str) -> bool:
        sender = sender_email.lower().strip()
        if sender in {a.lower() for a in self.addresses + self.learned}:
            return True
        for pattern in self.domains:
            p = pattern.lower().strip()
            if p.startswith("*@") and sender.endswith(p[1:]):
                return True
            if p == sender:
                return True
        return False


@dataclass
class AgentConfig:
    dry_run: bool
    daily_spend_cap_usd: float
    anthropic_model: str
    interview_slot_options: int
    interview_buffer_minutes: int
    interview_slot_minutes: int
    interview_search_days_ahead: int
    first_run_lookback_hours: int
    processed_id_cache_size: int
    sms_max_bytes: int
    max_messages_per_run: int

    working_hours: WorkingHours
    whitelist: Whitelist
    tone_sample: str

    google_client_id: str
    google_client_secret: str
    google_refresh_token: str
    anthropic_api_key: str
    smtp_user: str
    smtp_pass: str
    sms_to: str
    owner_email: str

    manual_mode: str | None = None


def _load_yaml(name: str) -> dict[str, Any]:
    return yaml.safe_load((CONFIG_DIR / name).read_text())


def _require_env(name: str) -> str:
    val = os.environ.get(name)
    if not val:
        raise RuntimeError(f"Missing required environment variable: {name}")
    return val.strip()


def load_config() -> AgentConfig:
    agent = _load_yaml("agent.yml")
    wh_raw = _load_yaml("working_hours.yml")
    wl_raw = _load_yaml("whitelist.yml")

    learned_path = STATE_DIR / "learned_whitelist.json"
    learned: list[str] = []
    if learned_path.exists():
        import json
        learned = json.loads(learned_path.read_text()).get("addresses", [])

    wh = WorkingHours(
        timezone=wh_raw["timezone"],
        sms_days=wh_raw["sms_days"],
        sms_hours=wh_raw["sms_hours"],
        morning_email_hour=wh_raw["morning_email_hour"],
        interview_earliest_hour=wh_raw["interview_window"]["earliest_hour"],
        interview_latest_hour=wh_raw["interview_window"]["latest_hour"],
        interview_weekdays_only=wh_raw["interview_window"]["weekdays_only"],
    )

    wl = Whitelist(
        domains=wl_raw.get("domains") or [],
        addresses=wl_raw.get("addresses") or [],
        learned=learned,
    )

    tone = (CONFIG_DIR / "tone_sample.txt").read_text()

    return AgentConfig(
        dry_run=bool(agent["dry_run"]),
        daily_spend_cap_usd=float(agent["daily_spend_cap_usd"]),
        anthropic_model=agent["anthropic_model"],
        interview_slot_options=int(agent["interview_slot_options"]),
        interview_buffer_minutes=int(agent["interview_buffer_minutes"]),
        interview_slot_minutes=int(agent["interview_slot_minutes"]),
        interview_search_days_ahead=int(agent["interview_search_days_ahead"]),
        first_run_lookback_hours=int(agent["first_run_lookback_hours"]),
        processed_id_cache_size=int(agent["processed_id_cache_size"]),
        sms_max_bytes=int(agent["sms_max_bytes"]),
        max_messages_per_run=int(agent["max_messages_per_run"]),
        working_hours=wh,
        whitelist=wl,
        tone_sample=tone,
        google_client_id=_require_env("GOOGLE_CLIENT_ID"),
        google_client_secret=_require_env("GOOGLE_CLIENT_SECRET"),
        google_refresh_token=_require_env("GOOGLE_REFRESH_TOKEN"),
        anthropic_api_key=_require_env("ANTHROPIC_API_KEY"),
        smtp_user=_require_env("SMTP_USER"),
        smtp_pass=_require_env("SMTP_PASS"),
        sms_to=_require_env("SMS_TO"),
        owner_email=_require_env("OWNER_EMAIL"),
        manual_mode=os.environ.get("MANUAL_MODE") or None,
    )
