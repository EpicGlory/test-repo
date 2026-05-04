"""Persistent state shared across runs (committed back to the repo)."""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path

from .config import STATE_DIR

LAST_RUN_FILE = STATE_DIR / "last_run.json"
LEARNED_WL_FILE = STATE_DIR / "learned_whitelist.json"


@dataclass
class State:
    last_run_utc: datetime | None = None
    processed_message_ids: list[str] = field(default_factory=list)
    last_morning_email_date: str | None = None

    @classmethod
    def load(cls) -> "State":
        if not LAST_RUN_FILE.exists():
            return cls()
        raw = json.loads(LAST_RUN_FILE.read_text())
        last_run = raw.get("last_run_utc")
        return cls(
            last_run_utc=datetime.fromisoformat(last_run) if last_run else None,
            processed_message_ids=raw.get("processed_message_ids", []),
            last_morning_email_date=raw.get("last_morning_email_date"),
        )

    def save(self, *, cache_size: int) -> None:
        ids = self.processed_message_ids[-cache_size:]
        payload = {
            "last_run_utc": self.last_run_utc.isoformat() if self.last_run_utc else None,
            "processed_message_ids": ids,
            "last_morning_email_date": self.last_morning_email_date,
        }
        LAST_RUN_FILE.write_text(json.dumps(payload, indent=2) + "\n")


def load_learned_whitelist() -> list[str]:
    if not LEARNED_WL_FILE.exists():
        return []
    return json.loads(LEARNED_WL_FILE.read_text()).get("addresses", [])


def save_learned_whitelist(addresses: list[str]) -> None:
    payload = {
        "addresses": sorted(set(addresses)),
        "updated_utc": datetime.now(timezone.utc).isoformat(),
    }
    LEARNED_WL_FILE.write_text(json.dumps(payload, indent=2) + "\n")
