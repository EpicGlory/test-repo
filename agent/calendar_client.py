"""Google Calendar wrapper: find free interview slots."""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build


@dataclass
class TimeSlot:
    start_local: datetime
    end_local: datetime

    def format_for_email(self, tz_name: str) -> str:
        """e.g., 'Tuesday, May 6 at 10:00 AM MT'"""
        tz_abbrev = "MT"
        return self.start_local.strftime("%A, %B %-d at %-I:%M %p ") + tz_abbrev


class CalendarClient:
    def __init__(self, client_id: str, client_secret: str, refresh_token: str):
        creds = Credentials(
            token=None,
            refresh_token=refresh_token,
            token_uri="https://oauth2.googleapis.com/token",
            client_id=client_id,
            client_secret=client_secret,
            scopes=["https://www.googleapis.com/auth/calendar.readonly"],
        )
        creds.refresh(Request())
        self.service = build("calendar", "v3", credentials=creds, cache_discovery=False)

    def find_free_slots(
        self,
        *,
        tz_name: str,
        earliest_hour: int,
        latest_hour: int,
        slot_minutes: int,
        buffer_minutes: int,
        days_ahead: int,
        slots_wanted: int,
        weekdays_only: bool,
        prefer_around: datetime | None = None,
    ) -> list[TimeSlot]:
        tz = ZoneInfo(tz_name)
        now_local = datetime.now(tz)
        # Start from tomorrow if it's already past noon local
        start_day = now_local.date() + timedelta(days=1 if now_local.hour >= 12 else 0)
        end_day = start_day + timedelta(days=days_ahead)

        time_min = datetime.combine(start_day, datetime.min.time(), tzinfo=tz)
        time_max = datetime.combine(end_day, datetime.min.time(), tzinfo=tz)

        busy = self._fetch_busy(time_min.astimezone(timezone.utc),
                                time_max.astimezone(timezone.utc))
        # Inflate busy ranges by buffer.
        buffer = timedelta(minutes=buffer_minutes)
        busy = [(s - buffer, e + buffer) for s, e in busy]

        candidates: list[TimeSlot] = []
        day = start_day
        while day < end_day and len(candidates) < slots_wanted * 4:
            if weekdays_only and day.weekday() >= 5:
                day += timedelta(days=1)
                continue
            for hour in range(earliest_hour, latest_hour):
                for minute in (0, 30):
                    start_l = datetime.combine(day, datetime.min.time(),
                                               tzinfo=tz).replace(hour=hour, minute=minute)
                    end_l = start_l + timedelta(minutes=slot_minutes)
                    if end_l.hour > latest_hour or (end_l.hour == latest_hour and end_l.minute > 0):
                        continue
                    if start_l <= now_local + timedelta(hours=12):
                        continue  # don't propose anything within the next 12 hours
                    s_utc = start_l.astimezone(timezone.utc)
                    e_utc = end_l.astimezone(timezone.utc)
                    if not _overlaps_any(s_utc, e_utc, busy):
                        candidates.append(TimeSlot(start_l, end_l))
            day += timedelta(days=1)

        if prefer_around:
            target = prefer_around.astimezone(tz)
            candidates.sort(key=lambda s: abs(s.start_local - target))
        return candidates[:slots_wanted]

    def _fetch_busy(self, time_min_utc: datetime, time_max_utc: datetime) -> list[tuple[datetime, datetime]]:
        body = {
            "timeMin": time_min_utc.isoformat().replace("+00:00", "Z"),
            "timeMax": time_max_utc.isoformat().replace("+00:00", "Z"),
            "items": [{"id": "primary"}],
        }
        resp = self.service.freebusy().query(body=body).execute()
        out: list[tuple[datetime, datetime]] = []
        for slot in resp["calendars"]["primary"].get("busy", []):
            s = datetime.fromisoformat(slot["start"].replace("Z", "+00:00"))
            e = datetime.fromisoformat(slot["end"].replace("Z", "+00:00"))
            out.append((s, e))
        return out


def _overlaps_any(start: datetime, end: datetime, busy: list[tuple[datetime, datetime]]) -> bool:
    for bs, be in busy:
        if start < be and end > bs:
            return True
    return False
