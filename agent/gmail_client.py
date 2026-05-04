"""Gmail API wrapper: fetch unread, label, trash, star, draft."""
from __future__ import annotations

import base64
from dataclasses import dataclass, field
from datetime import datetime, timezone
from email.mime.text import MIMEText
from typing import Any

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

LABEL_NAMES = [
    "auto/processed",
    "auto/would-trash",
    "auto/would-star",
    "auto/would-draft",
    "auto/mistake",
    "AutoDraft Interviews",
    "Jobs 2026",
    "Bills/Due",
    "Reply Needed",
    "Urgent",
]


@dataclass
class Message:
    id: str
    thread_id: str
    sender_email: str
    sender_name: str
    subject: str
    snippet: str
    body_text: str
    received_utc: datetime
    label_ids: list[str] = field(default_factory=list)
    has_list_unsubscribe: bool = False
    in_promotions: bool = False
    raw_headers: dict[str, str] = field(default_factory=dict)


def _build_credentials(client_id: str, client_secret: str, refresh_token: str) -> Credentials:
    creds = Credentials(
        token=None,
        refresh_token=refresh_token,
        token_uri="https://oauth2.googleapis.com/token",
        client_id=client_id,
        client_secret=client_secret,
        scopes=[
            "https://www.googleapis.com/auth/gmail.modify",
            "https://www.googleapis.com/auth/gmail.compose",
            "https://www.googleapis.com/auth/calendar.readonly",
        ],
    )
    creds.refresh(Request())
    return creds


def _parse_address(raw: str) -> tuple[str, str]:
    """Parse 'Display Name <email@x.com>' or just 'email@x.com'."""
    raw = (raw or "").strip()
    if "<" in raw and ">" in raw:
        name = raw.split("<")[0].strip().strip('"')
        email = raw.split("<")[1].split(">")[0].strip()
        return name, email.lower()
    return "", raw.lower()


def _decode_body(payload: dict[str, Any]) -> str:
    """Walk MIME tree and concatenate text/plain parts (fallback to text/html)."""
    def walk(part: dict[str, Any]) -> str:
        mime = part.get("mimeType", "")
        body = part.get("body", {})
        data = body.get("data")
        text = ""
        if data and mime.startswith("text/"):
            text = base64.urlsafe_b64decode(data).decode("utf-8", errors="replace")
        for child in part.get("parts", []) or []:
            text += "\n" + walk(child)
        return text

    plain = walk(payload)
    return plain.strip()


class GmailClient:
    def __init__(self, client_id: str, client_secret: str, refresh_token: str):
        creds = _build_credentials(client_id, client_secret, refresh_token)
        self.service = build("gmail", "v1", credentials=creds, cache_discovery=False)
        self._label_cache: dict[str, str] = {}

    def ensure_labels(self) -> None:
        existing = self.service.users().labels().list(userId="me").execute().get("labels", [])
        for lab in existing:
            self._label_cache[lab["name"]] = lab["id"]
        for name in LABEL_NAMES:
            if name not in self._label_cache:
                created = self.service.users().labels().create(
                    userId="me",
                    body={"name": name, "labelListVisibility": "labelShow",
                          "messageListVisibility": "show"},
                ).execute()
                self._label_cache[name] = created["id"]

    def label_id(self, name: str) -> str:
        return self._label_cache[name]

    def fetch_unread_since(self, since_utc: datetime, max_results: int) -> list[Message]:
        """Return unread messages in INBOX received after `since_utc`,
        excluding any already labeled `auto/processed`."""
        epoch_secs = int(since_utc.replace(tzinfo=timezone.utc).timestamp())
        query = f"is:unread in:inbox after:{epoch_secs} -label:auto/processed"
        out: list[Message] = []
        page_token: str | None = None
        while True:
            resp = self.service.users().messages().list(
                userId="me", q=query, maxResults=min(100, max_results - len(out)),
                pageToken=page_token,
            ).execute()
            for stub in resp.get("messages", []):
                msg = self._get_full(stub["id"])
                if msg:
                    out.append(msg)
                if len(out) >= max_results:
                    return out
            page_token = resp.get("nextPageToken")
            if not page_token:
                return out

    def _get_full(self, message_id: str) -> Message | None:
        full = self.service.users().messages().get(
            userId="me", id=message_id, format="full"
        ).execute()
        payload = full.get("payload", {})
        headers = {h["name"].lower(): h["value"] for h in payload.get("headers", [])}
        sender_name, sender_email = _parse_address(headers.get("from", ""))
        subject = headers.get("subject", "")
        internal = int(full.get("internalDate", "0")) / 1000
        received = datetime.fromtimestamp(internal, tz=timezone.utc)
        body = _decode_body(payload)
        label_ids = full.get("labelIds", [])
        return Message(
            id=full["id"],
            thread_id=full["threadId"],
            sender_email=sender_email,
            sender_name=sender_name,
            subject=subject,
            snippet=full.get("snippet", ""),
            body_text=body[:8000],  # cap body size sent to LLM
            received_utc=received,
            label_ids=label_ids,
            has_list_unsubscribe="list-unsubscribe" in headers,
            in_promotions="CATEGORY_PROMOTIONS" in label_ids,
            raw_headers=headers,
        )

    def add_labels(self, message_id: str, label_names: list[str]) -> None:
        ids = [self._label_cache[n] for n in label_names]
        self.service.users().messages().modify(
            userId="me", id=message_id, body={"addLabelIds": ids}
        ).execute()

    def remove_labels(self, message_id: str, label_names: list[str]) -> None:
        ids = [self._label_cache[n] for n in label_names]
        self.service.users().messages().modify(
            userId="me", id=message_id, body={"removeLabelIds": ids}
        ).execute()

    def trash(self, message_id: str) -> None:
        self.service.users().messages().trash(userId="me", id=message_id).execute()

    def star(self, message_id: str) -> None:
        self.service.users().messages().modify(
            userId="me", id=message_id, body={"addLabelIds": ["STARRED"]}
        ).execute()

    def create_draft_reply(
        self, *, thread_id: str, to_address: str, subject: str, body_text: str,
        in_reply_to: str | None,
    ) -> str:
        msg = MIMEText(body_text, _charset="utf-8")
        msg["To"] = to_address
        msg["Subject"] = subject if subject.lower().startswith("re:") else f"Re: {subject}"
        if in_reply_to:
            msg["In-Reply-To"] = in_reply_to
            msg["References"] = in_reply_to
        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
        body = {"message": {"raw": raw, "threadId": thread_id}}
        created = self.service.users().drafts().create(userId="me", body=body).execute()
        return created["id"]

    def list_messages_in_label(self, label_name: str, max_results: int = 50) -> list[str]:
        """Return message IDs currently in a given label (used for mistake-learning)."""
        if label_name not in self._label_cache:
            return []
        resp = self.service.users().messages().list(
            userId="me", labelIds=[self._label_cache[label_name]], maxResults=max_results,
        ).execute()
        return [m["id"] for m in resp.get("messages", [])]
