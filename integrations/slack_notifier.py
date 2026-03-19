"""
slack_notifier.py — Sends status updates to Slack during the presentation.

Reads SLACK_WEBHOOK_URL from .env.
Used by orchestrator.py to notify:
  - Demo triggered (email / meeting / research)
  - Demo completed
  - Presentation started / ended
  - Q&A started

Set SLACK_WEBHOOK_URL in .env to activate. If not set, all calls silently no-op.
"""

import os
import time
import requests
from dotenv import load_dotenv
from logger import get_logger

load_dotenv()
log          = get_logger("slack")
WEBHOOK_URL  = os.getenv("SLACK_WEBHOOK_URL", "")
PRESENTER    = os.getenv("PRESENTER_NAME",   "Presenter")
ORGANIZATION = os.getenv("ORGANIZATION",     "Workshop")


def _send(payload: dict) -> bool:
    """POST payload to Slack webhook. Silent no-op if URL not configured."""
    if not WEBHOOK_URL:
        return False
    try:
        resp = requests.post(WEBHOOK_URL, json=payload, timeout=5)
        ok   = resp.status_code == 200
        if not ok:
            log.warning(f"Slack HTTP {resp.status_code}: {resp.text[:100]}")
        return ok
    except Exception as e:
        log.warning(f"Slack notify failed: {e}")
        return False


def notify_presentation_started(total_slides: int, total_min: float):
    return _send({"text": (
        f"▶️ *{ORGANIZATION} — Presentation started*\n"
        f"Presenter: {PRESENTER} | {total_slides} slides | ~{total_min:.0f} min"
    )})


def notify_demo_triggered(demo_type: str):
    emoji = {"email": "📧", "meeting": "📋", "research": "🔬"}.get(demo_type, "🚀")
    return _send({"text": f"{emoji} *LIVE DEMO triggered:* `{demo_type.upper()}`"})


def notify_demo_complete(demo_type: str):
    return _send({"text": f"✅ *Demo complete:* `{demo_type.upper()}`"})


def notify_qa_started():
    return _send({"text": f"🎤 *Voice Q&A started* — {PRESENTER} is taking questions live"})


def notify_presentation_ended(slide_count: int):
    ts = time.strftime("%H:%M")
    return _send({"text": (
        f"🏁 *Presentation ended* at {ts}\n"
        f"{slide_count} slides delivered | Logs saved to logs/"
    )})


def notify_error(context: str, detail: str):
    return _send({"text": f"⚠️ *Error in `{context}`:*\n```{detail[:300]}```"})


if __name__ == "__main__":
    print("Sending test notification to Slack...")
    ok = notify_presentation_started(17, 45)
    print(f"Result: {'✅ sent' if ok else '❌ failed (check SLACK_WEBHOOK_URL in .env)'}")
