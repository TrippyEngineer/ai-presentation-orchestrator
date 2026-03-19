"""
n8n_trigger.py — Fires webhook triggers to n8n demo workflows.

All webhook URLs are loaded from .env — never hardcode them.
Demo payloads are illustrative samples; replace with your own content.
"""

import os
import requests
from dotenv import load_dotenv

load_dotenv()

EMAIL_WEBHOOK    = os.getenv("N8N_WEBHOOK_EMAIL")
MEETING_WEBHOOK  = os.getenv("N8N_WEBHOOK_MEETING")
RESEARCH_WEBHOOK = os.getenv("N8N_WEBHOOK_RESEARCH")


def trigger_email_demo():
    """
    Sample: inbound donor email requesting a programme update.
    Replace the payload with a realistic email for your own domain.
    """
    payload = {
        "Subject": "Q1 Programme Update Request",
        "From": "Sarah Johnson <sarah@example.org>",
        "body": (
            "Hi, we are preparing our Q1 donor review and would appreciate an update "
            "on programme outcomes in the target region. Specifically we need coverage rates, "
            "key achievements, and any emerging implementation challenges. "
            "Could you send this by Friday? Best regards, Sarah"
        ),
        "threadId": ""
    }
    response = requests.post(EMAIL_WEBHOOK, json=payload, timeout=30)
    print(f"Email Demo triggered: {response.status_code}")
    return response


def trigger_meeting_demo():
    """
    Sample: meeting transcript → structured outputs pipeline.
    Place your own transcript at demo_data/meeting_transcript.txt
    """
    with open("demo_data/meeting_transcript.txt") as f:
        transcript = f.read()
    payload = {
        "Meeting Name": "Country Team Meeting (Sample)",
        "Meeting Transcript": transcript
    }
    response = requests.post(MEETING_WEBHOOK, json=payload, timeout=30)
    print(f"Meeting Demo triggered: {response.status_code}")
    return response


def trigger_research_demo():
    """
    Sample: agentic research question → evidence brief.
    Replace the question and parameters with your own domain.
    """
    payload = {
        "Input Mode": "Research Question",
        "Research Question": (
            "What is the evidence on zinc supplementation effectiveness "
            "for reducing childhood diarrhea mortality and what are the barriers to scale-up?"
        ),
        "Meeting Transcript": "",
        "Program Area": "Nutrition",
        "Geographic Region": "Sub-Saharan Africa",
        "Audience Type": "Donor Brief",
        "Urgency": "Standard"
    }
    response = requests.post(RESEARCH_WEBHOOK, json=payload, timeout=90)
    print(f"Research Demo triggered: {response.status_code}")
    return response


if __name__ == "__main__":
    print("Testing all 3 webhooks...")
    print("\n1. Email Pipeline:")
    trigger_email_demo()
    print("\n2. Meeting Pipeline:")
    trigger_meeting_demo()
    print("\n3. Research Engine:")
    trigger_research_demo()
    print("\nAll triggers fired. Check n8n, Sheets, and Slack for outputs.")
