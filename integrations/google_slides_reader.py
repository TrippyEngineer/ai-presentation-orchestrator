"""
google_slides_reader.py — Extract slide content from Google Slides via API.

This is an alternative to slide_reader.py (which reads from a local PPTX).
Use this when your presentation lives in Google Slides.

SETUP (one-time):
  1. Go to https://console.cloud.google.com/
  2. Enable the Google Slides API
  3. Create a Service Account → download JSON key
  4. Share your presentation with the service account email
  5. Set GOOGLE_SLIDES_CREDENTIALS_FILE and GOOGLE_SLIDES_PRESENTATION_ID in .env

Docs: https://developers.google.com/slides/api/reference/rest
"""

import os
import json
from pathlib import Path
from dotenv import load_dotenv
from core.logger import get_logger

load_dotenv()
log = get_logger("gslides_reader")

PRESENTATION_ID   = os.getenv("GOOGLE_SLIDES_PRESENTATION_ID", "")
CREDENTIALS_FILE  = os.getenv("GOOGLE_SLIDES_CREDENTIALS_FILE", "credentials.json")
# Slides to skip (visual/blank slides, comma-separated in .env)
_skip_raw         = os.getenv("SKIP_SLIDES", "")
SKIP_SLIDES       = {int(x) for x in _skip_raw.split(",") if x.strip().isdigit()}


def _get_service():
    """Build and return a Google Slides API service client."""
    try:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
    except ImportError:
        raise ImportError(
            "Google API client not installed.\n"
            "Run: pip install google-api-python-client google-auth"
        )

    if not Path(CREDENTIALS_FILE).exists():
        raise FileNotFoundError(
            f"Service account credentials not found: {CREDENTIALS_FILE}\n"
            "Set GOOGLE_SLIDES_CREDENTIALS_FILE in .env"
        )

    creds = service_account.Credentials.from_service_account_file(
        CREDENTIALS_FILE,
        scopes=["https://www.googleapis.com/auth/presentations.readonly"]
    )
    return build("slides", "v1", credentials=creds, cache_discovery=False)


def extract_slide_content(presentation_id: str = PRESENTATION_ID) -> list[dict]:
    """
    Fetch all slides from Google Slides and extract text content.
    Returns list of dicts: [{slide_number, slide_id, content, has_content}]

    Drop-in replacement for slide_reader.extract_slide_content().
    """
    if not presentation_id:
        raise ValueError(
            "GOOGLE_SLIDES_PRESENTATION_ID not set in .env\n"
            "Find it in your Slides URL: /presentation/d/THIS_IS_THE_ID/edit"
        )

    service = _get_service()
    log.info(f"Fetching presentation: {presentation_id}")

    try:
        prs    = service.presentations().get(presentationId=presentation_id).execute()
        slides = prs.get("slides", [])
        log.info(f"Fetched {len(slides)} slides from Google Slides")
    except Exception as e:
        log.error(f"Google Slides API error: {e}", exc_info=True)
        raise

    result = []
    for i, slide in enumerate(slides):
        slide_num = i + 1
        texts     = []

        for element in slide.get("pageElements", []):
            shape = element.get("shape", {})
            text_content = shape.get("text", {})
            for tr in text_content.get("textElements", []):
                run = tr.get("textRun", {})
                t   = run.get("content", "").strip()
                if t:
                    texts.append(t)

        content     = "\n".join(texts)
        has_content = bool(content.strip()) and slide_num not in SKIP_SLIDES

        result.append({
            "slide_number": slide_num,
            "slide_id":     slide.get("objectId", ""),
            "content":      content,
            "has_content":  has_content,
        })

    return result


def get_presentation_title(presentation_id: str = PRESENTATION_ID) -> str:
    """Returns the title of the Google Slides presentation."""
    service = _get_service()
    prs     = service.presentations().get(presentationId=presentation_id).execute()
    return prs.get("title", "Untitled Presentation")


if __name__ == "__main__":
    print(f"\nFetching slides from: {PRESENTATION_ID}\n")
    slides = extract_slide_content()
    title  = get_presentation_title()
    print(f"Presentation: {title}")
    print(f"Total slides: {len(slides)}\n")
    for s in slides:
        status = "✅" if s["has_content"] else "⬛"
        words  = len(s["content"].split())
        print(f"  {status} Slide {s['slide_number']:2d}: {words:4d} words  [id: {s['slide_id'][:12]}...]")
        if s["content"]:
            preview = s["content"].replace("\n", " ")[:80]
            print(f"     {preview}...")
