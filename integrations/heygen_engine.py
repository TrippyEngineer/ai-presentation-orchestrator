"""
heygen_engine.py — HeyGen AI avatar voice synthesis.

Docs: https://docs.heygen.com/reference/tts-audio-v2
      https://docs.heygen.com/reference/create-avatar-video-v2

Two modes:
  1. TTS only  — generates an MP3 for the voice, used by the orchestrator
                 (fast, ~2–5s per slide, works for live presentation)
  2. Avatar video — generates a full talking-head video (slow, ~30–60s per slide,
                    use for pre-recorded content, not live orchestration)

Set VOICE_ENGINE=heygen in .env to activate this engine.
"""

import os
import time
import requests
from pathlib import Path
from dotenv import load_dotenv
from core.logger import get_logger

load_dotenv()
log = get_logger("heygen_engine")

API_KEY   = os.getenv("HEYGEN_API_KEY", "")
VOICE_ID  = os.getenv("HEYGEN_VOICE_ID", "")
AVATAR_ID = os.getenv("HEYGEN_AVATAR_ID", "")

BASE_URL  = "https://api.heygen.com"
CACHE_DIR = Path("cache")
CACHE_DIR.mkdir(exist_ok=True)

HEADERS = {
    "X-Api-Key": API_KEY,
    "Content-Type": "application/json",
    "Accept": "application/json",
}


# ── 1. TEXT-TO-SPEECH (fast — use for live presentation) ─────────────

def synthesise_tts(text: str, output_path: str, voice_id: str = VOICE_ID) -> bool:
    """
    Generate MP3 audio from text using HeyGen TTS API.
    Fast (~2–5s). Best for live presentation narration.

    Returns True on success, False on failure.
    """
    if not API_KEY:
        log.error("HEYGEN_API_KEY not set in .env")
        return False
    if not voice_id:
        log.error("HEYGEN_VOICE_ID not set in .env")
        return False
    if not text.strip():
        return False

    log.info(f"HeyGen TTS: {len(text.split())} words → {output_path}")

    try:
        resp = requests.post(
            f"{BASE_URL}/v2/audio/tts",
            headers=HEADERS,
            json={"text": text, "voice_id": voice_id},
            timeout=30,
        )

        if resp.status_code != 200:
            log.error(f"HeyGen TTS HTTP {resp.status_code}: {resp.text[:200]}")
            return False

        data = resp.json()

        # Response contains a signed URL to download the audio
        audio_url = (
            data.get("data", {}).get("url") or
            data.get("audio_url") or
            data.get("url")
        )

        if not audio_url:
            log.error(f"HeyGen TTS: no audio URL in response: {data}")
            return False

        # Download the audio file
        audio_resp = requests.get(audio_url, timeout=30)
        if audio_resp.status_code != 200:
            log.error(f"HeyGen audio download failed: HTTP {audio_resp.status_code}")
            return False

        Path(output_path).write_bytes(audio_resp.content)
        size = Path(output_path).stat().st_size
        log.info(f"HeyGen TTS saved: {output_path} ({size // 1024}KB)")
        return True

    except requests.Timeout:
        log.error("HeyGen TTS request timed out")
        return False
    except Exception as e:
        log.error(f"HeyGen TTS error: {e}", exc_info=True)
        return False


# ── 2. LIST AVAILABLE VOICES ──────────────────────────────────────────

def list_voices() -> list[dict]:
    """Fetch all available HeyGen TTS voices. Use to find your HEYGEN_VOICE_ID."""
    if not API_KEY:
        print("❌ Set HEYGEN_API_KEY in .env first")
        return []
    try:
        resp = requests.get(f"{BASE_URL}/v2/voices", headers=HEADERS, timeout=15)
        data = resp.json()
        voices = data.get("data", {}).get("voices", [])
        log.info(f"HeyGen voices fetched: {len(voices)}")
        return voices
    except Exception as e:
        log.error(f"list_voices error: {e}")
        return []


# ── 3. AVATAR VIDEO (slow — use for pre-recorded content) ─────────────

def generate_avatar_video(
    text:       str,
    output_path: str,
    avatar_id:  str = AVATAR_ID,
    voice_id:   str = VOICE_ID,
    poll_interval: int = 5,
    timeout:    int = 300,
) -> bool:
    """
    Generate a talking-head avatar video using HeyGen.
    Slow (~30–120s). Use for pre-recorded segment videos, NOT live orchestration.

    Returns True and saves MP4 to output_path on success.
    """
    if not all([API_KEY, avatar_id, voice_id]):
        log.error("HEYGEN_API_KEY, HEYGEN_AVATAR_ID, HEYGEN_VOICE_ID must all be set")
        return False

    log.info(f"HeyGen avatar video: {len(text.split())} words → {output_path}")

    # Step 1: Submit generation job
    try:
        resp = requests.post(
            f"{BASE_URL}/v2/video/generate",
            headers=HEADERS,
            json={
                "video_inputs": [{
                    "character": {
                        "type":      "avatar",
                        "avatar_id": avatar_id,
                        "scale":     1.0,
                    },
                    "voice": {
                        "type":     "text",
                        "input_text": text,
                        "voice_id": voice_id,
                    },
                }],
                "dimension": {"width": 1280, "height": 720},
                "aspect_ratio": "16:9",
            },
            timeout=30,
        )

        if resp.status_code not in (200, 201):
            log.error(f"HeyGen video submit HTTP {resp.status_code}: {resp.text[:300]}")
            return False

        video_id = resp.json().get("data", {}).get("video_id")
        if not video_id:
            log.error(f"No video_id in response: {resp.json()}")
            return False

        log.info(f"HeyGen job submitted: video_id={video_id}")
        print(f"  ⏳ HeyGen rendering video_id={video_id}...")

    except Exception as e:
        log.error(f"HeyGen submit error: {e}", exc_info=True)
        return False

    # Step 2: Poll for completion
    deadline = time.time() + timeout
    while time.time() < deadline:
        time.sleep(poll_interval)
        try:
            status_resp = requests.get(
                f"{BASE_URL}/v1/video_status.get?video_id={video_id}",
                headers=HEADERS,
                timeout=15,
            )
            data   = status_resp.json().get("data", {})
            status = data.get("status", "")
            print(f"  ⏳ Status: {status}", end="\r", flush=True)

            if status == "completed":
                video_url = data.get("video_url")
                if not video_url:
                    log.error("completed but no video_url")
                    return False

                dl = requests.get(video_url, timeout=60)
                Path(output_path).write_bytes(dl.content)
                size = Path(output_path).stat().st_size
                log.info(f"HeyGen video saved: {output_path} ({size // (1024*1024):.1f}MB)")
                print(f"  ✅ Saved: {output_path} ({size//(1024*1024):.1f}MB)   ")
                return True

            elif status in ("failed", "error"):
                log.error(f"HeyGen job failed: {data}")
                return False

        except Exception as e:
            log.warning(f"Poll error: {e}")

    log.error(f"HeyGen video timeout after {timeout}s for video_id={video_id}")
    return False


# ── 4. BATCH TTS (for pre_generate.py) ───────────────────────────────

def batch_synthesise_heygen(scripts: dict) -> dict:
    """
    Synthesise all slide scripts using HeyGen TTS.
    Drop-in replacement for voice_engine.batch_synthesise().

    scripts: {slide_num: text}
    Returns: {slide_num: mp3_path}
    """
    results = {}
    for slide_num, text in sorted(scripts.items()):
        if not text.strip():
            print(f"  ⏭ Slide {slide_num}: empty — skipping")
            continue
        mp3_path = str(CACHE_DIR / f"slide_{slide_num:02d}.mp3")
        if Path(mp3_path).exists():
            print(f"  ✅ Slide {slide_num}: already cached")
            results[slide_num] = mp3_path
            continue
        print(f"  🎙 Slide {slide_num}: HeyGen TTS...", end=" ", flush=True)
        ok = synthesise_tts(text, mp3_path)
        if ok:
            size = Path(mp3_path).stat().st_size
            print(f"✅ ({size // 1024}KB)")
            results[slide_num] = mp3_path
        else:
            print("❌ FAILED")
            log.error(f"Slide {slide_num}: HeyGen synthesis FAILED")
    return results


if __name__ == "__main__":
    print("\n── HeyGen Engine Test ────────────────────────────────")
    print("\n1. Listing voices (first 10)...")
    voices = list_voices()
    for v in voices[:10]:
        print(f"   {v.get('voice_id','?'):<30} {v.get('name','?')}")

    print("\n2. TTS test...")
    test_text = "Hello. This is a test of HeyGen text-to-speech synthesis."
    test_path = str(CACHE_DIR / "_heygen_tts_test.mp3")
    ok = synthesise_tts(test_text, test_path)
    print(f"   Result: {'✅ Saved to ' + test_path if ok else '❌ Failed'}")
