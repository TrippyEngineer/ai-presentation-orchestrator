"""
diagnose.py — Pre-presentation system diagnostic.

Run this before every session to verify all dependencies are ready.
"""

import sys
import os
import json
import subprocess
import time
import platform
from pathlib import Path

from logger import get_logger
from dotenv import load_dotenv

load_dotenv()
log = get_logger("diagnose")

PPTX_FILE = os.getenv("PPTX_FILE", "presentation.pptx")

print("\n" + "="*65)
print("  AI PRESENTATION ORCHESTRATOR — PRE-FLIGHT DIAGNOSTIC")
print(f"  Platform : {platform.system()} {platform.release()}")
print(f"  Python   : {sys.version.split()[0]}")
print(f"  Folder   : {Path('.').resolve()}")
print("="*65 + "\n")

passed = 0
failed = 0
warned = 0


def ok(label, detail=""):
    global passed
    passed += 1
    print(f"  ✅  {label}")
    if detail: print(f"      {detail}")
    log.info(f"PASS: {label} {detail}")


def fail(label, fix="", detail=""):
    global failed
    failed += 1
    print(f"  ❌  {label}")
    if detail: print(f"      {detail}")
    if fix:    print(f"      FIX: {fix}")
    log.error(f"FAIL: {label} | fix: {fix}")


def warn(label, detail=""):
    global warned
    warned += 1
    print(f"  ⚠️   {label}")
    if detail: print(f"      {detail}")
    log.warning(f"WARN: {label} | {detail}")


def find_ffmpeg_tool(tool: str):
    try:
        subprocess.run([tool, "-version"], capture_output=True, timeout=5)
        return tool
    except FileNotFoundError:
        pass
    if platform.system() == "Windows":
        try:
            result = subprocess.run(["where", tool], capture_output=True, text=True, timeout=5)
            if result.returncode == 0:
                full_path = result.stdout.strip().splitlines()[0]
                subprocess.run([full_path, "-version"], capture_output=True, timeout=5)
                return full_path
        except Exception:
            pass
        common_dirs = [
            r"C:\ffmpeg\bin",
            r"C:\Program Files\ffmpeg\bin",
            r"C:\Program Files (x86)\ffmpeg\bin",
            os.path.expanduser(r"~\ffmpeg\bin"),
        ]
        for d in common_dirs:
            candidate = os.path.join(d, tool + ".exe")
            if os.path.exists(candidate):
                try:
                    subprocess.run([candidate, "-version"], capture_output=True, timeout=5)
                    return candidate
                except Exception:
                    pass
    return None


# ── 1. PYTHON PACKAGES ────────────────────────────────────────────────
print("[ PYTHON PACKAGES ]")
for imp, pkg, cmd in [
    ("edge_tts",    "edge-tts",       "pip install edge-tts"),
    ("anthropic",   "anthropic",      "pip install anthropic"),
    ("dotenv",      "python-dotenv",  "pip install python-dotenv"),
    ("pptx",        "python-pptx",    "pip install python-pptx"),
    ("pyautogui",   "pyautogui",      "pip install pyautogui"),
    ("requests",    "requests",       "pip install requests"),
    ("pygame",      "pygame",         "pip install pygame"),
    ("pynput",      "pynput",         "pip install pynput  ← required for global keyboard control"),
]:
    try:
        __import__(imp)
        ok(pkg)
    except ImportError:
        fail(pkg, fix=cmd)
print()


# ── 2. FFMPEG ─────────────────────────────────────────────────────────
print("[ FFMPEG ]")
ffplay_path  = find_ffmpeg_tool("ffplay")
ffprobe_path = find_ffmpeg_tool("ffprobe")
if ffplay_path:  ok("ffplay found",  detail=ffplay_path)
else:            fail("ffplay not found",  fix="Install ffmpeg: https://www.gyan.dev/ffmpeg/builds/ and add bin\\ to PATH")
if ffprobe_path: ok("ffprobe found", detail=ffprobe_path)
else:            fail("ffprobe not found", fix="Same package — ffprobe is in the same bin\\ folder")
print()


# ── 3. .env FILE ──────────────────────────────────────────────────────
print("[ .env FILE ]")
if Path(".env").exists():
    ok(".env file found")
else:
    fail(".env file missing", fix="Copy .env.example to .env and fill in your values")

for var, hint in [
    ("ANTHROPIC_API_KEY",    "sk-ant-..."),
    ("N8N_WEBHOOK_EMAIL",    "http://localhost:5678/webhook/..."),
    ("N8N_WEBHOOK_MEETING",  "http://localhost:5678/webhook/..."),
    ("N8N_WEBHOOK_RESEARCH", "http://localhost:5678/webhook/..."),
]:
    val = os.getenv(var, "")
    if val:
        ok(f"{var}", detail=val[:25] + "...")
    else:
        fail(f"{var} not set", fix=f"Add {var}={hint} to .env")
print()


# ── 4. PPTX ───────────────────────────────────────────────────────────
print("[ PRESENTATION FILE ]")
pptx = Path(PPTX_FILE)
if pptx.exists():
    mb = pptx.stat().st_size / 1024 / 1024
    ok(f"PPTX found ({mb:.1f} MB) — {PPTX_FILE}")
    try:
        from pptx import Presentation
        prs = Presentation(str(pptx))
        n   = len(prs.slides)
        ok(f"Slide count: {n} slides")
        log.info(f"PPTX has {n} slides")
    except Exception as e:
        warn(f"Could not read PPTX: {e}")
else:
    fail(f"{PPTX_FILE} not found",
         fix="Set PPTX_FILE in .env or place presentation.pptx in this folder")
print()


# ── 5. DEMO DATA ──────────────────────────────────────────────────────
print("[ DEMO DATA ]")
transcript = Path("demo_data/meeting_transcript.txt")
if transcript.exists():
    words = len(transcript.read_text().split())
    ok(f"meeting_transcript.txt ({words} words)")
else:
    fail("demo_data/meeting_transcript.txt missing",
         fix="mkdir demo_data && create meeting_transcript.txt inside it")
print()


# ── 6. AUDIO CACHE ────────────────────────────────────────────────────
print("[ AUDIO CACHE ]")
manifest_path = Path("cache/manifest.json")
if manifest_path.exists():
    manifest  = json.loads(manifest_path.read_text())
    total     = len(manifest)
    have      = [k for k, v in manifest.items() if v.get("audio_path") and Path(v["audio_path"]).exists()]
    missing   = [k for k, v in manifest.items() if v.get("audio_path") and not Path(v["audio_path"]).exists()]
    total_dur = sum(v.get("duration", 0) for v in manifest.values())
    demos     = [k for k, v in manifest.items() if v.get("is_demo")]
    ok(f"manifest.json — {total} slides, {total_dur/60:.1f} min, demo slides: {demos}")
    if missing:
        fail(f"Missing audio files: slides {missing}", fix="python pre_generate.py")
    else:
        ok(f"All {len(have)} audio files present")
else:
    fail("cache/manifest.json missing", fix="python pre_generate.py")
print()


# ── 7. VOICE CONFIGURATION ────────────────────────────────────────────
print("[ VOICE CONFIGURATION ]")
try:
    from voice_engine import VOICE, RATE
    ok(f"Voice: {VOICE}   Rate: {RATE}")
except Exception as e:
    warn(f"Could not check voice config: {e}")
print()


# ── 8. AUDIO PLAYBACK TEST ────────────────────────────────────────────
print("[ AUDIO PLAYBACK — LIVE TEST ]")
test_mp3 = None
if manifest_path.exists():
    manifest = json.loads(manifest_path.read_text())
    for v in manifest.values():
        ap = v.get("audio_path")
        if ap and Path(ap).exists():
            test_mp3 = ap
            break

if test_mp3:
    print(f"  File: {test_mp3}")
    print("  Playing now — you should hear audio...")
    played = False

    try:
        import pygame
        if not pygame.mixer.get_init():
            pygame.mixer.pre_init(44100, -16, 2, 2048)
            pygame.mixer.init()
        pygame.mixer.music.load(test_mp3)
        pygame.mixer.music.play()
        t = time.time()
        while pygame.mixer.music.get_busy() and time.time() - t < 8:
            time.sleep(0.1)
        pygame.mixer.music.stop()
        played = True
        print("  (played via pygame)")
    except Exception:
        pass

    if not played and ffplay_path:
        try:
            subprocess.run(
                [ffplay_path, "-nodisp", "-autoexit", "-loglevel", "quiet", "-t", "6", test_mp3],
                timeout=15, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
            )
            played = True
            print("  (played via ffplay)")
        except Exception:
            pass

    if not played and platform.system() == "Windows":
        try:
            import ctypes
            winmm  = ctypes.windll.winmm
            alias  = "diagtest"
            p      = str(Path(test_mp3).resolve())
            winmm.mciSendStringW(f'open "{p}" type mpegvideo alias {alias}', None, 0, None)
            winmm.mciSendStringW(f"play {alias} wait", None, 0, None)
            winmm.mciSendStringW(f"close {alias}", None, 0, None)
            played = True
            print("  (played via winmm.dll)")
        except Exception:
            pass

    heard = input("  Did you hear audio? (y/n): ").strip().lower()
    if heard == "y":
        ok("Audio playback confirmed")
    else:
        fail("Audio playback failed", fix="pip install pygame OR fix ffmpeg PATH OR check speakers")
else:
    warn("No cached audio — run: python pre_generate.py")
print()


# ── 9. n8n ────────────────────────────────────────────────────────────
print("[ n8n CONNECTIVITY ]")
try:
    import requests
    r = requests.get("http://localhost:5678/healthz", timeout=5)
    if r.status_code == 200:
        ok("n8n running at localhost:5678")
    else:
        warn(f"n8n health check returned {r.status_code}")
except Exception:
    fail("n8n not running", fix="Open new terminal: npx n8n  (keep it open)")

for key in ["N8N_WEBHOOK_EMAIL", "N8N_WEBHOOK_MEETING", "N8N_WEBHOOK_RESEARCH"]:
    url = os.getenv(key, "")
    if url:
        ok(f"{key} configured")
    else:
        fail(f"{key} not set", fix=f"Add {key}=http://localhost:5678/webhook/... to .env")
print()


# ── 10. KEYBOARD LISTENER ─────────────────────────────────────────────
print("[ KEYBOARD CONTROL ]")
try:
    import pynput
    ok("pynput available — global keyboard listener will work (captures even when PPT has focus)")
except ImportError:
    warn("pynput not installed — keyboard only works when terminal is focused",
         detail="FIX: pip install pynput")
print()


# ── SUMMARY ───────────────────────────────────────────────────────────
print("="*65)
total_checks = passed + failed
print(f"  {passed}/{total_checks} passed  |  {failed} failed  |  {warned} warnings")
log.info(f"Diagnostic complete: {passed} pass, {failed} fail, {warned} warn")
if failed == 0:
    print("\n  ✅ ALL CLEAR.")
    print("  Run: python orchestrator.py")
    print()
    print("  CONTROLS during presentation:")
    print("    SPACEBAR   : pause / resume narration")
    print("    D          : skip demo wait")
    print("    Q          : quit")
else:
    print(f"\n  Fix the {failed} issue(s) above, then re-run: python diagnose.py")
print("="*65 + "\n")
