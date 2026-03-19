"""
pre_generate.py — Pre-generates all speech audio using Edge TTS (FREE).

Run:  python pre_generate.py
      python pre_generate.py --force   # regenerate all, ignore cache

Reads slide count and PPTX path from .env (see .env.example).
"""

import json
import os
import sys
import time
from pathlib import Path
from dotenv import load_dotenv

from logger import get_logger, log_slide_event
from slide_reader import extract_slide_content
from script_agent import generate_speech_for_slide, DEMO_SLIDES
from voice_engine import batch_synthesise, get_audio_duration, VOICE, RATE

load_dotenv()
log = get_logger("pre_generate")

CACHE_DIR  = Path("cache")
MANIFEST   = CACHE_DIR / "manifest.json"
SCRIPTS    = CACHE_DIR / "scripts.json"
PPTX_FILE  = os.getenv("PPTX_FILE", "presentation.pptx")
CACHE_DIR.mkdir(exist_ok=True)


def run_pre_generation(force_regenerate: bool = False):
    print("\n" + "="*65)
    print("  AI PRESENTATION — PRE-GENERATION")
    print("="*65)
    print(f"  Voice : {VOICE}")
    print(f"  Rate  : {RATE}")
    print()
    log.info(f"Pre-generation started. Voice={VOICE} Rate={RATE} force={force_regenerate}")

    # STEP 1: Read slides
    print("STEP 1: Reading slides from PPTX...")
    try:
        slides = extract_slide_content(PPTX_FILE)
    except FileNotFoundError as e:
        log.error(f"PPTX not found: {e}")
        print(f"  ❌ {e}")
        sys.exit(1)

    total = len(slides)
    print(f"  Found {total} slides\n")
    log.info(f"PPTX loaded: {total} slides")

    # STEP 2: Generate scripts via Claude
    print("STEP 2: Generating scripts via Claude...")
    existing = json.loads(SCRIPTS.read_text()) if SCRIPTS.exists() and not force_regenerate else {}
    scripts: dict[str, str] = {}

    for slide in slides:
        num = slide["slide_number"]
        key = str(num)
        if key in existing and not force_regenerate:
            scripts[key] = existing[key]
            print(f"  Slide {num:2d}: cached ({len(scripts[key].split())} words)")
            log.debug(f"Slide {num}: using cached script")
        else:
            print(f"  Slide {num:2d}: generating...", end=" ", flush=True)
            try:
                scripts[key] = generate_speech_for_slide(num, slide["content"])
                word_count = len(scripts[key].split())
                demo_tag = " [DEMO]" if num in DEMO_SLIDES else ""
                print(f"done ({word_count} words){demo_tag}")
                log_slide_event(num, "SCRIPT_GENERATED", f"{word_count} words")
                SCRIPTS.write_text(json.dumps(scripts, indent=2))
                time.sleep(0.2)
            except Exception as e:
                log.error(f"Slide {num}: generation failed: {e}", exc_info=True)
                print(f"FAILED ({e})")
                scripts[key] = ""

    SCRIPTS.write_text(json.dumps(scripts, indent=2))
    log.info(f"All scripts saved: {len(scripts)} slides")

    # STEP 3: Synthesise audio
    print(f"\nSTEP 3: Synthesising audio with Edge TTS ({VOICE}, {RATE})...")
    int_scripts = {int(k): v for k, v in scripts.items() if v.strip()}
    t0 = time.time()
    audio_files = batch_synthesise(int_scripts)
    elapsed = time.time() - t0
    print(f"  Done in {elapsed:.0f}s — {len(audio_files)} files generated")
    log.info(f"Batch synthesis done: {len(audio_files)} files in {elapsed:.0f}s")

    # STEP 4: Build manifest
    print("\nSTEP 4: Building manifest...")
    manifest: dict = {}
    for num_str, text in scripts.items():
        num = int(num_str)
        mp3 = audio_files.get(num)
        dur = get_audio_duration(mp3) if mp3 and Path(mp3).exists() else 3.0
        manifest[num_str] = {
            "slide_num":  num,
            "audio_path": mp3,
            "duration":   round(dur, 2),
            "speech":     text,
            "is_demo":    num in DEMO_SLIDES,
            "demo_type":  DEMO_SLIDES.get(num),
        }
        log_slide_event(num, "MANIFEST_ENTRY", f"duration={dur:.1f}s")

    MANIFEST.write_text(json.dumps(manifest, indent=2))
    total_dur = sum(v["duration"] for v in manifest.values())
    log.info(f"Manifest written: {len(manifest)} slides, {total_dur/60:.1f}min")

    print(f"\n{'='*65}")
    print(f"  COMPLETE — {len(audio_files)}/{total} files — {total_dur/60:.1f} min total")
    print(f"  Voice : {VOICE}  |  Rate : {RATE}")
    print(f"  Logs  : logs/ folder")
    print(f"\n  Next step:")
    print(f"    python diagnose.py     # verify everything")
    print(f"    python orchestrator.py # start presentation")
    print(f"{'='*65}\n")


if __name__ == "__main__":
    run_pre_generation("--force" in sys.argv)
