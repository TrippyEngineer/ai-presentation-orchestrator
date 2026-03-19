"""
orchestrator.py — Fully automatic AI presentation controller.

BEHAVIOUR:
- Press ENTER once to start. Everything runs automatically from there.
- Audio plays per slide. When it finishes, PowerPoint advances and next audio starts.
- SPACEBAR pauses/resumes narration at any time.
- Demo slides: n8n fires automatically, waits, then continues.
- Final slide: voice-driven Q&A (mic → Claude → spoken answer).
- Q quits at any point.

CONTROLS:
    SPACE    Pause / Resume narration
    Q        Quit
    D        Skip demo countdown
"""

import os
import sys
import time
import json
import threading
import subprocess
import platform
from pathlib import Path
from dotenv import load_dotenv

from logger import get_logger, log_slide_event, log_demo_event, log_n8n_event
from voice_engine import VOICE as TTS_VOICE, synthesise as tts_synthesise, play_audio as play_audio_blocking

load_dotenv()
log = get_logger("orchestrator")

# =====================================================================
# CONFIGURATION — all sensitive values come from .env
# =====================================================================
PPTX_FILE       = os.getenv("PPTX_FILE", "presentation.pptx")
PRESENTER_NAME  = os.getenv("PRESENTER_NAME", "the presenter")
PRESENTER_ROLE  = os.getenv("PRESENTER_ROLE", "AI expert")
ORGANIZATION    = os.getenv("ORGANIZATION",   "your organization")

CACHE_DIR       = Path("cache")
MANIFEST        = CACHE_DIR / "manifest.json"

DEMO_WAIT       = {"email": 22, "meeting": 30, "research": 60}
SLIDE_SETTLE_MS = 600
POST_DEMO_PAUSE = 4
QA_VOICE        = TTS_VOICE


# =====================================================================
# AUDIO ENGINE — non-blocking pygame with pause/resume
# =====================================================================

class PresentationAudio:
    def __init__(self):
        self._lock        = threading.Lock()
        self._done_event  = threading.Event()
        self.is_playing   = False
        self.is_paused    = False
        self._pygame_ok   = False
        self._slide_num   = 0
        self._init_pygame()

    def _init_pygame(self):
        try:
            import pygame
            pygame.mixer.pre_init(44100, -16, 2, 2048)
            pygame.mixer.init()
            self._pygame     = pygame
            self._pygame_ok  = True
            log.info("pygame mixer initialised")
        except Exception as e:
            log.warning(f"pygame unavailable ({e}) — blocking fallback")
            self._pygame_ok = False

    def play(self, path: str, slide_num: int = 0):
        if not Path(path).exists():
            log.error(f"Audio missing: {path}")
            self._done_event.set()
            return

        self._done_event.clear()
        self._slide_num  = slide_num
        self.is_playing  = True
        self.is_paused   = False

        if self._pygame_ok:
            try:
                with self._lock:
                    self._pygame.mixer.music.stop()
                    self._pygame.mixer.music.load(path)
                    self._pygame.mixer.music.play()
                threading.Thread(target=self._watch_pygame, args=(path,), daemon=True).start()
                return
            except Exception as e:
                log.error(f"pygame play failed: {e} — falling back")

        threading.Thread(target=self._blocking_thread, args=(path, slide_num), daemon=True).start()

    def _watch_pygame(self, path: str):
        while True:
            with self._lock:
                if not self._pygame.mixer.music.get_busy() and not self.is_paused:
                    break
            time.sleep(0.05)
        self.is_playing = False
        self.is_paused  = False
        self._done_event.set()

    def _blocking_thread(self, path: str, slide_num: int):
        play_audio_blocking(path)
        self.is_playing = False
        self.is_paused  = False
        self._done_event.set()

    def pause(self):
        if not self.is_playing or self.is_paused:
            return
        if self._pygame_ok:
            try:
                self._pygame.mixer.music.pause()
                self.is_paused = True
                log.info(f"Slide {self._slide_num}: PAUSED")
            except Exception as e:
                log.error(f"pause failed: {e}")

    def resume(self):
        if not self.is_paused:
            return
        if self._pygame_ok:
            try:
                self._pygame.mixer.music.unpause()
                self.is_paused = False
                log.info(f"Slide {self._slide_num}: RESUMED")
            except Exception as e:
                log.error(f"resume failed: {e}")

    def toggle_pause(self):
        if self.is_paused:
            self.resume()
        else:
            self.pause()

    def stop(self):
        if self._pygame_ok:
            try:
                self._pygame.mixer.music.stop()
            except Exception:
                pass
        self.is_playing = False
        self.is_paused  = False
        self._done_event.set()

    def wait(self, check_fn=None, poll_secs: float = 0.05) -> str:
        while not self._done_event.is_set():
            if check_fn:
                cmd = check_fn()
                if cmd == "pause":
                    self.toggle_pause()
                elif cmd == "stop":
                    self.stop()
                    return "stopped"
                elif cmd == "quit":
                    self.stop()
                    return "quit"
            self._done_event.wait(timeout=poll_secs)
        return "done"

    @property
    def active(self) -> bool:
        return self.is_playing or self.is_paused


# =====================================================================
# KEYBOARD LISTENER
# =====================================================================

class KeyboardListener:
    def __init__(self):
        self.space      = threading.Event()
        self.quit       = threading.Event()
        self.skip_demo  = threading.Event()
        self._pynput_ok = False
        self._listener  = None
        self._start()

    def _start(self):
        try:
            from pynput import keyboard as pk

            def on_press(key):
                try:
                    if key == pk.Key.space:
                        self.space.set()
                    elif hasattr(key, "char") and key.char:
                        c = key.char.lower()
                        if c == "q": self.quit.set()
                        elif c == "d": self.skip_demo.set()
                except Exception:
                    pass

            self._listener = pk.Listener(on_press=on_press, suppress=False)
            self._listener.start()
            self._pynput_ok = True
            log.info("pynput keyboard listener started")
        except ImportError:
            log.warning("pynput not installed — falling back to terminal keyboard")
            threading.Thread(target=self._fallback, daemon=True).start()

    def _fallback(self):
        try:
            import msvcrt
            while not self.quit.is_set():
                if msvcrt.kbhit():
                    ch = msvcrt.getch()
                    if ch == b" ":    self.space.set()
                    elif ch.lower() == b"q": self.quit.set()
                    elif ch.lower() == b"d": self.skip_demo.set()
                time.sleep(0.03)
        except ImportError:
            import tty, termios, select
            fd = sys.stdin.fileno()
            old = termios.tcgetattr(fd)
            try:
                tty.setraw(fd)
                while not self.quit.is_set():
                    if select.select([sys.stdin], [], [], 0.05)[0]:
                        ch = sys.stdin.read(1)
                        if ch == " ":      self.space.set()
                        elif ch.lower() == "q": self.quit.set()
                        elif ch.lower() == "d": self.skip_demo.set()
            finally:
                termios.tcsetattr(fd, termios.TCSADRAIN, old)

    def consume_space(self) -> bool:
        if self.space.is_set():
            self.space.clear()
            return True
        return False

    def consume_skip(self) -> bool:
        if self.skip_demo.is_set():
            self.skip_demo.clear()
            return True
        return False

    def stop(self):
        if self._listener:
            try: self._listener.stop()
            except Exception: pass


# =====================================================================
# POWERPOINT CONTROL
# =====================================================================

def open_powerpoint() -> bool:
    if not Path(PPTX_FILE).exists():
        log.error(f"PPTX not found: {PPTX_FILE}")
        print(f"  ❌ File not found: {PPTX_FILE}")
        return False

    path   = str(Path(PPTX_FILE).absolute())
    system = platform.system()

    if system == "Windows":
        candidates = [
            r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
            r"C:\Program Files\Microsoft Office\root\Office15\POWERPNT.EXE",
            r"C:\Program Files\Microsoft Office\Office16\POWERPNT.EXE",
            r"C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE",
            r"C:\Program Files (x86)\Microsoft Office\Office16\POWERPNT.EXE",
        ]
        for exe in candidates:
            if Path(exe).exists():
                subprocess.Popen([exe, "/S", path])
                print("  ✅ PowerPoint opened in slideshow mode")
                return True
        os.startfile(path)
        print("  ⚠️ PowerPoint opened — press F5 to go fullscreen")
        return True

    elif system == "Darwin":
        subprocess.Popen(["open", "-a", "Microsoft PowerPoint", path])
        time.sleep(5)
        result = subprocess.run(
            ["osascript", "-e",
             'tell application "Microsoft PowerPoint" to activate\n'
             'tell application "Microsoft PowerPoint"\n'
             '  start slide show of slide show settings of active presentation\n'
             'end tell'],
            capture_output=True, timeout=10
        )
        if result.returncode == 0:
            print("  ✅ Slideshow started via AppleScript")
        else:
            print("  ⚠️ Press F5 in PowerPoint to start the slideshow")
        return True

    else:
        for cmd in [["libreoffice", "--impress", "--show", path],
                    ["soffice",     "--impress", "--show", path]]:
            try:
                subprocess.Popen(cmd)
                print("  ✅ Opened with LibreOffice")
                return True
            except FileNotFoundError:
                continue
        print("  ❌ LibreOffice not found — open the file manually")
        return False


def _focus_ppt():
    system = platform.system()
    try:
        if system == "Windows":
            import ctypes
            user32  = ctypes.windll.user32
            found   = [None]
            keywords = ["powerpoint", "slide show", "slideshow"]
            PROC    = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_int, ctypes.c_int)

            def cb(hwnd, _):
                if not user32.IsWindowVisible(hwnd): return True
                n = user32.GetWindowTextLengthW(hwnd)
                if n == 0: return True
                buf = ctypes.create_unicode_buffer(n + 1)
                user32.GetWindowTextW(hwnd, buf, n + 1)
                if any(kw in buf.value.lower() for kw in keywords):
                    found[0] = hwnd
                    return False
                return True

            user32.EnumWindows(PROC(cb), 0)
            if found[0]:
                user32.ShowWindow(found[0], 9)
                user32.SetForegroundWindow(found[0])
                time.sleep(0.3)
                return

        elif system == "Darwin":
            subprocess.run(
                ["osascript", "-e", 'tell application "Microsoft PowerPoint" to activate'],
                capture_output=True, timeout=3
            )
            time.sleep(0.3)
            return

        else:
            for hint in ["PowerPoint", "Impress", "Slide"]:
                r = subprocess.run(
                    ["xdotool", "search", "--name", hint, "windowactivate", "--sync"],
                    capture_output=True, timeout=3
                )
                if r.returncode == 0:
                    return
    except Exception as e:
        log.debug(f"_focus_ppt: {e}")


def _send_right_to_ppt():
    try:
        import pyautogui
        pyautogui.FAILSAFE = False
        _focus_ppt()
        time.sleep(0.25)
        pyautogui.press("right")
        time.sleep(0.4)
    except ImportError:
        log.warning("pyautogui not installed — cannot advance slide automatically")
        print("  ⚠️ Install pyautogui: pip install pyautogui")


def _go_to_slide_1():
    try:
        import pyautogui
        pyautogui.FAILSAFE = False
        _focus_ppt()
        time.sleep(0.3)
        pyautogui.press("1")
        time.sleep(0.08)
        pyautogui.press("enter")
        time.sleep(0.8)
    except ImportError:
        pass


# =====================================================================
# n8n DEMO TRIGGERS
# =====================================================================

def _post(key: str, payload: dict, timeout: int = 30) -> bool:
    url = os.getenv(key)
    if not url:
        log.warning(f"{key} not in .env")
        print(f"  ⚠️ {key} not set — skipping demo trigger")
        return False
    try:
        import requests
        r  = requests.post(url, json=payload, timeout=timeout)
        ok = r.status_code in (200, 201, 202)
        log_n8n_event(key, r.status_code)
        if ok:
            print(f"  ✅ {key}: HTTP {r.status_code}")
        else:
            print(f"  ⚠️ {key}: HTTP {r.status_code} — check workflow is ACTIVE")
        return ok
    except Exception as e:
        log_n8n_event(key, error=str(e))
        print(f"  ❌ {key} failed: {e}")
        return False


def trigger_email_demo() -> bool:
    return _post("N8N_WEBHOOK_EMAIL", {
        "Subject": "Q1 Programme Update Request",
        "From":    "XYX <xyx@example.org>",
        "body":    (
            "Hi, we are preparing our Q1 donor review and would appreciate "
            "an update on programme outcomes in the target region. Specifically "
            "we need coverage rates, key achievements, and any emerging implementation "
            "challenges. Could you send this by Friday? Best regards, XYX"
        ),
        "threadId": ""
    })


def trigger_meeting_demo() -> bool:
    tx = Path("demo_data/meeting_transcript.txt")
    if not tx.exists():
        print(f"  ❌ Missing: {tx}")
        return False
    return _post("N8N_WEBHOOK_MEETING", {
        "Meeting Name":       "Country Team Meeting (Sample)",
        "Meeting Transcript": tx.read_text()
    })


def trigger_research_demo() -> bool:
    return _post("N8N_WEBHOOK_RESEARCH", {
        "Input Mode":        "Research Question",
        "Research Question": (
            "What is the evidence on zinc supplementation effectiveness "
            "for reducing childhood diarrhea mortality and what are the barriers to scale-up?"
        ),
        "Meeting Transcript": "",
        "Program Area":       "Nutrition",
        "Geographic Region":  "Sub-Saharan Africa",
        "Audience Type":      "Donor Brief",
        "Urgency":            "Standard"
    }, timeout=90)


DEMO_TRIGGERS = {
    "email":    trigger_email_demo,
    "meeting":  trigger_meeting_demo,
    "research": trigger_research_demo,
}


def _demo_countdown(seconds: int, kbd: KeyboardListener):
    end = time.time() + seconds
    while time.time() < end:
        if kbd.consume_skip() or kbd.quit.is_set():
            print("  ⏭ Demo wait skipped")
            return
        remaining = int(end - time.time())
        print(f"  ⏳ Demo running — {remaining:3d}s remaining (D=skip) ", end="\r", flush=True)
        time.sleep(0.5)
    print("  ✅ Demo complete                                   ")


def _write_demo_status(demo_type: str, status: str):
    try:
        (CACHE_DIR / "demo_status.json").write_text(
            json.dumps({"demo": demo_type, "status": status, "ts": time.time()})
        )
    except Exception:
        pass


# =====================================================================
# VOICE Q&A — mic → Google STT → Claude → Edge TTS → speaker
# =====================================================================

def _listen(timeout: int = 12) -> str:
    try:
        import speech_recognition as sr
    except ImportError:
        print("  ❌ pip install SpeechRecognition pyaudio")
        return ""

    rec = sr.Recognizer()
    rec.energy_threshold       = 300
    rec.dynamic_energy_threshold = True
    rec.pause_threshold        = 1.2

    try:
        with sr.Microphone() as src:
            print("  🎤 Listening...", flush=True)
            rec.adjust_for_ambient_noise(src, duration=0.4)
            audio = rec.listen(src, timeout=timeout, phrase_time_limit=20)
            text  = rec.recognize_google(audio, language="en-US")
            log.info(f"STT: {text[:80]}")
            return text.strip()
    except sr.WaitTimeoutError:
        return ""
    except sr.UnknownValueError:
        print("  ⚠️ Didn't catch that — try again")
        return ""
    except Exception as e:
        log.error(f"STT error: {e}")
        return ""


def _claude_answer(question: str, history: list) -> str:
    try:
        from anthropic import Anthropic
        client = Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
    except Exception:
        return "I'm having a technical issue right now."

    system = (
        f"You are {PRESENTER_NAME}, a {PRESENTER_ROLE} at {ORGANIZATION}. "
        f"You just finished a live workshop on AI tools and automation. "
        f"Answer questions in 3-4 spoken sentences. No bullet points. No markdown. "
        f"Plain conversational English. Direct. If unsure, say so honestly."
    )

    history.append({"role": "user", "content": question})
    try:
        from anthropic import Anthropic
        client   = Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
        resp     = client.messages.create(
            model="claude-opus-4-5",
            max_tokens=200,
            system=system,
            messages=history
        )
        answer = resp.content[0].text.strip()
        history.append({"role": "assistant", "content": answer})
        log.info(f"Q&A answer: {answer[:80]}")
        return answer
    except Exception as e:
        log.error(f"Claude Q&A error: {e}")
        return "I'm having a technical issue — let me answer that directly."


def _speak(text: str):
    tmp = str(CACHE_DIR / "_qa.mp3")
    if tts_synthesise(text, tmp, voice=QA_VOICE):
        play_audio_blocking(tmp)
    else:
        print(f"  [TTS failed] {text}")


def run_voice_qa(kbd: KeyboardListener):
    log.info("Voice Q&A started")
    print("\n" + "═"*60)
    print("  VOICE Q&A — audience speaks, AI answers aloud")
    print("  Press Q to end Q&A")
    print("═"*60 + "\n")

    _speak("We have a few minutes for questions. Please speak clearly and I'll answer right away.")

    history   = []
    q_count   = 0
    timeouts  = 0

    while not kbd.quit.is_set():
        print(f"\n  [{q_count+1}] Waiting for question (Q to end)...")
        question = _listen(timeout=12)

        if not question:
            timeouts += 1
            if timeouts >= 3:
                _speak("It looks like we're done. Thank you all for a great session.")
                break
            _speak("Go ahead with your question.")
            continue

        timeouts = 0
        q_count += 1
        print(f"  Q: {question}")

        if any(p in question.lower() for p in ["that's all", "no more", "we're done", "thank you all"]):
            _speak("Thank you. It was a pleasure presenting to all of you today.")
            break

        print("  Generating answer...", flush=True)
        answer = _claude_answer(question, history)
        print(f"  A: {answer[:80]}{'...' if len(answer) > 80 else ''}")
        _speak(answer)

    log.info(f"Voice Q&A ended: {q_count} questions")
    print(f"\n  Q&A ended — {q_count} question(s) answered\n")


# =====================================================================
# MAIN
# =====================================================================

def validate_n8n():
    try:
        import requests
        for key in ["N8N_WEBHOOK_EMAIL", "N8N_WEBHOOK_MEETING", "N8N_WEBHOOK_RESEARCH"]:
            url = os.getenv(key, "")
            if not url:
                print(f"  ⚠️ {key} not set in .env")
                continue
            try:
                base = url.split("/webhook/")[0] + "/healthz"
                r    = requests.get(base, timeout=4)
                mark = "✅" if r.status_code == 200 else "⚠️ "
                print(f"  {mark} n8n {key.split('_')[-1].lower()}: HTTP {r.status_code}")
            except Exception:
                print(f"  ❌ n8n unreachable for {key}")
    except ImportError:
        print("  ⚠️ requests not installed — skipping n8n check")


def run_presentation():
    print("\n" + "═"*60)
    print("  AI PRESENTATION ORCHESTRATOR")
    print("═"*60)

    if not MANIFEST.exists():
        print("\n  ❌ cache/manifest.json not found")
        print("  Run: python pre_generate.py\n")
        sys.exit(1)

    manifest      = json.loads(MANIFEST.read_text())
    sorted_slides = sorted(manifest.items(), key=lambda x: int(x[0]))
    total         = len(sorted_slides)
    total_dur     = sum(v.get("duration", 0) for v in manifest.values())
    demo_slides   = [int(k) for k, v in manifest.items() if v.get("is_demo")]

    missing = [k for k, v in manifest.items()
               if v.get("audio_path") and not Path(v["audio_path"]).exists()]
    if missing:
        print(f"\n  ❌ Missing audio for slides: {missing}")
        print("  Run: python pre_generate.py --force\n")
        sys.exit(1)

    log.info(f"Starting: {total} slides, {total_dur/60:.1f}min, demos={demo_slides}")

    print(f"""
Slides      : {total}
Duration    : ~{total_dur/60:.0f} min
Demo slides : {demo_slides}

CONTROLS
──────────────────────────────
SPACE   Pause / Resume narration
D       Skip demo countdown
Q       Quit

Everything else is automatic.
""")

    print("  Checking n8n...")
    validate_n8n()
    print()
    input("  ▶ Press ENTER to open the presentation and start... ")

    opened = open_powerpoint()
    wait_s = 7 if opened else 0
    if opened:
        print(f"  ⏳ Waiting {wait_s}s for the app to load...")
        time.sleep(wait_s)
    else:
        input("  Open the PPTX manually in fullscreen, then press ENTER... ")

    _go_to_slide_1()
    time.sleep(1.0)

    audio = PresentationAudio()
    kbd   = KeyboardListener()

    if kbd._pynput_ok:
        print("  ✅ Keyboard listener active (SPACE works even when PPT has focus)\n")
    else:
        print("  ⚠️ pynput missing — keep this terminal focused for SPACE to work\n")
        print("  Install for better control: pip install pynput\n")

    print(f"  {'─'*55}")
    print(f"  PRESENTATION STARTING NOW")
    print(f"  {'─'*55}\n")
    log.info("Presentation loop starting")

    time.sleep(1.5)

    for idx, (num_str, entry) in enumerate(sorted_slides):

        if kbd.quit.is_set():
            break

        slide_num  = int(num_str)
        audio_path = entry.get("audio_path", "")
        duration   = entry.get("duration",   3.0)
        is_demo    = entry.get("is_demo",     False)
        demo_type  = entry.get("demo_type",   "")
        is_final   = (idx == total - 1)

        demo_tag = f" 🔴 DEMO:{demo_type.upper()}" if is_demo else ""
        print(f"  SLIDE {slide_num:2d}/{total}{demo_tag} ({duration:.0f}s)", flush=True)
        log_slide_event(slide_num, "START")

        if idx > 0:
            _send_right_to_ppt()
            time.sleep(SLIDE_SETTLE_MS / 1000)

        if audio_path and Path(audio_path).exists():
            audio.play(audio_path, slide_num)

            def _check_kbd():
                if kbd.quit.is_set():       return "quit"
                if kbd.consume_space():
                    state = "⏸ PAUSED" if not audio.is_paused else "▶ Resumed"
                    print(f"\r  {state}          ", flush=True)
                    return "pause"
                return None

            result = audio.wait(check_fn=_check_kbd, poll_secs=0.05)
            if result == "quit":
                break
        else:
            log.warning(f"Slide {slide_num}: no audio — sleeping {duration:.0f}s")
            time.sleep(duration)

        log_slide_event(slide_num, "DONE")

        if is_demo and demo_type in DEMO_TRIGGERS:
            print(f"\n  🚀 Triggering {demo_type.upper()} demo...")
            log_demo_event(demo_type, "TRIGGERING")
            _write_demo_status(demo_type, "running")

            ok = DEMO_TRIGGERS[demo_type]()
            if ok:
                log_demo_event(demo_type, "OK")
                _demo_countdown(DEMO_WAIT[demo_type], kbd)
                _write_demo_status(demo_type, "complete")
            else:
                log_demo_event(demo_type, "FAILED")
                print("  ❌ Trigger failed — fire manually from localhost:5678")

            if kbd.quit.is_set():
                break
            time.sleep(POST_DEMO_PAUSE)
            print()

        if is_final:
            print("\n  ✅ Presentation complete.\n")
            ans = input("  Start voice Q&A? (y/n): ").strip().lower()
            if ans == "y":
                run_voice_qa(kbd)
            break

    audio.stop()
    kbd.stop()
    log.info("Presentation ended")
    print(f"\n  {'─'*55}")
    print("  SESSION ENDED | logs → logs/")
    print(f"  {'─'*55}\n")


if __name__ == "__main__":
    run_presentation()
