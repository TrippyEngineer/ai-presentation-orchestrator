"""
voice_engine.py — FREE voice synthesis via Microsoft Edge TTS.

Configure via .env:
  TTS_VOICE   Edge TTS voice name (default: en-US-AndrewNeural)
  TTS_RATE    Speech rate offset  (default: -3%)

Fallback chain: pygame → ffplay → afplay (macOS) → winmm (Windows) → mpg123 → sleep
"""

import os
import asyncio
import subprocess
import sys
import time
import threading
from pathlib import Path

from core.logger import get_logger, log_audio_event
from dotenv import load_dotenv

load_dotenv()
log = get_logger("voice_engine")

VOICE     = os.getenv("TTS_VOICE", "en-US-AndrewNeural")
RATE      = os.getenv("TTS_RATE",  "-3%")
CACHE_DIR = Path("cache")
CACHE_DIR.mkdir(exist_ok=True)


async def _synthesise_async(text: str, output_path: str, voice: str = VOICE) -> bool:
    try:
        import edge_tts
        communicate = edge_tts.Communicate(text, voice, rate=RATE)
        await communicate.save(output_path)
        log.debug(f"Synthesised → {output_path} ({len(text.split())} words, voice={voice})")
        return True
    except ImportError:
        log.error("edge-tts not installed. Run: pip install edge-tts")
        return False
    except Exception as e:
        log.error(f"Edge TTS error for {output_path}: {e}", exc_info=True)
        return False


def synthesise(text: str, output_path: str, voice: str = VOICE) -> bool:
    if not text.strip():
        log.warning(f"synthesise() called with empty text for {output_path}")
        return False
    return asyncio.run(_synthesise_async(text, output_path, voice))


class AudioPlayer:
    """
    Non-blocking pygame audio player with pause/resume support.

    Usage:
        player = AudioPlayer()
        player.play("cache/slide_03.mp3")   # returns immediately
        player.pause()
        player.resume()
        player.wait_until_done()
        player.stop()
    """

    def __init__(self):
        self._lock = threading.Lock()
        self.is_playing = False
        self.is_paused  = False
        self._done      = threading.Event()
        self._init_pygame()
        log.debug("AudioPlayer initialised")

    def _init_pygame(self):
        try:
            import pygame
            if not pygame.mixer.get_init():
                pygame.mixer.pre_init(44100, -16, 2, 2048)
                pygame.mixer.init()
            self._pygame = pygame
            log.debug("pygame mixer ready")
        except ImportError:
            self._pygame = None
            log.warning("pygame not available — will fall back to blocking players")

    def play(self, path: str, slide_num: int = 0) -> bool:
        if not Path(path).exists():
            log.error(f"Audio file not found: {path}")
            return False

        self._done.clear()
        self._slide_num = slide_num

        if self._pygame:
            try:
                with self._lock:
                    self._pygame.mixer.music.stop()
                    self._pygame.mixer.music.load(path)
                    self._pygame.mixer.music.play()
                self.is_playing = True
                self.is_paused  = False
                threading.Thread(target=self._monitor_pygame,
                                 args=(path,), daemon=True).start()
                log_audio_event(slide_num, "PLAY_START", f"file={Path(path).name}")
                return True
            except Exception as e:
                log.error(f"pygame play error: {e}", exc_info=True)

        threading.Thread(target=self._play_blocking,
                         args=(path, slide_num), daemon=True).start()
        self.is_playing = True
        self.is_paused  = False
        return True

    def _monitor_pygame(self, path: str):
        while self._pygame.mixer.music.get_busy():
            time.sleep(0.05)
        with self._lock:
            self.is_playing = False
            self.is_paused  = False
        log_audio_event(self._slide_num, "PLAY_DONE", f"file={Path(path).name}")
        self._done.set()

    def _play_blocking(self, path: str, slide_num: int):
        log.debug(f"Slide {slide_num}: blocking fallback playback")
        _play_audio_blocking(path)
        with self._lock:
            self.is_playing = False
            self.is_paused  = False
        log_audio_event(slide_num, "PLAY_DONE_BLOCKING", f"file={Path(path).name}")
        self._done.set()

    def pause(self):
        if not self.is_playing or self.is_paused:
            return
        if self._pygame:
            try:
                self._pygame.mixer.music.pause()
                self.is_paused = True
                log_audio_event(self._slide_num, "PAUSED")
            except Exception as e:
                log.error(f"pause error: {e}")

    def resume(self):
        if not self.is_paused:
            return
        if self._pygame:
            try:
                self._pygame.mixer.music.unpause()
                self.is_paused = False
                log_audio_event(self._slide_num, "RESUMED")
            except Exception as e:
                log.error(f"resume error: {e}")

    def toggle_pause(self):
        if self.is_paused:
            self.resume()
        else:
            self.pause()

    def stop(self):
        if self._pygame:
            try:
                self._pygame.mixer.music.stop()
            except Exception:
                pass
        with self._lock:
            self.is_playing = False
            self.is_paused  = False
        self._done.set()

    def wait_until_done(self, timeout: float = 300.0) -> bool:
        finished = self._done.wait(timeout=timeout)
        if not finished:
            log.warning(f"Audio wait timeout ({timeout}s) — moving on")
        return finished

    @property
    def active(self) -> bool:
        return self.is_playing and not self.is_paused


def play_audio(path: str, duration_hint: float = 0) -> float:
    """Blocking play. Returns elapsed seconds."""
    if not Path(path).exists():
        log.warning(f"Audio not found: {path} — sleeping {duration_hint:.1f}s")
        if duration_hint > 0:
            time.sleep(duration_hint)
        return duration_hint

    abs_path = str(Path(path).resolve())
    start    = time.time()
    played   = _play_audio_blocking(abs_path)

    if not played:
        log.warning(f"All audio players failed for {path} — sleeping {duration_hint:.1f}s")
        time.sleep(max(duration_hint, 1))

    elapsed = time.time() - start
    log.debug(f"play_audio: {Path(path).name} finished in {elapsed:.1f}s")
    return elapsed


def _play_audio_blocking(abs_path: str) -> bool:
    if _play_pygame(abs_path):    return True
    if _play_ffplay(abs_path):    return True
    import platform
    if platform.system() == "Darwin"  and _play_afplay(abs_path):  return True
    if platform.system() == "Windows" and _play_winmm(abs_path):   return True
    if _play_mpg123(abs_path): return True
    return False


def _play_pygame(path: str) -> bool:
    try:
        import pygame
        if not pygame.mixer.get_init():
            pygame.mixer.pre_init(44100, -16, 2, 2048)
            pygame.mixer.init()
        pygame.mixer.music.stop()
        pygame.mixer.music.load(path)
        pygame.mixer.music.play()
        while pygame.mixer.music.get_busy():
            time.sleep(0.05)
        pygame.mixer.music.stop()
        return True
    except ImportError:
        return False
    except Exception as e:
        log.debug(f"pygame blocking failed: {e}")
        try:
            import pygame
            pygame.mixer.quit()
        except Exception:
            pass
        return False


def _play_ffplay(path: str) -> bool:
    try:
        result = subprocess.run(
            ["ffplay", "-nodisp", "-autoexit", "-loglevel", "quiet", path],
            timeout=600, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )
        return result.returncode == 0
    except (FileNotFoundError, subprocess.TimeoutExpired):
        return False


def _play_afplay(path: str) -> bool:
    try:
        result = subprocess.run(
            ["afplay", path],
            timeout=600, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )
        return result.returncode == 0
    except (FileNotFoundError, subprocess.TimeoutExpired):
        return False


def _play_winmm(path: str) -> bool:
    try:
        import ctypes
        winmm = ctypes.windll.winmm
        alias = "presentation_audio"
        ret = winmm.mciSendStringW(f'open "{path}" type mpegvideo alias {alias}', None, 0, None)
        if ret != 0:
            return False
        winmm.mciSendStringW(f"play {alias} wait", None, 0, None)
        winmm.mciSendStringW(f"close {alias}", None, 0, None)
        return True
    except Exception:
        return False


def _play_mpg123(path: str) -> bool:
    try:
        result = subprocess.run(
            ["mpg123", "-q", path],
            timeout=600, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )
        return result.returncode == 0
    except (FileNotFoundError, subprocess.TimeoutExpired):
        return False


def batch_synthesise(scripts: dict, voice: str = VOICE) -> dict:
    """Synthesise all scripts. Skips already-cached files."""
    results = {}
    for slide_num, text in sorted(scripts.items()):
        if not text.strip():
            log.debug(f"Slide {slide_num}: empty script — skipping")
            print(f"  ⏭ Slide {slide_num}: empty — skipping")
            continue
        mp3_path = str(CACHE_DIR / f"slide_{slide_num:02d}.mp3")
        if Path(mp3_path).exists():
            print(f"  ✅ Slide {slide_num}: already cached")
            results[slide_num] = mp3_path
            continue
        print(f"  🔊 Slide {slide_num}: synthesising... ", end="", flush=True)
        ok = synthesise(text, mp3_path, voice)
        if ok:
            size = Path(mp3_path).stat().st_size
            print(f"✅ ({size//1024}KB)")
            results[slide_num] = mp3_path
        else:
            print("❌ FAILED")
            log.error(f"Slide {slide_num}: synthesis FAILED")
    return results


def get_audio_duration(path: str) -> float:
    try:
        result = subprocess.run(
            ["ffprobe", "-v", "error", "-show_entries", "format=duration",
             "-of", "default=noprint_wrappers=1:nokey=1", path],
            capture_output=True, text=True, timeout=10
        )
        if result.returncode == 0 and result.stdout.strip():
            return float(result.stdout.strip())
    except Exception:
        pass
    try:
        return max(os.path.getsize(path) / 6000, 1.0)
    except Exception:
        return 5.0


def speak_and_wait(text: str, slide_num: int, voice: str = VOICE) -> float:
    tmp_path = str(CACHE_DIR / f"_tmp_slide_{slide_num}.mp3")
    ok = synthesise(text, tmp_path, voice)
    if not ok:
        fallback = len(text.split()) / 2.5
        time.sleep(fallback)
        return fallback
    return play_audio(tmp_path)


if __name__ == "__main__":
    print("\n" + "="*60)
    print("  AUDIO PLAYBACK TEST")
    print(f"  Voice: {VOICE}   Rate: {RATE}")
    print("="*60)
    test_text = (
        "Hello. This is an audio playback test. "
        "If you can hear this clearly, the voice configuration is working correctly."
    )
    test_path = str(CACHE_DIR / "_audio_test.mp3")
    print("\nGenerating test audio via Edge TTS (needs internet)...")
    ok = synthesise(test_text, test_path)
    if not ok:
        print("❌ Synthesis failed. Check: pip install edge-tts, and internet connection.")
        sys.exit(1)
    print("Playing now...")
    duration = play_audio(test_path)
    if duration > 0.5:
        print(f"\n✅ Audio works! ({duration:.1f}s)  Voice: {VOICE}")
    else:
        print("\n❌ Audio may not have played. Check speakers and volume.")
