"""
logger.py — Centralized logging for the AI Presentation Orchestrator.

All modules import from here. Logs go to:
- logs/presentation_YYYYMMDD.log (rotating, DEBUG level)
- Console stderr (WARNING+ only, to keep terminal clean during presentation)

Usage:
    from logger import get_logger
    log = get_logger(__name__)
    log.info("Slide 3 audio started")
    log.error("n8n webhook failed", exc_info=True)
"""

import logging
import logging.handlers
from pathlib import Path
from datetime import datetime

LOG_DIR = Path("logs")
LOG_DIR.mkdir(exist_ok=True)

_loggers = {}


def get_logger(name: str) -> logging.Logger:
    """Return a named logger. Creates it on first call, reuses after."""
    if name in _loggers:
        return _loggers[name]

    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)

    if logger.handlers:
        _loggers[name] = logger
        return logger

    log_file = LOG_DIR / f"presentation_{datetime.now().strftime('%Y%m%d')}.log"
    fh = logging.handlers.RotatingFileHandler(
        log_file,
        maxBytes=10 * 1024 * 1024,
        backupCount=7,
        encoding="utf-8",
    )
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(logging.Formatter(
        "%(asctime)s [%(name)-20s] %(levelname)-8s %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    ))

    ch = logging.StreamHandler()
    ch.setLevel(logging.WARNING)
    ch.setFormatter(logging.Formatter("%(levelname)s [%(name)s]: %(message)s"))

    logger.addHandler(fh)
    logger.addHandler(ch)

    _loggers[name] = logger
    return logger


def get_presentation_logger() -> logging.Logger:
    return get_logger("presentation")


def log_slide_event(slide_num: int, event: str, detail: str = ""):
    log = get_logger("slides")
    msg = f"SLIDE {slide_num:02d} | {event}"
    if detail:
        msg += f" | {detail}"
    log.info(msg)


def log_audio_event(slide_num: int, event: str, detail: str = ""):
    log = get_logger("audio")
    msg = f"SLIDE {slide_num:02d} | {event}"
    if detail:
        msg += f" | {detail}"
    log.info(msg)


def log_demo_event(demo_type: str, event: str, detail: str = ""):
    log = get_logger("demo")
    msg = f"DEMO:{demo_type.upper()} | {event}"
    if detail:
        msg += f" | {detail}"
    log.info(msg)


def log_n8n_event(webhook_key: str, status_code: int = None,
                  error: str = None, payload_summary: str = ""):
    log = get_logger("n8n")
    if error:
        log.error(f"{webhook_key} | FAILED | {error}")
    elif status_code:
        level = logging.INFO if 200 <= status_code < 300 else logging.WARNING
        log.log(level, f"{webhook_key} | HTTP {status_code} | {payload_summary}")
    else:
        log.info(f"{webhook_key} | TRIGGERED | {payload_summary}")


if __name__ == "__main__":
    log = get_logger("test")
    log.debug("Debug message — in file only")
    log.info("Info message — in file only")
    log.warning("Warning message — visible on console")
    log.error("Error message — visible on console")
    log_slide_event(3, "AUDIO_START", "duration=12.4s")
    log_demo_event("email", "TRIGGERED", "HTTP 200")
    print(f"\nLog written to: {LOG_DIR}/")
    print("Check logs/ folder to verify.")
