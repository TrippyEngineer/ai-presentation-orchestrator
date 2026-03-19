"""
slide_controller.py — Controls PowerPoint/Impress: open, focus, navigate.

Works on Windows, macOS, and Linux (LibreOffice).
Uses EnumWindows with partial matching on Windows (reliable across all PPT versions).
"""

import os
import sys
import time
import subprocess
import platform
from pathlib import Path

try:
    import pyautogui
    pyautogui.FAILSAFE = False
    pyautogui.PAUSE = 0.05
except ImportError:
    pyautogui = None

KEYWORDS = ["powerpoint", "slide show", "slideshow", "impress", "presentation"]


def open_presentation(pptx_path: str) -> bool:
    """Open PPTX directly into fullscreen slideshow mode."""
    abs_path = str(Path(pptx_path).resolve())
    if not Path(abs_path).exists():
        print(f"  ❌ File not found: {abs_path}")
        return False

    system = platform.system()

    if system == "Windows":
        ppt_locations = [
            r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
            r"C:\Program Files\Microsoft Office\root\Office15\POWERPNT.EXE",
            r"C:\Program Files\Microsoft Office\Office16\POWERPNT.EXE",
            r"C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE",
            r"C:\Program Files (x86)\Microsoft Office\Office16\POWERPNT.EXE",
        ]
        for exe in ppt_locations:
            if Path(exe).exists():
                subprocess.Popen([exe, "/S", abs_path])
                print("  ✅ PowerPoint opened in slideshow mode (/S flag)")
                return True
        os.startfile(abs_path)
        print("  ⚠️ Opened file — press F5 in PowerPoint to go fullscreen.")
        return True

    elif system == "Darwin":
        subprocess.Popen(["open", "-a", "Microsoft PowerPoint", abs_path])
        time.sleep(5)
        script = (
            'tell application "Microsoft PowerPoint" to activate\n'
            'tell application "Microsoft PowerPoint"\n'
            '  start slide show of slide show settings of active presentation\n'
            'end tell'
        )
        result = subprocess.run(["osascript", "-e", script], capture_output=True, timeout=10)
        if result.returncode == 0:
            print("  ✅ Slideshow started via AppleScript")
        else:
            print("  ⚠️ Press F5 in PowerPoint to start the slideshow.")
        return True

    else:
        for cmd in [["libreoffice", "--impress", "--show", abs_path],
                    ["soffice",     "--impress", "--show", abs_path]]:
            try:
                subprocess.Popen(cmd)
                print("  ✅ Opened with LibreOffice Impress")
                return True
            except FileNotFoundError:
                continue
        print("  ❌ LibreOffice not found. Open the file manually.")
        return False


def focus_presentation_window():
    system = platform.system()
    if system == "Windows":
        _focus_windows()
    elif system == "Darwin":
        _focus_macos()
    else:
        _focus_linux()
    time.sleep(0.4)


def _focus_windows():
    try:
        import ctypes
        user32 = ctypes.windll.user32
        found = [None]
        PROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_int, ctypes.c_int)

        def cb(hwnd, _):
            if not user32.IsWindowVisible(hwnd):
                return True
            n = user32.GetWindowTextLengthW(hwnd)
            if n == 0:
                return True
            buf = ctypes.create_unicode_buffer(n + 1)
            user32.GetWindowTextW(hwnd, buf, n + 1)
            if any(kw in buf.value.lower() for kw in KEYWORDS):
                found[0] = hwnd
                return False
            return True

        user32.EnumWindows(PROC(cb), 0)
        if found[0]:
            user32.ShowWindow(found[0], 9)
            user32.SetForegroundWindow(found[0])
            return
    except Exception:
        pass
    try:
        import pygetwindow as gw
        for hint in ["PowerPoint", "Slide Show"]:
            wins = gw.getWindowsWithTitle(hint)
            if wins:
                wins[0].activate()
                return
    except ImportError:
        pass


def _focus_macos():
    try:
        subprocess.run(
            ["osascript", "-e", 'tell application "Microsoft PowerPoint" to activate'],
            capture_output=True, timeout=3
        )
    except Exception:
        pass


def _focus_linux():
    for hint in ["PowerPoint", "Impress", "Presentation", "Slide"]:
        try:
            r = subprocess.run(
                ["xdotool", "search", "--name", hint, "windowactivate", "--sync"],
                capture_output=True, timeout=3
            )
            if r.returncode == 0:
                return
        except (FileNotFoundError, Exception):
            break


def next_slide():
    if pyautogui is None:
        print("  ❌ pyautogui missing. Run: pip install pyautogui")
        return
    focus_presentation_window()
    pyautogui.press("right")
    time.sleep(0.5)


def prev_slide():
    if pyautogui is None:
        return
    focus_presentation_window()
    pyautogui.press("left")
    time.sleep(0.5)


def go_to_slide(slide_num: int):
    if pyautogui is None:
        print("  ❌ pyautogui missing. Run: pip install pyautogui")
        return
    focus_presentation_window()
    time.sleep(0.4)
    for digit in str(slide_num):
        pyautogui.press(digit)
        time.sleep(0.08)
    pyautogui.press("enter")
    time.sleep(0.8)


if __name__ == "__main__":
    import os
    pptx = os.getenv("PPTX_FILE", "presentation.pptx")
    print(f"slide_controller.py test — opening {pptx}")
    open_presentation(pptx)
    time.sleep(6)
    print("Going to slide 1...")
    go_to_slide(1)
    time.sleep(2)
    print("Advancing to slide 2...")
    next_slide()
    print("Done.")
