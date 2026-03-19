# RUNBOOK — Presentation Day Checklist

## FOLDER STRUCTURE

```
AI_Presentation/
  orchestrator.py
  pre_generate.py
  script_agent.py
  slide_controller.py
  slide_reader.py
  voice_engine.py
  n8n_trigger.py
  diagnose.py
  logger.py
  presentation.pptx          ← your slide deck
  .env                        ← never commit
  demo_data/
    meeting_transcript.txt   ← sample transcript
  cache/                     ← created automatically by pre_generate.py
    manifest.json
    slide_01.mp3
    slide_02.mp3
    ... etc
```

## .env FILE MUST CONTAIN

```
ANTHROPIC_API_KEY=sk-ant-your-key-here
HEYGEN_API_KEY=your-heygen-key
HEYGEN_AVATAR_ID=your-avatar-id
HEYGEN_VOICE_ID=your-voice-id
GOOGLE_SLIDES_PRESENTATION_ID=your-presentation-id
N8N_WEBHOOK_EMAIL=http://localhost:5678/webhook/email-demo
N8N_WEBHOOK_MEETING=http://localhost:5678/webhook/meeting-demo
N8N_WEBHOOK_RESEARCH=http://localhost:5678/webhook/research-demo
SLACK_WEBHOOK_URL=https://hooks.slack.com/services/YOUR/WEBHOOK/URL
```

---

### STEP 1 — Open one terminal, go to your folder

```
cd AI_Presentation
```

Keep this terminal open for everything below.

---

### STEP 2 — Install all dependencies (run once)

```
pip install edge-tts anthropic python-dotenv python-pptx pyautogui requests pygame pynput
```

---

### STEP 3 — Install ffmpeg (needed for audio duration detection)

**Windows:**
1. Go to: https://www.gyan.dev/ffmpeg/builds/
2. Download: `ffmpeg-release-essentials.zip`
3. Unzip anywhere, e.g. `C:\ffmpeg`
4. Add `C:\ffmpeg\bin` to your Windows PATH:
   - Search "environment variables" in Start
   - System Properties → Environment Variables
   - Under "System variables" find "Path" → Edit → New → paste `C:\ffmpeg\bin`
   - Click OK on all windows
5. Open a NEW terminal and test: `ffplay -version`

**Mac:** `brew install ffmpeg`
**Linux:** `sudo apt install ffmpeg`

---

### STEP 4 — Create demo_data folder and transcript

```
mkdir demo_data
```

Create `demo_data/meeting_transcript.txt` and paste in a realistic sample meeting transcript.
Do not use real meeting data — create a representative sample.

---

### STEP 5 — Pre-generate all audio (takes 5–10 minutes)

```
python pre_generate.py
```

Watch the output. It will:
- Read your slides
- Call Claude to write scripts for each slide
- Generate MP3 files via Edge TTS (or HeyGen if configured)
- Save everything to `cache/`

When done you will see: `COMPLETE — N/N files — X.X min total`

If it fails on a slide, just re-run. Cached slides are skipped automatically.

---

### STEP 6 — Run the diagnostic

```
python diagnose.py
```

Every line should say ✅. If anything says ❌, fix it using the FIX instruction shown.
The audio test will play a real slide out loud — confirm you hear it.

---

### STEP 7 — Do ONE full dry run

Start n8n in a separate terminal first:

```
npx n8n
```

Then in your main terminal:

```
python orchestrator.py
```

- Presentation opens automatically in fullscreen
- Audio starts playing
- Slides advance automatically
- At demo slides, n8n triggers fire
- Let it run to the end

After the dry run: close everything, restart, and you are ready.

---

## ============================================================
## PRESENTATION DAY — EXACT SEQUENCE
## ============================================================

**TERMINAL 1** — Start n8n (keep this window open all day):

```
npx n8n
```

Wait for `n8n ready on port 5678`, then leave it.

Open browser → http://localhost:5678
Check all 3 workflows: **Email Pipeline**, **Meeting Pipeline**, **Research Engine**
Each must show a **GREEN Active** toggle. If grey/inactive → click to activate.

---

### 30 MINUTES BEFORE

**TERMINAL 2** — Run diagnostic one more time:

```
python diagnose.py
```

All lines must say ✅. Fix anything that doesn't.

---

### 15 MINUTES BEFORE

- Connect laptop to projector/screen
- Set display to **EXTEND** (presenter notes on your screen) or **DUPLICATE** (audience sees same)
- Turn laptop volume up to 80–100%
- Check external speakers if the room has them

---

### 5 MINUTES BEFORE

Close all other applications except:
- Terminal 1 (n8n — do **NOT** close this)
- Terminal 2 (your presentation terminal)

Prevent popups:
- **Windows:** Action Center → Focus Assist → Alarms Only
- Quit Slack, Teams, Outlook, Chrome (unless needed for demo)
- Disable antivirus scan scheduling if possible

---

### WHEN YOU ARE READY TO START

In Terminal 2:

```
python orchestrator.py
```

Verify the checklist printed on screen:
- n8n is running (Terminal 1 is open)
- All 3 workflows are Active in n8n
- Volume is up
- Projector is connected

Then press **ENTER**.

- Presentation opens automatically in fullscreen
- Audio starts playing automatically
- Slides advance automatically
- Demos fire automatically at configured demo slides

**YOUR JOB DURING THE PRESENTATION:**
- Stand near the room, speak naturally alongside the audio if needed
- Watch Terminal 2 — it shows current slide number and demo countdown
- Do not touch keyboard or mouse unless something goes wrong

---

## IF SOMETHING GOES WRONG

| Problem | Fix |
|---------|-----|
| Audio stops / orchestrator crashes | `Ctrl+C` → take over manually with arrow keys |
| Slide didn't advance | Press `→` RIGHT ARROW once |
| Demo didn't fire | Open http://localhost:5678 → trigger manually from n8n UI |
| Presentation app not responding to keyboard | Click once on the window to give it focus → then `→` |

---

### AFTER THE PRESENTATION

```
Ctrl+C   ← Terminal 1 (stops n8n)
Ctrl+C   ← Terminal 2 (stops orchestrator if running)
```

Close the presentation app.

---

## ============================================================
## QUICK REFERENCE CARD
## ============================================================

```
BEFORE:
  Terminal 1:  npx n8n
  Browser:     http://localhost:5678  →  all 3 workflows ACTIVE
  Terminal 2:  python diagnose.py    (all ✅)

START:
  Terminal 2:  python orchestrator.py  →  press ENTER

DURING:
  SPACE  = pause/resume narration
  D      = skip demo countdown
  Q      = quit
  Watch Terminal 2. Do not touch anything else.

EMERGENCY:
  Ctrl+C → use arrow keys manually
```
