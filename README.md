# AI Presentation Orchestrator

Fully automated AI presentation system with live avatar narration, n8n workflow demos,
and voice-driven Q&A. Built with HeyGen, Claude, Edge TTS, and n8n.

## The Two-Phase Design

```
PHASE 1: Pre-generation  (run 15–20 min before showtime)
  → Reads all slides from PPTX or Google Slides
  → Generates speech scripts via Claude
  → Submits ALL to HeyGen simultaneously (or synthesises via Edge TTS)
  → Downloads all audio/video to cache/
  → Creates cache/manifest.json

PHASE 2: Live presentation  (run when ready to present)
  → Loads cached audio/video (no waiting)
  → For each slide: plays narration → triggers demo (if needed) → advances slide
  → Everything timed from actual audio/video duration — no guessing
```

---

## Folder Structure

```
AI_PRESENTATION/
├── .env.example                   ← template — copy and fill in
├── credentials.example.json       ← Google OAuth template
├── presentation.pptx              ← your slide deck (never commit)
├── pre_generate.py                ← run FIRST
├── orchestrator.py                ← run to present
├── script_agent.py
├── voice_engine.py
├── heygen_engine.py
├── slide_controller.py
├── slide_reader.py
├── google_slides_reader.py
├── n8n_trigger.py
├── slack_notifier.py
├── diagnose.py
├── logger.py
├── workflow_monitor.html
├── cache/                         ← auto-created by pre_generate.py
│   ├── manifest.json
│   ├── scripts.json
│   ├── slide_01.mp3               ← or .mp4 if using HeyGen avatar
│   └── ...
└── demo_data/
    ├── meeting_transcript.txt     ← sample transcript
    └── research_question.txt
```

---

## Setup

### 1. Clone and install dependencies

```bash
git clone <your-repo-url>
cd AI_PRESENTATION
pip install -r requirements.txt
```

Or install manually:

```bash
pip install anthropic python-pptx requests pyautogui python-dotenv pygame edge-tts pynput
```

**Also install ffmpeg** (needed for audio duration detection):

- **Windows:** Download from https://www.gyan.dev/ffmpeg/builds/ → unzip → add `bin\` to PATH
- **Mac:** `brew install ffmpeg`
- **Linux:** `sudo apt install ffmpeg`

---

### 2. Configure your `.env`

```bash
cp .env.example .env
```

Fill in `.env` with your actual keys. See `.env.example` for all required variables.

---

### 3. Find your HeyGen IDs (one-time, if using HeyGen)

```bash
python heygen_engine.py
```

This prints your available avatars and voices. Copy the IDs into `.env`.

---

### 4. Google Slides setup (optional — if not using a local PPTX)

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Enable the **Google Slides API**
3. Create a **Service Account** → download the JSON key → save as `credentials.json`
4. Share your presentation with the service account email (Viewer access)
5. Set `GOOGLE_SLIDES_PRESENTATION_ID` and `GOOGLE_SLIDES_CREDENTIALS_FILE` in `.env`

---

## Pre-Generate Audio/Video

Run **15–20 minutes before** your presentation:

```bash
python pre_generate.py
```

This will:
- Read all slides
- Call Claude to write narration scripts
- Synthesise audio (Edge TTS) or generate avatar videos (HeyGen)
- Save everything to `cache/`

When complete you will see: `COMPLETE — N/N files — X.X min total`

If it crashes partway through, just re-run — cached slides are skipped automatically.

To regenerate everything from scratch:

```bash
python pre_generate.py --force
```

---

## Verify Everything

```bash
python diagnose.py
```

Every line should say ✅. Fix anything marked ❌ before presenting.

---

## Run the Presentation

### Step 1 — Start n8n (keep this terminal open)

```bash
npx n8n
```

Wait for `n8n ready on port 5678`, then open http://localhost:5678 and make sure
all 3 workflows (Email, Meeting, Research) show a **green Active** toggle.

### Step 2 — Start the orchestrator

```bash
python orchestrator.py
```

Press **ENTER** when prompted. The orchestrator will:
1. Open your presentation in fullscreen slideshow mode
2. Play narration audio for each slide
3. Automatically advance slides when audio finishes
4. Fire n8n demo triggers at configured demo slides
5. Launch voice Q&A on the final slide

### Controls During Presentation

| Key | Action |
|-----|--------|
| `SPACE` | Pause / resume narration |
| `D` | Skip demo countdown |
| `Q` | Quit |

---

## Demo Slides

Configure in `.env` via `DEMO_SLIDES=8:email,10:meeting,12:research` (slide:type pairs).

| Type | What it triggers |
|------|-----------------|
| `email` | Email pipeline n8n workflow |
| `meeting` | Meeting transcript → structured outputs |
| `research` | Agentic research engine |

---

## Troubleshooting

**Narration doesn't play**
→ Run `python diagnose.py` → check audio section
→ Ensure ffmpeg/pygame is installed

**n8n webhook fails**
→ Confirm n8n is running at http://localhost:5678
→ Check all 3 workflows are Active (green toggle)
→ Trigger manually from the n8n UI as fallback

**Slide didn't advance**
→ Press `→` (RIGHT ARROW) manually
→ Ensure pyautogui is installed: `pip install pyautogui`

**Want to restart from a specific slide**
→ `Ctrl+C` to stop → manually go to that slide in the presentation app → re-run orchestrator

---

## Workflow Monitor

Open `workflow_monitor.html` in any browser to see a live visual of the 3 demo pipelines.
No server required — it's a static HTML file.

---

## After the Presentation

```bash
# Stop n8n
Ctrl+C   (Terminal 1)

# Stop orchestrator (if still running)
Ctrl+C   (Terminal 2)
```

Logs are saved to `logs/presentation_YYYYMMDD.log`.
