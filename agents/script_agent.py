"""
script_agent.py — Generates per-slide speech scripts via Claude.

Configure via .env:
  ANTHROPIC_API_KEY   your Anthropic key
  PRESENTER_NAME      e.g. "Alex Johnson"
  PRESENTER_ROLE      e.g. "digital health and AI expert"
  ORGANIZATION        e.g. "GlobalHealth NGO"
  AUDIENCE_CONTEXT    e.g. "Programme managers and M&E staff"
  DOMAIN_CONTEXT      e.g. "malaria programmes, maternal health, field teams in Kenya"
  TOTAL_SLIDES        number of slides in your deck (default: 17)
  DEMO_SLIDES         comma-separated slide:type pairs, e.g. "8:email,10:meeting,12:research"
"""

import os
from dotenv import load_dotenv
from anthropic import Anthropic
from core.logger import get_logger

load_dotenv()
log = get_logger("script_agent")

client = Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))

PRESENTER_NAME   = os.getenv("PRESENTER_NAME",   "the presenter")
PRESENTER_ROLE   = os.getenv("PRESENTER_ROLE",   "digital health and AI expert")
ORGANIZATION     = os.getenv("ORGANIZATION",     "your organization")
AUDIENCE_CONTEXT = os.getenv("AUDIENCE_CONTEXT", "Programme managers, M&E staff, country office teams")
DOMAIN_CONTEXT   = os.getenv("DOMAIN_CONTEXT",   "public health programmes, field teams, donor reporting")
TOTAL_SLIDES     = int(os.getenv("TOTAL_SLIDES",  "17"))

# Parse DEMO_SLIDES from env: "8:email,10:meeting,12:research"
_demo_raw = os.getenv("DEMO_SLIDES", "8:email,10:meeting,12:research")
DEMO_SLIDES: dict[int, str] = {}
for pair in _demo_raw.split(","):
    pair = pair.strip()
    if ":" in pair:
        num, dtype = pair.split(":", 1)
        try:
            DEMO_SLIDES[int(num.strip())] = dtype.strip()
        except ValueError:
            pass

BLACK_SLIDES: set[int] = set()

# Default per-slide tone/length config for a 17-slide deck.
# Adjust to match your own deck structure.
SLIDE_CONTEXT: dict[int, dict] = {
    1:  {"tone": "energetic, warm opening — set up the room, acknowledge the audience",              "words": (100, 140)},
    2:  {"tone": "grounded, relatable — use an analogy to make AI feel familiar not scary",          "words": (140, 175)},
    3:  {"tone": "clear contrast, build intrigue — the shift from tool to infrastructure",           "words": (150, 185)},
    4:  {"tone": "practical and exciting — show breadth, connect to their actual work",              "words": (140, 175)},
    5:  {"tone": "informative, empowering — they can start today, most tools are free",             "words": (120, 155)},
    6:  {"tone": "direct, action-oriented — role-specific value, tomorrow not someday",             "words": (130, 165)},
    7:  {"tone": "building tension — make the problem feel real and personal",                       "words": (120, 155)},
    8:  {"tone": "calm, observational — short, build anticipation before email demo",               "words": ( 70,  95)},
    9:  {"tone": "observational, narrate the email pipeline running on screen",                     "words": ( 60,  80)},
    10: {"tone": "calm, observational — short setup before meeting demo",                           "words": ( 70,  95)},
    11: {"tone": "observational — describe the four meeting outputs appearing",                     "words": ( 60,  80)},
    12: {"tone": "calm but building — contrast with previous two, agentic AI is different",         "words": ( 80, 110)},
    13: {"tone": "curious, observational — let the agentic reasoning unfold on screen",            "words": ( 70,  90)},
    14: {"tone": "serious, grounded, human — this is the most important slide",                     "words": (150, 185)},
    15: {"tone": "reveal moment — proud, warm, let the architecture land",                          "words": (155, 195)},
    16: {"tone": "practical, encouraging — give them a clear first step this week",                 "words": (130, 165)},
    17: {"tone": "quiet, human, closing — leave them with one lasting thought before Q&A",          "words": ( 70, 100)},
}

SYSTEM_PROMPT = f"""You are {PRESENTER_NAME}, a {PRESENTER_ROLE} presenting to colleagues
at a capacity-building AI workshop.

Your audience: {AUDIENCE_CONTEXT}.
They are smart, mission-driven, and overworked. They care deeply about outcomes.
Many have heard about AI but haven't seen what it can actually do for their specific work.

Your voice: calm, confident, warm. You speak like a trusted colleague genuinely excited
about what you've built — not a salesperson, not a professor.

Make the scripts conversational and engaging:
- Open slides with a scene, a question, or a vivid example
- Connect AI capabilities to the domain: {DOMAIN_CONTEXT}
- Use short, punchy sentences next to longer explanatory ones — create rhythm
- Give the audience the "aha" moment: the moment they realise this is possible for THEM
- Occasionally use second-person ("you", "your team") to pull them in

Never say:
- "This slide shows..."
- "As you can see..."
- "In conclusion..."
- "Let me now..."
- "Moving on to..."
- "Next I'll talk about..."

Just speak naturally, directly, as if the audience is in the room with you."""


def generate_speech_for_slide(slide_number: int, slide_content: str,
                               custom_instruction: str = "") -> str:
    """Generate speech for a single slide."""
    log.info(f"Generating script for slide {slide_number}")

    if slide_number in BLACK_SLIDES:
        return ""

    ctx = SLIDE_CONTEXT.get(slide_number, {"tone": "clear and direct", "words": (130, 165)})
    min_words, max_words = ctx["words"]

    if max_words == 0:
        return ""

    tone = ctx["tone"]
    is_demo = slide_number in DEMO_SLIDES
    demo_type = DEMO_SLIDES.get(slide_number, "")

    demo_instruction = ""
    if is_demo:
        demo_instructions = {
            "email": (
                "You are about to trigger the email workflow live on screen. "
                "Tell the audience to watch carefully — an inbound email is about to arrive "
                "and you are not going to touch anything. One sentence of context about the problem, "
                "then build the tension. End by directing their eyes to the screen. Keep it tight."
            ),
            "meeting": (
                "You are about to trigger the meeting pipeline live. "
                "Give them one line about the everyday meeting pain — the black hole of no follow-through. "
                "Then: one transcript in, four structured outputs out. You are not touching anything. "
                "Short. Crisp. Direct their eyes to the screen."
            ),
            "research": (
                "This is the most impressive demo — agentic AI, fundamentally different from the previous two. "
                "The previous workflows followed fixed paths — event in, output out. "
                "This one thinks. It plans its own research strategy, evaluates the quality of what it finds, "
                "and decides whether to broaden, narrow, or arbitrate. "
                "Then it writes a calibrated brief. No human in the loop. "
                "Build the distinction clearly — then hand over to the demo."
            ),
        }
        demo_instruction = demo_instructions.get(demo_type, "")

    custom_block = f"\nSPECIAL INSTRUCTION: {custom_instruction}\n" if custom_instruction else ""

    prompt = f"""SLIDE {slide_number} of {TOTAL_SLIDES}.
Tone: {tone}
Target length: {min_words}–{max_words} words (strict — stay within range)
{custom_block}
Slide content:
---
{slide_content.strip() if slide_content.strip() else "(No text — this slide is mostly visual)"}
---
{"DEMO CONTEXT: " + demo_instruction if demo_instruction else ""}

STYLE REMINDERS:
- Open with something that grabs attention (a question, a scene, a surprising fact)
- Use domain-specific context where natural: {DOMAIN_CONTEXT}
- Vary sentence rhythm — short punchy lines, then a fuller explanation
- Second-person "you" at least once to pull the audience in

Write exactly what {PRESENTER_NAME} says when this slide is on screen.
Output ONLY the speech text — no labels, no stage directions, no quotes, no markdown."""

    try:
        response = client.messages.create(
            model="claude-opus-4-5",
            max_tokens=600,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": prompt}]
        )
        speech = response.content[0].text.strip()
        words = len(speech.split())
        log.info(f"Slide {slide_number}: generated {words} words")

        if words < min_words * 0.7:
            log.warning(f"Slide {slide_number}: {words} words — too short, retrying")
            print(f"  ⚠️ Slide {slide_number}: too short ({words} words), retrying...")
            return generate_speech_for_slide(slide_number, slide_content, custom_instruction)

        return speech

    except Exception as e:
        log.error(f"Slide {slide_number}: Claude API error — {e}", exc_info=True)
        raise


if __name__ == "__main__":
    test_content = "You already use AI as a tool. Today you will see AI as infrastructure."
    speech = generate_speech_for_slide(1, test_content)
    print(f"\nTest speech ({len(speech.split())} words):\n{speech}")
