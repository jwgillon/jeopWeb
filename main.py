import io
import logging
import os
import traceback
from typing import Any

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, Response
from fastapi.staticfiles import StaticFiles
from pptx import Presentation
import httpx
from pydantic import BaseModel

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

app = FastAPI(title="JeoparTy Generator API")

@app.get("/health")
def health():
    return {"status": "ok"}

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["POST", "GET"],
    allow_headers=["*"],
)

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(HERE, "template.pptm")
POINT_VALUES = [200, 400, 600, 800, 1000]
MASTER_API_KEY = os.environ.get("GEMINI_API_KEY", "")


class GenerateRequest(BaseModel):
    game_data: dict[str, Any]
    theme: str = "Game"
    api_key: str = ""


def set_shape_text(slide, shape_name: str, text: str) -> bool:
    for shape in slide.shapes:
        if shape.name != shape_name:
            continue
        if not shape.has_text_frame:
            log.warning("Shape %s has no text frame", shape_name)
            return False
        tf = shape.text_frame
        if not tf.paragraphs:
            log.warning("Shape %s has no paragraphs", shape_name)
            return False
        para = tf.paragraphs[0]
        if not para.runs:
            run = para.add_run()
            run.text = text
        else:
            para.runs[0].text = text
            for run in para.runs[1:]:
                run.text = ""
        return True
    log.warning("Shape not found: %s", shape_name)
    return False

@app.get("/has-master-key")
async def has_master_key():
    """Let the frontend know if a master key is available."""
    return {"available": bool(MASTER_API_KEY)}


class GeminiRequest(BaseModel):
    prompt: str


@app.post("/gemini")
async def call_gemini(req: GeminiRequest):
    """Call Gemini using the master key stored in environment."""
    if not MASTER_API_KEY:
        raise HTTPException(400, "No master API key configured on server")

    # await asyncio.sleep(6)  # stay under 5 RPM limit

    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={MASTER_API_KEY}"
    payload = {
        "contents": [{"parts": [{"text": req.prompt}]}],
        "generationConfig": {"temperature": 0.7, "maxOutputTokens": 8192}
    }

    async with httpx.AsyncClient(timeout=120) as client:
        res = await client.post(url, json=payload)
    if res.status_code != 200:
        raise HTTPException(502, f"Gemini API error: {res.text[:200]}")
    data = res.json()
    text = data.get("candidates", [{}])[0].get("content", {}).get("parts", [{}])[0].get("text", "")
    if not text:
        raise HTTPException(502, "Empty response from Gemini")
    log.info("Gemini via master key returned %d chars", len(text))
    return {"text": text}



@app.post("/generate")
async def generate(req: GenerateRequest):
    data = req.game_data

    for key in ["categories", "clues", "answers",
                "finalJeopardyClue", "finalJeopardyAnswer", "finalJeopardyTopic"]:
        if key not in data:
            raise HTTPException(400, f"Missing field: {key}")

    categories = data["categories"]
    clues      = data["clues"]
    answers    = data["answers"]

    log.info("Categories: %s", categories)

    if len(categories) != 5:
        raise HTTPException(400, f"Expected 5 categories, got {len(categories)}")
    if len(clues) != 5 or len(answers) != 5:
        raise HTTPException(400, f"Expected 5 difficulty rows")

    if not os.path.exists(TEMPLATE_PATH):
        raise HTTPException(500, f"Template not found: {TEMPLATE_PATH}")

    try:
        prs = Presentation(TEMPLATE_PATH)
        log.info("Template loaded. Slides: %d", len(prs.slides))

        # ── 1. Main panel category titles (Slide 1, index 0) ──
        panel_slide = prs.slides[0]
        for i, cat in enumerate(categories, start=1):
            ok = set_shape_text(panel_slide, f"Title_Cat{i}", cat)
            log.info("Title_Cat%d -> '%s' [%s]", i, cat, "OK" if ok else "MISS")

        # ── 2. Hidden data slides ──
        q_data_slide  = prs.slides[61]
        a_data_slide  = prs.slides[62]
        fj_data_slide = prs.slides[63]

        for i, cat in enumerate(categories, start=1):
            set_shape_text(q_data_slide, f"Data_Cat{i}", cat)

        # ── 3. Game question & answer slides directly ──
        for diff_idx, points in enumerate(POINT_VALUES):
            for cat_idx in range(5):
                cat_num = cat_idx + 1
                clue   = clues[diff_idx][cat_idx]   if diff_idx < len(clues)   and cat_idx < len(clues[diff_idx])   else ""
                answer = answers[diff_idx][cat_idx] if diff_idx < len(answers) and cat_idx < len(answers[diff_idx]) else ""

                q_shape_name = f"Q_Cat{cat_num}_{points}"
                a_shape_name = f"A_Cat{cat_num}_{points}"

                q_found = a_found = False
                for slide in prs.slides:
                    if not q_found:
                        q_found = set_shape_text(slide, q_shape_name, str(clue))
                    if not a_found:
                        a_found = set_shape_text(slide, a_shape_name, str(answer))
                    if q_found and a_found:
                        break

                set_shape_text(q_data_slide, f"Data_Q_Cat{cat_num}_{points}", str(clue))
                set_shape_text(a_data_slide, f"Data_A_Cat{cat_num}_{points}", str(answer))

                if not q_found or not a_found:
                    log.warning("MISS: %s=%s %s=%s", q_shape_name, q_found, a_shape_name, a_found)
                else:
                    log.info("✓ Cat%d %d: '%s' / '%s'", cat_num, points, str(clue)[:40], str(answer))

        # ── 4. Final Jeopardy slides directly ──
        fj_topic  = str(data["finalJeopardyTopic"])
        fj_clue   = str(data["finalJeopardyClue"])
        fj_answer = str(data["finalJeopardyAnswer"])

        for slide in prs.slides:
            set_shape_text(slide, "FinalJeopardyTopic", fj_topic)
            set_shape_text(slide, "FinalJeopardyClue",  fj_clue)
            set_shape_text(slide, "FinalJeopardyAnswer", fj_answer)

        set_shape_text(fj_data_slide, "Data_FJ_Topic",  fj_topic)
        set_shape_text(fj_data_slide, "Data_FJ_Clue",   fj_clue)
        set_shape_text(fj_data_slide, "Data_FJ_Answer",  fj_answer)

        log.info("Final Jeopardy: %s / %s", fj_topic, fj_answer)

        # ── 5. Build filename and return ──
        raw_theme = str(req.theme)
        slug = "".join(c for c in raw_theme[:20] if c.isalnum() or c in " -_").strip() or "Game"

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        file_bytes = buf.read()
        log.info("File built: %d bytes", len(file_bytes))

        return Response(
            content=file_bytes,
            media_type="application/vnd.ms-powerpoint.presentation.macroEnabled.12",
            headers={"Content-Disposition": f'attachment; filename="AI Jeopardy - {slug}.pptm"'},
        )

    except HTTPException:
        raise
    except Exception as e:
        log.error("Generation failed: %s", traceback.format_exc())
        raise HTTPException(500, f"Failed to generate file: {str(e)}")


# Serve static files
static_dir = os.path.join(HERE, "static")
if os.path.exists(static_dir):
    app.mount("/static", StaticFiles(directory=static_dir), name="static")


@app.get("/", response_class=HTMLResponse)
async def root():
    html_path = os.path.join(HERE, "index.html")
    if not os.path.exists(html_path):
        raise HTTPException(404, "index.html not found")
    with open(html_path) as f:
        return f.read()
