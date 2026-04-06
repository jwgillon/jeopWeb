import logging
import os
import tempfile
import traceback
from typing import Any

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, Response
from fastapi.staticfiles import StaticFiles
from pptx import Presentation
from pydantic import BaseModel

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

app = FastAPI(title="JeoparTy Generator API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["POST", "GET"],
    allow_headers=["*"],
)

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(HERE, "template.pptm")
POINT_VALUES = [200, 400, 600, 800, 1000]


class GenerateRequest(BaseModel):
    game_data: dict[str, Any]
    theme: str = "Game"


def set_shape_text(slide, shape_name: str, text: str) -> bool:
    for shape in slide.shapes:
        if shape.name != shape_name:
            continue
        if not shape.has_text_frame:
            log.warning("Shape %s found but has no text frame", shape_name)
            return False
        tf = shape.text_frame
        if not tf.paragraphs:
            log.warning("Shape %s has no paragraphs", shape_name)
            return False
        para = tf.paragraphs[0]
        if not para.runs:
            # No runs — create one
            run = para.add_run()
            run.text = text
        else:
            # Set first run, clear the rest
            para.runs[0].text = text
            for run in para.runs[1:]:
                run.text = ""
        return True
    log.warning("Shape not found: %s", shape_name)
    return False


@app.post("/generate")
async def generate(req: GenerateRequest):
    data = req.game_data
    log.info("Generate request received. Keys: %s", list(data.keys()))

    # Validate
    for key in ["categories", "clues", "answers",
                "finalJeopardyClue", "finalJeopardyAnswer", "finalJeopardyTopic"]:
        if key not in data:
            raise HTTPException(400, f"Missing field: {key}")

    categories = data["categories"]
    clues = data["clues"]
    answers = data["answers"]

    log.info("Categories: %s", categories)
    log.info("Clues rows: %d, Answers rows: %d", len(clues), len(answers))

    if len(categories) != 5:
        raise HTTPException(400, f"Expected 5 categories, got {len(categories)}")
    if len(clues) != 5:
        raise HTTPException(400, f"Expected 5 clue rows, got {len(clues)}")
    if len(answers) != 5:
        raise HTTPException(400, f"Expected 5 answer rows, got {len(answers)}")

    if not os.path.exists(TEMPLATE_PATH):
        log.error("Template not found at %s", TEMPLATE_PATH)
        raise HTTPException(500, f"Template file not found at {TEMPLATE_PATH}")

    log.info("Loading template from %s", TEMPLATE_PATH)

    try:
        prs = Presentation(TEMPLATE_PATH)
        log.info("Template loaded. Slide count: %d", len(prs.slides))

        q_slide  = prs.slides[59]   # Data slide - questions
        a_slide  = prs.slides[60]   # Data slide - answers
        fj_slide = prs.slides[61]   # Data slide - final jeopardy

        # Inject categories
        for i, cat in enumerate(categories, start=1):
            ok = set_shape_text(q_slide, f"Data_Cat{i}", cat)
            log.info("Data_Cat%d -> '%s' [%s]", i, cat, "OK" if ok else "MISS")

        # Inject clues and answers
        for diff_idx, points in enumerate(POINT_VALUES):
            for cat_idx in range(5):
                cat_num = cat_idx + 1
                clue   = clues[diff_idx][cat_idx]   if diff_idx < len(clues)   and cat_idx < len(clues[diff_idx])   else ""
                answer = answers[diff_idx][cat_idx] if diff_idx < len(answers) and cat_idx < len(answers[diff_idx]) else ""
                ok_q = set_shape_text(q_slide, f"Data_Q_Cat{cat_num}_{points}", str(clue))
                ok_a = set_shape_text(a_slide, f"Data_A_Cat{cat_num}_{points}", str(answer))
                if not ok_q or not ok_a:
                    log.warning("Miss at diff=%d pts=%d cat=%d  q=%s a=%s", diff_idx, points, cat_num, ok_q, ok_a)

        # Inject Final Jeopardy
        set_shape_text(fj_slide, "Data_FJ_Topic",  str(data["finalJeopardyTopic"]))
        set_shape_text(fj_slide, "Data_FJ_Clue",   str(data["finalJeopardyClue"]))
        set_shape_text(fj_slide, "Data_FJ_Answer",  str(data["finalJeopardyAnswer"]))
        log.info("Final Jeopardy injected: %s", data["finalJeopardyTopic"])

        # Build filename slug from theme (first 15 chars, sanitized)
        raw_theme = str(req.theme)
        slug = "".join(c for c in raw_theme[:15] if c.isalnum() or c in " -_").strip()
        if not slug:
            slug = "Game"

        # Save to bytes and return — avoids temp file race condition
        import io
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        file_bytes = buf.read()
        log.info("File built successfully, size: %d bytes", len(file_bytes))

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


# Serve static files (images etc.) if folder exists
static_dir = os.path.join(HERE, "static")
if os.path.exists(static_dir):
    app.mount("/static", StaticFiles(directory=static_dir), name="static")

# Serve the frontend
@app.get("/", response_class=HTMLResponse)
async def root():
    html_path = os.path.join(HERE, "index.html")
    if not os.path.exists(html_path):
        raise HTTPException(404, "index.html not found")
    with open(html_path) as f:
        return f.read()
