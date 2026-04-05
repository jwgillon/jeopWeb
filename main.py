from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from typing import Any
import copy
import json
import os
import tempfile
import traceback

from pptx import Presentation
from pptx.util import Pt

app = FastAPI(title="JeoparTy Generator API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["POST", "GET"],
    allow_headers=["*"],
)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "template.pptm")

POINT_VALUES = [200, 400, 600, 800, 1000]


class GenerateRequest(BaseModel):
    game_data: dict[str, Any]


def set_shape_text(slide, shape_name: str, text: str):
    """Find a shape by name and replace its text, preserving formatting."""
    for shape in slide.shapes:
        if shape.name == shape_name and shape.has_text_frame:
            tf = shape.text_frame
            # Clear all paragraphs except the first, then set text
            # Preserve the run formatting from the first run of the first paragraph
            if tf.paragraphs and tf.paragraphs[0].runs:
                first_run = tf.paragraphs[0].runs[0]
                # Save font properties
                font = first_run.font
                bold = font.bold
                size = font.size
                color = font.color.rgb if font.color and font.color.type else None
                name = font.name
            else:
                bold = size = color = name = None

            # Clear existing text
            for para in tf.paragraphs:
                for run in para.runs:
                    run.text = ""

            # Set new text on first run of first paragraph
            if tf.paragraphs and tf.paragraphs[0].runs:
                run = tf.paragraphs[0].runs[0]
                run.text = text
                if bold is not None:
                    run.font.bold = bold
                if size is not None:
                    run.font.size = size
                if color is not None:
                    run.font.color.rgb = color
                if name is not None:
                    run.font.name = name
            elif tf.paragraphs:
                # No runs — add one
                para = tf.paragraphs[0]
                run = para.add_run()
                run.text = text
            return True
    return False


@app.post("/generate")
async def generate(req: GenerateRequest):
    data = req.game_data

    # Validate required keys
    required = ["categories", "clues", "answers",
                "finalJeopardyClue", "finalJeopardyAnswer", "finalJeopardyTopic"]
    for key in required:
        if key not in data:
            raise HTTPException(400, f"Missing field: {key}")

    categories = data["categories"]
    clues = data["clues"]       # clues[difficulty_idx][cat_idx]
    answers = data["answers"]   # answers[difficulty_idx][cat_idx]

    if len(categories) != 5:
        raise HTTPException(400, "Expected exactly 5 categories")
    if len(clues) != 5 or len(answers) != 5:
        raise HTTPException(400, "Expected 5 difficulty rows")

    try:
        prs = Presentation(TEMPLATE_PATH)

        # Slide 60 = index 59 (questions data slide)
        q_slide = prs.slides[59]
        # Slide 61 = index 60 (answers data slide)
        a_slide = prs.slides[60]
        # Slide 62 = index 61 (final jeopardy data slide)
        fj_slide = prs.slides[61]

        # --- Inject categories ---
        for cat_idx, cat_name in enumerate(categories, start=1):
            set_shape_text(q_slide, f"Data_Cat{cat_idx}", cat_name)

        # --- Inject clues & answers ---
        for diff_idx, points in enumerate(POINT_VALUES):
            for cat_idx in range(5):
                cat_num = cat_idx + 1
                clue = clues[diff_idx][cat_idx] if diff_idx < len(clues) and cat_idx < len(clues[diff_idx]) else ""
                answer = answers[diff_idx][cat_idx] if diff_idx < len(answers) and cat_idx < len(answers[diff_idx]) else ""

                q_name = f"Data_Q_Cat{cat_num}_{points}"
                a_name = f"Data_A_Cat{cat_num}_{points}"

                set_shape_text(q_slide, q_name, str(clue))
                set_shape_text(a_slide, a_name, str(answer))

        # --- Inject Final Jeopardy ---
        set_shape_text(fj_slide, "Data_FJ_Topic", str(data["finalJeopardyTopic"]))
        set_shape_text(fj_slide, "Data_FJ_Clue", str(data["finalJeopardyClue"]))
        set_shape_text(fj_slide, "Data_FJ_Answer", str(data["finalJeopardyAnswer"]))

        # Save to temp file
        with tempfile.NamedTemporaryFile(suffix=".pptm", delete=False) as tmp:
            tmp_path = tmp.name

        prs.save(tmp_path)

        return FileResponse(
            tmp_path,
            media_type="application/vnd.ms-powerpoint.presentation.macroEnabled.12",
            filename="JeoparTy_Game.pptm",
            background=None,
        )

    except HTTPException:
        raise
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(500, f"Failed to generate file: {str(e)}")


# Serve the frontend
if os.path.exists(os.path.join(os.path.dirname(__file__), "index.html")):
    @app.get("/", response_class=HTMLResponse)
    async def root():
        with open(os.path.join(os.path.dirname(__file__), "index.html")) as f:
            return f.read()
