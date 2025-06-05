from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pptx import Presentation
import shutil
import os
import language_tool_python
import uuid

app = FastAPI()

# CORS for local dev / Netlify frontend
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Change this to your frontend domain in prod
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

UPLOAD_DIR = "uploads"
OUTPUT_DIR = "output"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

tool = language_tool_python.LanguageTool('en-US')

@app.post("/api/process-ppt")
async def process_ppt(
    file: UploadFile = File(...),
    report: str = Form(...),
    options: str = Form(...)
):
    try:
        # Save uploaded file temporarily
        temp_filename = f"{uuid.uuid4()}_{file.filename}"
        temp_path = os.path.join(UPLOAD_DIR, temp_filename)

        with open(temp_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        prs = Presentation(temp_path)
        amended_count = 0

        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    original_text = shape.text
                    corrected_text = tool.correct(original_text)
                    if corrected_text != original_text:
                        shape.text = corrected_text
                        amended_count += 1

        # Save corrected file
        corrected_filename = f"corrected_{file.filename}"
        corrected_path = os.path.join(OUTPUT_DIR, corrected_filename)
        prs.save(corrected_path)

        # Return JSON with count and file path
        return JSONResponse(
            content={
                "amendedSlidesCount": amended_count,
                "fileUrl": f"/api/download/{corrected_filename}"
            }
        )

    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})


@app.get("/api/download/{filename}")
def download_file(filename: str):
    path = os.path.join(OUTPUT_DIR, filename)
    if os.path.exists(path):
        return FileResponse(path, media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation', filename=filename)
    return JSONResponse(status_code=404, content={"error": "File not found"})
