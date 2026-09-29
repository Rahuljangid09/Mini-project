from fastapi import FastAPI, UploadFile, File, HTTPException, Request
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from typing import List

import uuid
import os
import shutil

from app.core.format_detector import FORMAT_2019, detect_format, detect_format
from app.core.nep_parser import parse_ledger_nep
from app.core.pdf_parser import extract_text_from_pdf, parse_ledger
from app.core.excel_writer import generate_excel
from app.core.analysis_builder import build_analysis_data
from app.state import app_state

app = FastAPI()
app.mount("/static", StaticFiles(directory="app/static"), name="static")

UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

templates = Jinja2Templates(directory="app/templates")


@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    return templates.TemplateResponse(request, "index.html", {})


@app.post("/upload-pdf")
async def upload_pdf(request: Request, files: List[UploadFile] = File(...)):
    pdf_files = [f for f in files if f.filename.lower().endswith(".pdf")]

    if not pdf_files:
        raise HTTPException(status_code=400, detail="No PDF files were uploaded.")

    all_students = []
    failed_files = []

    for file in pdf_files:
        unique_name = f"{uuid.uuid4()}_{file.filename}"
        pdf_path = os.path.join(UPLOAD_DIR, unique_name)

        try:
            with open(pdf_path, "wb") as buffer:
                shutil.copyfileobj(file.file, buffer)

            fmt = detect_format(pdf_path)
            if fmt == FORMAT_2019:
                raw_text = extract_text_from_pdf(pdf_path)
                students = parse_ledger(raw_text)
            else:  # FORMAT_2024_NEP
                students = parse_ledger_nep(pdf_path)

            if not students:
                failed_files.append(file.filename)
                continue

            all_students.extend(students)

        except Exception:
            # One bad/corrupt PDF shouldn't take down the whole batch.
            failed_files.append(file.filename)
            continue

    if not all_students:
        raise HTTPException(
            status_code=422,
            detail=f"Could not parse any student records from the uploaded PDF(s): {failed_files}",
        )

    excel_path = generate_excel(all_students)
    app_state.set_analysis_data(build_analysis_data(all_students))

    filename = os.path.basename(excel_path)

    return templates.TemplateResponse(
    request, "index.html",
    {"download_link": f"/download/{filename}", "failed_files": failed_files},
)


@app.get("/dashboard", response_class=HTMLResponse)
def dashboard(request: Request):
    return templates.TemplateResponse(
    request, "dashboard.html",
    {"analysis_data": app_state.get_analysis_data()},
)


@app.get("/download/{filename}")
def download_file(filename: str):
    file_path = os.path.join("outputs", filename)
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="File not found.")
    return FileResponse(
        path=file_path,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Cache-Control": "no-store, no-cache, must-revalidate"},
    )
