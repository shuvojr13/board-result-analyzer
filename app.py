"""
app.py - FastAPI Web Application for Bangladesh SSC/HSC Result Extraction
Handles file upload, Pandas data processing, SSE progress streaming, and Excel/CSV export.
"""

import asyncio
import io
import json
import logging
import os
import sys
import uuid
from datetime import datetime
from typing import Any, Dict, List, Optional

if sys.platform == "win32":
    try:
        asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
    except Exception:
        pass

import pandas as pd
from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from scraper import scrape_single_student, scrape_student_batch
from analysis import compute_result_analysis, generate_analysis_excel_bytes

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("app")

# Ensure necessary directories exist
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
STATIC_DIR = os.path.join(BASE_DIR, "static")
EXPORTS_DIR = os.path.join(BASE_DIR, "exports")
UPLOADS_DIR = os.path.join(BASE_DIR, "uploads")

for folder in [EXPORTS_DIR, UPLOADS_DIR, STATIC_DIR, TEMPLATES_DIR]:
    os.makedirs(folder, exist_ok=True)

app = FastAPI(
    title="Bangladesh Board Result Extractor",
    description="Automated SSC/HSC Data Extractor with Playwright & Pandas",
    version="1.0.0"
)

# Enable CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Mount static files
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")
templates = Jinja2Templates(directory=TEMPLATES_DIR)

# In-memory store for active jobs and progress queues
# job_id -> {"queue": asyncio.Queue, "records": List, "results": List, "status": str, "files": dict}
JOBS: Dict[str, Dict[str, Any]] = {}


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    """Renders the main Bootstrap user interface."""
    return templates.TemplateResponse(request=request, name="index.html")


@app.get("/sample-csv")
async def download_sample_csv():
    """Provides a sample CSV template for users to download."""
    sample_path = os.path.join(BASE_DIR, "sample_students.csv")
    if not os.path.exists(sample_path):
        sample_df = pd.DataFrame({
            "Roll": [100101, 100102, 100103, 100104, 100105],
            "Registration": [1512345678, 1512345679, 1512345680, 1512345681, 1512345682]
        })
        sample_df.to_csv(sample_path, index=False)
    return FileResponse(
        sample_path,
        media_type="text/csv",
        filename="sample_students.csv"
    )


@app.post("/api/start-extraction")
async def start_extraction(
    file: UploadFile = File(...),
    exam: str = Form(...),
    year: str = Form(...),
    board: str = Form(...),
    demo_mode: bool = Form(False)
):
    """
    Accepts student CSV/Excel file, parses Roll/Registration numbers using Pandas,
    and initializes an extraction job with an SSE event queue.
    """
    content = await file.read()
    filename = file.filename.lower()

    try:
        if filename.endswith(".csv"):
            df = pd.read_csv(io.BytesIO(content), dtype=str)
        elif filename.endswith((".xlsx", ".xls")):
            df = pd.read_excel(io.BytesIO(content), dtype=str)
        else:
            raise HTTPException(status_code=400, detail="Unsupported file format. Please upload a .csv or .xlsx file.")
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Failed to read file: {str(e)}")

    # Normalize column names (case-insensitive search for Roll and Registration)
    roll_col = next((c for c in df.columns if "roll" in str(c).lower()), None)
    reg_col = next((c for c in df.columns if any(k in str(c).lower() for k in ["reg", "registration"])), None)

    if not roll_col:
        raise HTTPException(
            status_code=400,
            detail="Could not find a 'Roll' column in your CSV/Excel file. Header must contain 'Roll'."
        )

    # Filter out empty rolls
    df = df.dropna(subset=[roll_col])
    records = []
    for _, row in df.iterrows():
        r = str(row[roll_col]).strip().split(".")[0]  # Strip any float decimals
        reg_val = str(row[reg_col]).strip().split(".")[0] if reg_col and pd.notna(row[reg_col]) else ""
        if r and r != "nan":
            records.append({"Roll": r, "Registration": reg_val if reg_val != "nan" else ""})

    if not records:
        raise HTTPException(status_code=400, detail="The uploaded file contains no valid Roll numbers.")

    job_id = str(uuid.uuid4())
    event_queue = asyncio.Queue()

    JOBS[job_id] = {
        "queue": event_queue,
        "records": records,
        "results": [],
        "exam": exam,
        "year": year,
        "board": board,
        "demo_mode": demo_mode,
        "status": "processing",
        "created_at": datetime.now(),
        "files": {}
    }

    # Launch background task for scraping
    asyncio.create_task(run_extraction_task(job_id))

    return {
        "job_id": job_id,
        "total_students": len(records),
        "exam": exam,
        "year": year,
        "board": board,
        "demo_mode": demo_mode
    }


@app.post("/api/check-individual")
async def check_individual(
    exam: str = Form(...),
    year: str = Form(...),
    board: str = Form(...),
    roll: str = Form(...),
    reg: str = Form(""),
    demo_mode: bool = Form(False)
):
    """
    Fetches the result for an individual student (Roll + Registration)
    via Playwright or simulated demo mode, returning full marksheet data.
    """
    roll = roll.strip()
    reg = reg.strip()
    if not roll:
        raise HTTPException(status_code=400, detail="Roll number is required.")

    try:
        result = await scrape_single_student(
            exam=exam,
            year=year,
            board=board,
            roll=roll,
            reg=reg,
            demo_mode=demo_mode
        )
        return {"success": True, "data": result}
    except Exception as e:
        logger.exception("Error during individual student check:")
        error_msg = str(e).strip() or "Connection to board server timed out or failed. Please enable 'Test / Demo Mode' to preview."
        raise HTTPException(status_code=500, detail=f"Failed to check result: {error_msg}")


async def run_extraction_task(job_id: str):
    """
    Background worker that runs the Playwright scraping loop,
    streams live progress through the job's queue,
    and generates the final Excel / CSV downloads with Pandas.
    """
    job = JOBS.get(job_id)
    if not job:
        return

    queue: asyncio.Queue = job["queue"]
    records = job["records"]
    exam = job["exam"]
    year = job["year"]
    board = job["board"]
    demo_mode = job["demo_mode"]

    async def on_progress_callback(current: int, total: int, student_result: Dict[str, Any]):
        job["results"].append(student_result)
        percent = int((current / total) * 100)
        await queue.put({
            "type": "progress",
            "current": current,
            "total": total,
            "percent": percent,
            "data": student_result
        })

    try:
        # Run batch scraping
        results = await scrape_student_batch(
            student_records=records,
            exam=exam,
            year=year,
            board=board,
            demo_mode=demo_mode,
            on_progress=on_progress_callback
        )

        # Build Pandas DataFrame from scraped results
        flattened_records = []
        all_subjects = set()

        for r in results:
            subj_grades = r.get("Subject_Grades", {})
            for s in subj_grades.keys():
                all_subjects.add(s)

        sorted_subjects = sorted(list(all_subjects))

        for r in results:
            row_dict = {
                "Roll": r.get("Roll"),
                "Registration": r.get("Registration"),
                "Name": r.get("Name"),
                "Father Name": r.get("Father Name", ""),
                "Mother Name": r.get("Mother Name", ""),
                "Institute": r.get("Institute", ""),
                "Board": r.get("Board"),
                "Group": r.get("Group", ""),
                "Type": r.get("Type", "REGULAR"),
                "Exam": r.get("Exam"),
                "Year": r.get("Year"),
                "Result Status": r.get("Result"),
                "GPA": r.get("GPA"),
                "Subject Grades Summary": r.get("Subject_Grades_Summary", ""),
            }

            # Add columns for individual subjects
            subj_grades = r.get("Subject_Grades", {})
            for subj in sorted_subjects:
                row_dict[subj] = subj_grades.get(subj, "-")

            flattened_records.append(row_dict)

        results_df = pd.DataFrame(flattened_records)

        # Generate exported files (Excel & CSV)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_filename = f"Board_Results_{exam}_{year}_{board}_{timestamp}.xlsx"
        csv_filename = f"Board_Results_{exam}_{year}_{board}_{timestamp}.csv"

        excel_filepath = os.path.join(EXPORTS_DIR, f"{job_id}_{excel_filename}")
        csv_filepath = os.path.join(EXPORTS_DIR, f"{job_id}_{csv_filename}")

        # Save CSV
        results_df.to_csv(csv_filepath, index=False, encoding="utf-8-sig")

        # Save formatted Excel with openpyxl
        with pd.ExcelWriter(excel_filepath, engine="openpyxl") as writer:
            results_df.to_excel(writer, index=False, sheet_name="Results")
            worksheet = writer.sheets["Results"]

            # Auto-fit columns
            for col in worksheet.columns:
                max_len = max(len(str(cell.value or "")) for cell in col)
                col_letter = col[0].column_letter
                worksheet.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 40)

        # Generate comprehensive result analysis and analysis Excel
        analysis_data = compute_result_analysis(results)
        job["analysis"] = analysis_data

        analysis_filename = f"Board_Results_Analysis_{exam}_{year}_{board}_{timestamp}.xlsx"
        analysis_filepath = os.path.join(EXPORTS_DIR, f"{job_id}_{analysis_filename}")

        meta = {"exam": exam, "year": year, "board": board}
        analysis_stream = generate_analysis_excel_bytes(analysis_data, meta)
        with open(analysis_filepath, "wb") as f_out:
            f_out.write(analysis_stream.getvalue())

        job["status"] = "completed"
        job["files"] = {
            "excel_path": excel_filepath,
            "excel_filename": excel_filename,
            "csv_path": csv_filepath,
            "csv_filename": csv_filename,
            "analysis_excel_path": analysis_filepath,
            "analysis_excel_filename": analysis_filename
        }

        # Send completion event
        await queue.put({
            "type": "completed",
            "total": len(results),
            "download_xlsx": f"/api/download/{job_id}/xlsx",
            "download_csv": f"/api/download/{job_id}/csv",
            "download_analysis_xlsx": f"/api/download-analysis/{job_id}",
            "analysis": analysis_data
        })

    except Exception as e:
        logger.exception("Error during extraction task execution:")
        job["status"] = "failed"
        await queue.put({
            "type": "error",
            "message": f"Extraction encountered an error: {str(e)}"
        })


@app.get("/api/progress-stream/{job_id}")
async def progress_stream(job_id: str):
    """
    Server-Sent Events (SSE) endpoint to stream live extraction progress to the frontend UI.
    """
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job ID not found.")

    async def event_generator():
        queue: asyncio.Queue = job["queue"]
        while True:
            try:
                # Wait for next event with a timeout
                event_data = await asyncio.wait_for(queue.get(), timeout=60.0)
                yield f"data: {json.dumps(event_data)}\n\n"

                # Stop streaming once job completes or fails
                if event_data.get("type") in ["completed", "error"]:
                    break
            except asyncio.TimeoutError:
                # Keep-alive heartbeat comment
                yield ": keep-alive\n\n"
            except asyncio.CancelledError:
                break

    return StreamingResponse(
        event_generator(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no"
        }
    )


@app.get("/api/download/{job_id}/{file_format}")
async def download_result_file(job_id: str, file_format: str):
    """
    Downloads the resulting Excel or CSV file.
    """
    job = JOBS.get(job_id)
    if not job or "files" not in job:
        raise HTTPException(status_code=404, detail="Result file not found or extraction not completed.")

    files = job["files"]
    if file_format.lower() == "xlsx":
        path = files.get("excel_path")
        filename = files.get("excel_filename")
        media = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    elif file_format.lower() == "csv":
        path = files.get("csv_path")
        filename = files.get("csv_filename")
        media = "text/csv"
    else:
        raise HTTPException(status_code=400, detail="Invalid format requested. Use 'xlsx' or 'csv'.")

    if not path or not os.path.exists(path):
        raise HTTPException(status_code=404, detail="Requested file was not found on server.")

    return FileResponse(
        path=path,
        filename=filename,
        media_type=media
    )


@app.get("/api/analysis/{job_id}")
async def get_job_analysis(job_id: str):
    """
    Returns the comprehensive analysis for a completed extraction job.
    """
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job ID not found.")

    if "analysis" in job and job["analysis"]:
        return {"success": True, "analysis": job["analysis"]}

    results = job.get("results", [])
    if not results:
        raise HTTPException(status_code=400, detail="No extracted results found for this job.")

    analysis_data = compute_result_analysis(results)
    job["analysis"] = analysis_data
    return {"success": True, "analysis": analysis_data}


@app.get("/api/download-analysis/{job_id}")
async def download_analysis_file(job_id: str):
    """
    Downloads the comprehensive Analysis Excel workbook (.xlsx).
    """
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job ID not found.")

    files = job.get("files", {})
    path = files.get("analysis_excel_path")
    filename = files.get("analysis_excel_filename")

    if not path or not os.path.exists(path):
        results = job.get("results", [])
        if not results:
            raise HTTPException(status_code=404, detail="No student records available for analysis.")
        analysis_data = job.get("analysis") or compute_result_analysis(results)
        meta = {"exam": job.get("exam", ""), "year": job.get("year", ""), "board": job.get("board", "")}
        analysis_stream = generate_analysis_excel_bytes(analysis_data, meta)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"Board_Results_Analysis_{timestamp}.xlsx"
        path = os.path.join(EXPORTS_DIR, f"{job_id}_{filename}")
        with open(path, "wb") as f_out:
            f_out.write(analysis_stream.getvalue())
        job.setdefault("files", {})["analysis_excel_path"] = path
        job["files"]["analysis_excel_filename"] = filename

    return FileResponse(
        path=path,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.post("/api/analyze-file")
async def analyze_uploaded_file(file: UploadFile = File(...)):
    """
    Parses an uploaded CSV or Excel result file, computes analysis,
    and returns full analysis data JSON.
    """
    content = await file.read()
    filename = file.filename.lower()
    try:
        if filename.endswith(".csv"):
            df = pd.read_csv(io.BytesIO(content), dtype=str)
        elif filename.endswith((".xlsx", ".xls")):
            df = pd.read_excel(io.BytesIO(content), dtype=str)
        else:
            raise HTTPException(status_code=400, detail="Please upload a .csv or .xlsx file.")
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Failed to read file: {str(e)}")

    records = df.to_dict("records")
    if not records:
        raise HTTPException(status_code=400, detail="The uploaded file contains no student rows.")

    analysis_data = compute_result_analysis(records)
    return {
        "success": True,
        "filename": file.filename,
        "total_records": len(records),
        "analysis": analysis_data
    }


@app.post("/api/export-analysis-file")
async def export_analysis_from_uploaded_file(file: UploadFile = File(...)):
    """
    Accepts an uploaded CSV/Excel results file and directly streams
    the comprehensive Analysis Excel workbook.
    """
    content = await file.read()
    filename = file.filename.lower()
    try:
        if filename.endswith(".csv"):
            df = pd.read_csv(io.BytesIO(content), dtype=str)
        elif filename.endswith((".xlsx", ".xls")):
            df = pd.read_excel(io.BytesIO(content), dtype=str)
        else:
            raise HTTPException(status_code=400, detail="Please upload a .csv or .xlsx file.")
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Failed to read file: {str(e)}")

    records = df.to_dict("records")
    if not records:
        raise HTTPException(status_code=400, detail="The uploaded file contains no student rows.")

    analysis_data = compute_result_analysis(records)
    meta = {"exam": "Board Result", "year": "", "board": "All"}
    excel_stream = generate_analysis_excel_bytes(analysis_data, meta)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    export_name = f"Board_Result_Analysis_{timestamp}.xlsx"
    return StreamingResponse(
        excel_stream,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{export_name}"'}
    )


if __name__ == "__main__":
    import uvicorn
    print("\n" + "="*60)
    print(" Bangladesh SSC/HSC Board Result Analyzer & Data Extractor")
    print(" Server starting on: http://localhost:8000")
    print("="*60 + "\n")
    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=True)
