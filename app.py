import asyncio
import io
import json
import subprocess
import sys
import tempfile
import uuid
import zipfile
from pathlib import Path
from urllib.parse import quote

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import HTMLResponse, StreamingResponse

from pptx_to_jpeg import build_expanded_pptx, pptx_to_jpegs

MAX_UPLOAD_BYTES = 50 * 1024 * 1024  # 50 MB
TEMPLATES_DIR = Path(__file__).parent / "templates"

app = FastAPI()

# In-memory job store
jobs: dict = {}


@app.get("/", response_class=HTMLResponse)
async def index():
    return (TEMPLATES_DIR / "index.html").read_text(encoding="utf-8")


# ── Upload endpoints ──────────────────────────────────────────────────────────

def _create_job(filename: str, contents: bytes, job_type: str) -> str:
    job_id = str(uuid.uuid4())
    jobs[job_id] = {
        "type": job_type,
        "filename": filename,
        "data": contents,
        "status": "pending",
        "result_bytes": None,
        "error": None,
    }
    return job_id


async def _accept_upload(file: UploadFile) -> bytes:
    if not file.filename.lower().endswith(".pptx"):
        raise HTTPException(status_code=400, detail="File must be a .pptx file.")
    contents = await file.read()
    if len(contents) > MAX_UPLOAD_BYTES:
        raise HTTPException(status_code=413, detail="File exceeds 50 MB limit.")
    return contents


@app.post("/convert")
async def convert_jpeg(file: UploadFile = File(...)):
    contents = await _accept_upload(file)
    return {"job_id": _create_job(file.filename, contents, "jpeg")}


@app.post("/convert-pdf")
async def convert_pdf(file: UploadFile = File(...)):
    contents = await _accept_upload(file)
    return {"job_id": _create_job(file.filename, contents, "pdf")}


# ── Conversion workers ────────────────────────────────────────────────────────

def _run_jpeg_conversion(job_id: str):
    job = jobs[job_id]
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        input_path = tmp_path / "input.pptx"
        expanded_path = tmp_path / "expanded.pptx"
        jpegs_dir = tmp_path / "jpegs"
        jpegs_dir.mkdir()
        input_path.write_bytes(job["data"])

        job["status"] = "parsing"
        manifest = build_expanded_pptx(str(input_path), str(expanded_path))

        job["status"] = "converting"
        jpeg_paths = pptx_to_jpegs(str(expanded_path), str(jpegs_dir), dpi=150)

        job["status"] = "packaging"
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for i, (_, _, label) in enumerate(manifest):
                if i < len(jpeg_paths):
                    zf.write(jpeg_paths[i], arcname=f"{label}.jpg")
        buf.seek(0)
        job["result_bytes"] = buf.read()
        job["status"] = "done"


def _run_pdf_conversion(job_id: str):
    job = jobs[job_id]
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        input_path = tmp_path / "input.pptx"
        expanded_path = tmp_path / "expanded.pptx"
        input_path.write_bytes(job["data"])

        job["status"] = "parsing"
        build_expanded_pptx(str(input_path), str(expanded_path))

        job["status"] = "converting"
        soffice_script = Path("/mnt/skills/public/pptx/scripts/office/soffice.py")
        if soffice_script.exists():
            cmd = [sys.executable, str(soffice_script), "--headless",
                   "--convert-to", "pdf", str(expanded_path), "--outdir", str(tmp_path)]
        else:
            cmd = ["soffice", "--headless", "--convert-to", "pdf",
                   "--outdir", str(tmp_path), str(expanded_path)]

        result = subprocess.run(cmd, capture_output=True, text=True)
        pdf_files = list(tmp_path.glob("*.pdf"))
        if not pdf_files:
            raise RuntimeError(
                f"LibreOffice failed.\nSTDOUT: {result.stdout}\nSTDERR: {result.stderr}"
            )

        job["result_bytes"] = pdf_files[0].read_bytes()
        job["status"] = "done"


# ── SSE progress stream ───────────────────────────────────────────────────────

@app.get("/stream/{job_id}")
async def stream_progress(job_id: str):
    if job_id not in jobs:
        raise HTTPException(status_code=404, detail="Job not found.")

    job_type = jobs[job_id]["type"]

    async def generate():
        def event(step: str, pct: int, **kwargs) -> str:
            return f"data: {json.dumps({'step': step, 'pct': pct, **kwargs})}\n\n"

        loop = asyncio.get_running_loop()

        if job_type == "pdf":
            yield event("Parsing animations and building slide states…", 15)
            future = loop.run_in_executor(None, _run_pdf_conversion, job_id)
            prev_status = ""
            while not future.done():
                status = jobs.get(job_id, {}).get("status", "")
                if status != prev_status:
                    prev_status = status
                    if status == "converting":
                        yield event("Converting to PDF — this may take a minute…", 50)
                await asyncio.sleep(0.3)

        else:  # jpeg
            yield event("Parsing animations and building slide states…", 15)
            future = loop.run_in_executor(None, _run_jpeg_conversion, job_id)
            prev_status = ""
            while not future.done():
                status = jobs.get(job_id, {}).get("status", "")
                if status != prev_status:
                    prev_status = status
                    if status == "converting":
                        yield event("Converting to PDF — this may take a minute…", 40)
                    elif status == "packaging":
                        yield event("Rendering images and packaging ZIP…", 80)
                await asyncio.sleep(0.3)

        try:
            future.result()
        except Exception as exc:
            jobs[job_id]["status"] = "error"
            yield event("Conversion failed.", 0, error=str(exc))
            return

        yield event("Done! Download starting…", 100, done=True)

    return StreamingResponse(
        generate(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


# ── Download ──────────────────────────────────────────────────────────────────

@app.get("/download/{job_id}")
async def download(job_id: str):
    if job_id not in jobs or jobs[job_id]["status"] != "done":
        raise HTTPException(status_code=404, detail="Job not found or not ready.")

    job = jobs.pop(job_id)
    stem = Path(job["filename"]).stem

    if job["type"] == "pdf":
        filename = f"{stem}.pdf"
        media_type = "application/pdf"
    else:
        filename = f"{stem}_slides.zip"
        media_type = "application/zip"

    encoded_name = quote(filename)
    return StreamingResponse(
        io.BytesIO(job["result_bytes"]),
        media_type=media_type,
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{encoded_name}"},
    )
