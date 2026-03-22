import asyncio
import io
import json
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

# In-memory job store: job_id -> {filename, data, status, zip_bytes, error}
jobs: dict = {}


@app.get("/", response_class=HTMLResponse)
async def index():
    return (TEMPLATES_DIR / "index.html").read_text(encoding="utf-8")


@app.post("/convert")
async def convert_start(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".pptx"):
        raise HTTPException(status_code=400, detail="File must be a .pptx file.")
    contents = await file.read()
    if len(contents) > MAX_UPLOAD_BYTES:
        raise HTTPException(status_code=413, detail="File exceeds 50 MB limit.")

    job_id = str(uuid.uuid4())
    jobs[job_id] = {
        "filename": file.filename,
        "data": contents,
        "status": "pending",
        "zip_bytes": None,
        "error": None,
    }
    return {"job_id": job_id}


def _run_conversion(job_id: str):
    """Blocking conversion — called from a thread pool executor."""
    job = jobs[job_id]
    contents = job["data"]

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        input_path = tmp_path / "input.pptx"
        expanded_path = tmp_path / "expanded.pptx"
        jpegs_dir = tmp_path / "jpegs"
        jpegs_dir.mkdir()
        input_path.write_bytes(contents)

        job["status"] = "parsing"
        manifest = build_expanded_pptx(str(input_path), str(expanded_path))

        job["status"] = "converting"
        jpeg_paths = pptx_to_jpegs(str(expanded_path), str(jpegs_dir), dpi=150)

        job["status"] = "packaging"
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for i, (_, _, label) in enumerate(manifest):
                if i < len(jpeg_paths):
                    zf.write(jpeg_paths[i], arcname=f"{label}.jpg")
        zip_buffer.seek(0)
        job["zip_bytes"] = zip_buffer.read()
        job["status"] = "done"


@app.get("/stream/{job_id}")
async def stream_progress(job_id: str):
    if job_id not in jobs:
        raise HTTPException(status_code=404, detail="Job not found.")

    async def generate():
        def event(step: str, pct: int, **kwargs) -> str:
            return f"data: {json.dumps({'step': step, 'pct': pct, **kwargs})}\n\n"

        yield event("Parsing animations and building slide states…", 15)

        loop = asyncio.get_running_loop()
        future = loop.run_in_executor(None, _run_conversion, job_id)

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
            jobs[job_id]["error"] = str(exc)
            yield event("Conversion failed.", 0, error=str(exc))
            return

        yield event("Done! Download starting…", 100, done=True)

    return StreamingResponse(
        generate(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.get("/download/{job_id}")
async def download(job_id: str):
    if job_id not in jobs or jobs[job_id]["status"] != "done":
        raise HTTPException(status_code=404, detail="Job not found or not ready.")

    job = jobs.pop(job_id)
    stem = Path(job["filename"]).stem
    zip_name = f"{stem}_slides.zip"
    encoded_name = quote(zip_name)

    return StreamingResponse(
        io.BytesIO(job["zip_bytes"]),
        media_type="application/zip",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{encoded_name}"},
    )
