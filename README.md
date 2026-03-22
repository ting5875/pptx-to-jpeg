# PPTX Converter

A web app that converts PowerPoint files to JPEG images or PDF, intelligently expanding entrance animations into separate frames.

## What it does

Two conversion modes via a tabbed UI:

**JPEG Slides tab**
- Parses each slide's entrance animations and expands them into one image per click state
- A slide with 3 entrance animations produces 4 JPEGs: `slide_02_state_0.jpg` through `slide_02_state_3.jpg`
- Slides with no animations produce a single image: `slide_01.jpg`
- All images are bundled into a ZIP for download

**PDF tab**
- Same animation expansion, but outputs a single multi-page PDF
- Each animation state becomes a separate page

Both modes show a real-time progress bar during conversion (powered by Server-Sent Events).

Upload limit: 50 MB.

## Tech stack

| Layer | Tool |
|---|---|
| Web framework | FastAPI + Uvicorn |
| PPTX parsing | lxml (reads OOXML animation XML directly) |
| PPTX → PDF | LibreOffice (`soffice --headless`) |
| PDF → JPEG | Poppler (`pdftoppm`) |
| Frontend | Plain HTML + JS (no framework) |

## Local setup

**Prerequisites:** Docker Desktop

```bash
git clone https://github.com/ting5875/pptx-to-jpeg.git
cd pptx-to-jpeg
docker build -t pptx-converter .
docker run -p 8000:8000 pptx-converter
```

Open http://localhost:8000 in your browser.

> The first build takes 5–10 minutes — LibreOffice and the CJK/Liberation fonts are large (~700 MB image total).

## Project structure

```
app.py                  # FastAPI app — upload, SSE progress, download endpoints
pptx_to_jpeg.py         # Core logic — animation parsing & PPTX state expansion
templates/index.html    # Frontend UI
Dockerfile
requirements.txt
railway.toml            # Railway.app deployment config
```
