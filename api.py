"""
Vendor Form Filler API
=======================
A simple API that wraps form_filler_engine.py.

POST /fill  →  upload a form file  →  get filled file back

That's it. n8n calls this endpoint with the file, gets the result.
"""

import os
import tempfile
import shutil
import httpx
from pathlib import Path
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse

from form_filler_engine import fill_form

app = FastAPI(title="Vendor Form Filler API")

# Google Doc ID — the long string in your doc URL:
# https://docs.google.com/document/d/THIS_PART/edit
GOOGLE_DOC_ID = os.getenv("GOOGLE_DOC_ID", "")

# Fallback to local file if no Google Doc configured
LOCAL_MASTER_PATH = os.getenv("MASTER_DATA_PATH", "/app/master_data.md")


def fetch_master_data() -> str:
    """Fetch latest master data from Google Doc, or fall back to local file."""
    if GOOGLE_DOC_ID:
        url = f"https://docs.google.com/document/d/{GOOGLE_DOC_ID}/export?format=txt"
        try:
            resp = httpx.get(url, timeout=10, follow_redirects=True)
            resp.raise_for_status()
            return resp.text
        except Exception as e:
            print(f"⚠️ Failed to fetch Google Doc: {e}. Falling back to local file.")

    if os.path.exists(LOCAL_MASTER_PATH):
        return Path(LOCAL_MASTER_PATH).read_text(encoding="utf-8")

    raise RuntimeError("No master data available. Set GOOGLE_DOC_ID or provide local file.")


@app.get("/health")
def health():
    """Health check for Railway / uptime monitors."""
    return {"status": "ok"}


@app.get("/master")
def preview_master():
    """Preview the current master data (for debugging)."""
    try:
        text = fetch_master_data()
        source = "google_doc" if GOOGLE_DOC_ID else "local_file"
        return {"source": source, "data": text[:500] + "..." if len(text) > 500 else text}
    except Exception as e:
        return {"error": str(e)}


@app.post("/fill")
async def fill(file: UploadFile = File(...)):
    """
    Upload a blank vendor form (xlsx, docx, or pdf).
    Returns the filled form in the same format.
    """
    # Validate extension
    ext = Path(file.filename).suffix.lower()
    if ext not in (".xlsx", ".docx", ".pdf"):
        raise HTTPException(400, f"Unsupported format: {ext}. Send .xlsx, .docx, or .pdf")

    # Check master data exists
    try:
        master_text = fetch_master_data()
    except RuntimeError as e:
        raise HTTPException(500, str(e))

    # Save uploaded file to temp dir
    tmp_dir = tempfile.mkdtemp()
    try:
        input_path = os.path.join(tmp_dir, file.filename)
        with open(input_path, "wb") as f:
            content = await file.read()
            f.write(content)

        # Write master data to temp file (engine reads from file path)
        master_path = os.path.join(tmp_dir, "master_data.md")
        with open(master_path, "w", encoding="utf-8") as f:
            f.write(master_text)

        # Fill the form
        output_filename = f"FILLED_{file.filename}"
        output_path = os.path.join(tmp_dir, output_filename)

        result = fill_form(master_path, input_path, output_path)

        # Determine response MIME type
        mime_map = {
            ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            ".pdf": "application/pdf",
            ".txt": "text/plain",
        }

        # For non-fillable PDFs, the engine creates a _fill_guide.txt
        actual_output = output_path
        if not os.path.exists(output_path):
            # Check for fill guide
            guide_path = output_path.replace(".pdf", "_fill_guide.txt")
            if os.path.exists(guide_path):
                actual_output = guide_path
            else:
                raise HTTPException(500, "Form filling failed — no output generated")

        actual_ext = Path(actual_output).suffix.lower()
        mime = mime_map.get(actual_ext, "application/octet-stream")

        return FileResponse(
            actual_output,
            media_type=mime,
            filename=Path(actual_output).name,
        )

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500, f"Error filling form: {str(e)}")
    finally:
        # Cleanup happens after response is sent (FastAPI handles this)
        # For safety, we don't delete tmp_dir here since FileResponse needs it
        pass


if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)