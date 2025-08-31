from fastapi import FastAPI, Request
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
import uvicorn
import json
import tempfile
import os
import logging
import shutil
import time
from pptx_generator import create_pptx_from_json

# Configuration
DOWNLOAD_BASE_URL = os.getenv("DOWNLOAD_BASE_URL", "http://localhost:8010")

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(title="PowerPoint Slide Generator")

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Mount static files
app.mount("/static", StaticFiles(directory="static"), name="static")

def cleanup_old_files():
    """Clean up old PPTX files from /tmp directory"""
    try:
        current_time = time.time()
        for filename in os.listdir("/tmp"):
            if filename.startswith("pptx_") or filename.endswith(".pptx"):
                file_path = f"/tmp/{filename}"
                # Remove files older than 1 hour
                if current_time - os.path.getctime(file_path) > 3600:
                    os.remove(file_path)
                    logger.info(f"Cleaned up old file: {filename}")
    except Exception as e:
        logger.error(f"Error during cleanup: {str(e)}")

@app.get("/", response_class=HTMLResponse)
async def root():
    """Serve the main form page"""
    try:
        with open("templates/form.html", "r") as f:
            return f.read()
    except FileNotFoundError:
        return HTMLResponse(content="<h1>Template not found</h1>", status_code=500)

@app.post("/ai-generate-pptx")
async def ai_generate_pptx(request: Request):
    """Generate PowerPoint from AI-enhanced JSON data (for n8n workflow)"""
    try:
        # Accept raw JSON from n8n workflow
        data = await request.json()
        logger.info(f"Received AI-enhanced data: {json.dumps(data, indent=2)}")

        # Handle different data formats from N8N
        slides_data = None
        
        if isinstance(data, list):
            # If it's an array, take the first element
            if len(data) > 0:
                first_item = data[0]
                if "slides" in first_item:
                    slides_data = first_item["slides"]
                else:
                    # Treat the array items as individual slides
                    slides_data = data
            else:
                return {"error": "Empty array received", "status": "error"}
        elif isinstance(data, dict):
            if "slides" in data:
                # Standard format: {"slides": [...]}
                slides_data = data["slides"]
            else:
                # Single slide object: {"title": "...", "headline": "...", "content": "..."}
                slides_data = [data]
        else:
            return {"error": "Invalid data format. Expected object or array", "status": "error"}

        if not slides_data:
            return {"error": "No slides data found", "status": "error"}

        # Create the expected format for the generator
        slides_json = {"slides": slides_data}

        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
            success = create_pptx_from_json(slides_json, tmp.name)
            if not success:
                return {"error": "Failed to create PowerPoint presentation", "status": "error"}

            # Store file in a more persistent location
            import shutil
            persistent_path = f"/tmp/pptx_{tmp.name.split('/')[-1]}"
            shutil.move(tmp.name, persistent_path)

            # Generate download URL using configured base URL
            download_url = f"{DOWNLOAD_BASE_URL}/download/{persistent_path.split('/')[-1]}"
            logger.info(f"Generated download URL: {download_url}")

            return {
                "download_url": download_url,
                "filename": "slides.pptx",
                "status": "success",
                "slides_processed": len(slides_data)
            }

    except Exception as e:
        logger.error(f"Error in AI PPTX generation: {str(e)}")
        return {"error": f"Failed to generate presentation: {str(e)}", "status": "error"}

@app.post("/generate-pptx")
async def generate_pptx(request: Request):
    """Generate PowerPoint from JSON data (for web form)"""
    try:
        # Accept JSON from web form
        data = await request.json()
        logger.info(f"Received data: {json.dumps(data, indent=2)}")

        if not data or "slides" not in data:
            return {"error": "Invalid data format. Expected {'slides': [...]}", "status": "error"}

        slides_json = data  # Data is already parsed JSON

        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
            success = create_pptx_from_json(slides_json, tmp.name)
            if not success:
                return {"error": "Failed to create PowerPoint presentation", "status": "error"}

            # Store file in a more persistent location
            import shutil
            persistent_path = f"/tmp/pptx_{tmp.name.split('/')[-1]}"
            shutil.move(tmp.name, persistent_path)

            # Generate download URL using configured base URL
            download_url = f"{DOWNLOAD_BASE_URL}/download/{persistent_path.split('/')[-1]}"
            logger.info(f"Generated download URL: {download_url}")

            return {
                "download_url": download_url,
                "filename": "slides.pptx",
                "status": "success",
                "slides_processed": len(slides_json.get("slides", []))
            }

    except Exception as e:
        logger.error(f"Error in PPTX generation: {str(e)}")
        return {"error": f"Failed to generate presentation: {str(e)}", "status": "error"}

@app.get("/download/{filename}")
async def download_file(filename: str):
    """Download the generated PowerPoint file"""
    try:
        # Handle both old and new file naming schemes
        if filename.startswith("pptx_"):
            file_path = f"/tmp/{filename}"
        else:
            # Legacy support for old naming
            file_path = f"/tmp/{filename}"

        logger.info(f"Attempting to download file: {file_path}")

        if not os.path.exists(file_path):
            logger.error(f"File not found: {file_path}")
            return {"error": "File not found", "status": "error"}

        # Check file size to ensure it's not empty
        file_size = os.path.getsize(file_path)
        logger.info(f"File size: {file_size} bytes")

        if file_size == 0:
            logger.error(f"File is empty: {file_path}")
            return {"error": "Generated file is empty", "status": "error"}

        return FileResponse(
            file_path,
            filename="slides.pptx",
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

    except Exception as e:
        logger.error(f"Error downloading file: {str(e)}")
        return {"error": f"Failed to download file: {str(e)}", "status": "error"}

@app.on_event("startup")
async def startup_event():
    """Clean up old files on startup"""
    cleanup_old_files()
    logger.info("Application started and old files cleaned up")

@app.post("/test-professional-slide")
async def test_professional_slide():
    """Generate a test slide with professional styling"""
    test_data = {
        "slides": [
            {
                "title": "Professional Business Presentation",
                "headline": "Executive Summary",
                "content": "• Strategic business objectives achieved\n• Revenue growth of 25% YoY\n• Market leadership maintained\n• Innovation pipeline strengthened",
                "chartType": "bar",
                "chartData": {
                    "labels": ["Q1", "Q2", "Q3", "Q4"],
                    "values": [120, 150, 180, 200],
                    "colors": ["#60A5FA", "#3B82F6", "#1D4ED8", "#1E3A8A"],
                    "title": "Quarterly Revenue Growth"
                }
            }
        ]
    }

    try:
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
            success = create_pptx_from_json(test_data, tmp.name)
            if not success:
                return {"error": "Failed to create test presentation", "status": "error"}

            # Store file in a more persistent location
            persistent_path = f"/tmp/test_pptx_{tmp.name.split('/')[-1]}"
            shutil.move(tmp.name, persistent_path)

            download_url = f"{DOWNLOAD_BASE_URL}/download/{persistent_path.split('/')[-1]}"

            return {
                "download_url": download_url,
                "filename": "professional_test.pptx",
                "status": "success",
                "message": "Professional slide generated with dark blue background and white text"
            }

    except Exception as e:
        logger.error(f"Error generating test slide: {str(e)}")
        return {"error": f"Failed to generate test: {str(e)}", "status": "error"}

@app.get("/cleanup")
async def manual_cleanup():
    """Manually trigger cleanup of old files"""
    try:
        cleanup_old_files()
        return {"status": "success", "message": "Old files cleaned up"}
    except Exception as e:
        logger.error(f"Error during manual cleanup: {str(e)}")
        return {"status": "error", "message": str(e)}

@app.get("/health")
async def health_check():
    """Health check endpoint"""
    return {"status": "healthy", "service": "pptx-generator"}

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8010)
