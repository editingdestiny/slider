from fastapi import FastAPI, Request
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
import uvicorn
import json
import tempfile
import os
import logging
import shutil
import time
import httpx
from pathlib import Path
from dotenv import load_dotenv
from pptx import Presentation
import io
from general_presentation import create_general_presentation  # Main generator with charts

def is_valid_pptx(file_content: bytes) -> bool:
    """Validate if the byte content is a valid PPTX file."""
    if not file_content:
        logger.error("Validation failed: File content is empty.")
        return False
    try:
        # Use BytesIO to treat the byte content as a file
        file_stream = io.BytesIO(file_content)
        # Try to open the presentation. If it fails, it's corrupted.
        Presentation(file_stream)
        logger.info("PPTX validation successful.")
        return True
    except Exception as e:
        logger.error(f"Validation failed: The file is not a valid PPTX package or is corrupted: {e}")
        return False

def create_presentation_with_real_charts(data, output_path):
    """
    Create presentation using general_presentation for real chart graphics
    """
    try:
        logger.info("Using general_presentation for real chart graphics")
        
        # Extract search phrase for context
        search_phrase = "Business Analysis"
        if isinstance(data, dict):
            search_phrase = data.get('search_phrase', 'Business Analysis')
            if 'slides' in data and data['slides']:
                # Try to extract topic from first slide title
                first_slide = data['slides'][0]
                if first_slide.get('title'):
                    search_phrase = first_slide['title']
        
        # Use general_presentation which has real chart capabilities
        presentation = create_general_presentation(data, search_phrase)
        if presentation:
            presentation.save(output_path)
            logger.info(f"Successfully created presentation with real charts: {output_path}")
            return True
        else:
            logger.error("general_presentation failed to create presentation")
            return False
            
    except Exception as e:
        logger.error(f"Error in chart presentation creation: {str(e)}")
        return False

# Load environment variables
load_dotenv()

# Configuration
DOWNLOAD_BASE_URL = os.getenv("DOWNLOAD_BASE_URL", "https://slider.sd-ai.co.uk")
N8N_WEBHOOK_URL = os.getenv("N8N_WEBHOOK_URL", "https://sd-n8n.duckdns.org/webhook/slider")  # Production default

# Pydantic models for request validation
class CustomizationOptions(BaseModel):
    slide_bg_color: str = "#0F1632"
    title_font_color: str = "#FFFFFF"
    title_bg_color: str = "#44546A"
    body_text_color: str = "#FFFFFF"
    title_position: str = "left"
    font_size: int = 16

class SlideGenerationRequest(BaseModel):
    search_phrase: str
    number_of_slides: int = 5  # Default to 5 slides
    customization: CustomizationOptions = None

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

# Request deduplication tracking
recent_requests = {}
REQUEST_COOLDOWN = 3  # seconds

@app.post("/generate-slides-from-search")
async def generate_slides_from_search(request: SlideGenerationRequest):
    """Generate PowerPoint slides by triggering n8n webhook and return the file"""
    global recent_requests
    current_time = time.time()
    request_id = f"{request.search_phrase}_{int(current_time)}"
    
    # Deduplication check with more precise timing
    request_key = f"{request.search_phrase}_{request.number_of_slides}"
    
    if request_key in recent_requests:
        time_diff = current_time - recent_requests[request_key]
        if time_diff < REQUEST_COOLDOWN:
            logger.warning(f"[{request_id}] DUPLICATE REQUEST BLOCKED: Same request made {time_diff:.2f}s ago")
            return {"error": f"Duplicate request detected. Please wait {REQUEST_COOLDOWN - time_diff:.1f} seconds before trying again.", "status": "rate_limited"}
    
    recent_requests[request_key] = current_time
    
    # Clean up old entries (older than 10 seconds)
    cutoff_time = current_time - 10
    recent_requests = {k: v for k, v in recent_requests.items() if v > cutoff_time}
    
    try:
        logger.info(f"[{request_id}] NEW REQUEST: Triggering n8n webhook for: {request.search_phrase}, {request.number_of_slides} slides")
        
        # Prepare payload for n8n webhook
        webhook_payload = {
            "search_phrase": request.search_phrase,
            "number_of_slides": request.number_of_slides,
            "customization": request.customization.dict() if request.customization else None,
            "timestamp": current_time
        }
        
        logger.info(f"[{request_id}] Sending to n8n: {webhook_payload}")
        
        # Trigger n8n webhook and expect binary file response
        async with httpx.AsyncClient(timeout=None) as client:
            try:
                response = await client.post(N8N_WEBHOOK_URL, json=webhook_payload)

                # --- TEMPORARY LOGGING START ---
                # logger.info(f"N8N_RESPONSE_STATUS: {response.status_code}")
                # logger.info(f"N8N_RESPONSE_HEADERS: {response.headers}")
                # logger.info(f"N8N_RESPONSE_BODY (first 200 bytes): {response.content[:200]}")
                # --- TEMPORARY LOGGING END ---
                
                response.raise_for_status()
                
                logger.info(f"[{request_id}] n8n response received, status: {response.status_code}")
                
                # Check if response is binary (PowerPoint file)
                content_type = response.headers.get('content-type', '')
                if 'application/vnd.openxmlformats-officedocument.presentationml.presentation' in content_type:
                    # Return the PowerPoint file directly
                    logger.info(f"[{request_id}] Received PowerPoint file from n8n webhook, returning file")
                    
                    # --- VALIDATION STEP ---
                    if not is_valid_pptx(response.content):
                        logger.error(f"[{request_id}] Validation failed: Received corrupted file from n8n.")
                        return {
                            "error": "Received a corrupted presentation file from the generation service. Please try again.",
                            "status": "error"
                        }
                    # --- END VALIDATION ---

                    filename = f"{request.search_phrase.replace(' ', '_')}_Analysis.pptx"
                    
                    # Save temporarily to return as FileResponse
                    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
                        tmp.write(response.content)
                        
                        return FileResponse(
                            path=tmp.name,
                            filename=filename,
                            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
                else:
                    # Handle JSON response (fallback)
                    webhook_result = response.json()
                    logger.info(f"n8n webhook returned JSON: {webhook_result}")
                    
                    return {
                        "status": "success",
                        "message": f"Request submitted to n8n for processing: {request.search_phrase}",
                        "webhook_response": webhook_result,
                        "search_phrase": request.search_phrase,
                        "number_of_slides": request.number_of_slides
                    }
                    
            except httpx.HTTPError as e:
                logger.error(f"Error calling n8n webhook: {e}")
                return {"error": f"Failed to trigger n8n webhook: {str(e)}", "status": "error"}

    except Exception as e:
        logger.error(f"Error in slide generation: {str(e)}")
        return {"error": f"Failed to generate presentation: {str(e)}", "status": "error"}

@app.post("/create-presentation")
async def create_presentation(content_data: dict):
    """Direct endpoint for n8n to create general business presentations - returns file directly"""
    try:
        search_phrase = content_data.get("search_phrase", "Analysis")
        logger.info(f"Creating general presentation for: {search_phrase}")
        
        # Handle data field if nested
        data = content_data.get("data", content_data)
        customization = content_data.get("customization")

        # If data is a string, parse it as JSON
        if isinstance(data, str):
            try:
                data = json.loads(data)
            except json.JSONDecodeError:
                logger.error("Failed to parse 'data' string as JSON.")
                return {"error": "Invalid format for 'data' field.", "status": "error"}
        
        # Check if data has slides directly or if we need to convert from old format
        if "slides" not in data:
            logger.warning(f"No 'slides' key found in data. Available keys: {list(data.keys())}")
            
            # Try to convert old ESG format to general slides
            if "executiveSummary" in data:
                logger.info("Converting old ESG format to general slides")
                slides = []
                
                # Executive Summary slide
                exec_summary = data.get("executiveSummary", {})
                if isinstance(exec_summary, dict):
                    key_finding = exec_summary.get('keyFinding', 'Key business findings and insights')
                else:
                    key_finding = str(exec_summary) if exec_summary else 'Key business findings and insights'
                
                slides.append({
                    "title": f"Executive Summary: {search_phrase}",
                    "headline": "Key Business Overview",
                    "content": f"• {key_finding}\n• Market opportunities and strategic implications\n• Risk assessment and mitigation strategies\n• Recommended next steps for implementation"
                })
                
                # Impact Analysis slide
                if "impactAnalysis" in data:
                    impact = data["impactAnalysis"]
                    if isinstance(impact, dict):
                        financial = impact.get('financial', 'Positive ROI expected')
                    else:
                        financial = 'Positive ROI expected'
                    
                    slides.append({
                        "title": "Impact Analysis",
                        "headline": "Business Impact Assessment",
                        "content": f"• Financial impact: {financial}\n• Operational efficiency improvements\n• Strategic positioning advantages\n• Long-term business sustainability"
                    })
                
                # Regional/Market Data slide
                if "regionalData" in data and data["regionalData"]:
                    regional_data = data["regionalData"]
                    if isinstance(regional_data, list) and regional_data:
                        regional = regional_data[0] if isinstance(regional_data[0], dict) else {}
                    elif isinstance(regional_data, dict):
                        regional = regional_data
                    else:
                        regional = {}
                    
                    region = regional.get('region', 'Global market') if isinstance(regional, dict) else 'Global market'
                    trend = regional.get('trend', 'Positive growth trajectory') if isinstance(regional, dict) else 'Positive growth trajectory'
                    
                    slides.append({
                        "title": "Market Analysis",
                        "headline": "Regional and Market Insights",
                        "content": f"• Region: {region}\n• Growth trends: {trend}\n• Market drivers and opportunities\n• Competitive landscape assessment"
                    })
                
                data = {"slides": slides}
            else:
                # Create a meaningful slide from available data
                available_keys = [k for k in data.keys() if k not in ['search_phrase', 'number_of_slides', 'timestamp']]
                content_points = []
                
                for key in available_keys[:4]:  # Take up to 4 keys
                    value = data.get(key, "")
                    if isinstance(value, (str, int, float)) and str(value).strip():
                        content_points.append(f"• {key.replace('_', ' ').title()}: {str(value)[:100]}")
                    elif isinstance(value, dict) and value:
                        content_points.append(f"• {key.replace('_', ' ').title()}: Analysis available")
                    elif isinstance(value, list) and value:
                        content_points.append(f"• {key.replace('_', ' ').title()}: {len(value)} items identified")
                
                if not content_points:
                    content_points = [
                        f"• Comprehensive analysis of {search_phrase}",
                        "• Strategic business opportunities identified",
                        "• Risk assessment and mitigation strategies",
                        "• Implementation roadmap and recommendations"
                    ]
                
                data = {
                    "slides": [
                        {
                            "title": f"Business Analysis: {search_phrase}",
                            "headline": "Comprehensive Business Intelligence",
                            "content": "\n".join(content_points)
                        }
                    ]
                }
        
        # Always create general slides presentation using rich visuals
        logger.info("Creating general slides presentation with charts and tables")
        
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
            # Try to use the new rich presentation generator first
            presentation = create_general_presentation(data, search_phrase, customization)
            if presentation:
                presentation.save(tmp.name)
                
                # --- VALIDATION STEP ---
                with open(tmp.name, "rb") as f:
                    file_content = f.read()
                if not is_valid_pptx(file_content):
                    logger.error("Validation failed: Generated a corrupted file locally.")
                    return {
                        "error": "The server generated a corrupted presentation file. Please check the logs.",
                        "status": "error"
                    }
                # --- END VALIDATION ---
                
                success = True
            else:
                logger.error("Failed to generate presentation with general_presentation")
                return {"error": "Failed to create PowerPoint presentation", "status": "error"}
            
            if not success:
                return {"error": "Failed to create PowerPoint presentation", "status": "error"}
            
            # Return file directly (hardcoded for n8n compatibility)
            filename = f"{search_phrase.replace(' ', '_')}_Presentation.pptx"
            logger.info(f"Returning presentation file directly: {filename}")
            return FileResponse(
                path=tmp.name,
                filename=filename,
                media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
            
    except Exception as e:
        logger.error(f"Error creating presentation: {str(e)}")
        return {"error": f"Failed to create presentation: {str(e)}", "status": "error"}

@app.get("/health")
async def health_check():
    """Health check endpoint"""
    return {"status": "healthy", "service": "pptx-generator"}

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8010)
