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
from pptx_generator import create_pptx_from_json
from general_presentation import create_general_presentation

# Load environment variables
load_dotenv()

# Configuration
DOWNLOAD_BASE_URL = os.getenv("DOWNLOAD_BASE_URL", "https://slider.sd-ai.co.uk")
N8N_WEBHOOK_URL = "https://sd-n8n.duckdns.org/webhook/slider"

# Pydantic models for request validation
class SlideGenerationRequest(BaseModel):
    search_phrase: str
    number_of_slides: int = 5  # Default to 5 slides

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

@app.post("/generate-slides-from-search")
async def generate_slides_from_search(request: SlideGenerationRequest):
    """Generate PowerPoint slides by triggering n8n webhook and return the file"""
    try:
        logger.info(f"Triggering n8n webhook for: {request.search_phrase}, {request.number_of_slides} slides")
        
        # Prepare payload for n8n webhook
        webhook_payload = {
            "search_phrase": request.search_phrase,
            "number_of_slides": request.number_of_slides,
            "timestamp": time.time()
        }
        
        # Trigger n8n webhook and expect binary file response
        async with httpx.AsyncClient(timeout=60.0) as client:
            try:
                response = await client.post(N8N_WEBHOOK_URL, json=webhook_payload)
                response.raise_for_status()
                
                # Check if response is binary (PowerPoint file)
                content_type = response.headers.get('content-type', '')
                if 'application/vnd.openxmlformats-officedocument.presentationml.presentation' in content_type:
                    # Return the PowerPoint file directly
                    logger.info(f"Received PowerPoint file from n8n webhook, returning file")
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
    """Direct endpoint for n8n to create general business presentations"""
    try:
        search_phrase = content_data.get("search_phrase", "Analysis")
        return_file = content_data.get("return_file", False)
        logger.info(f"Creating general presentation for: {search_phrase}, return_file: {return_file}")
        
        # Handle data field if nested
        data = content_data.get("data", content_data)
        
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
            presentation = create_general_presentation(data, search_phrase)
            if presentation:
                presentation.save(tmp.name)
                success = True
            else:
                # Fallback to basic generator if rich generator fails
                logger.warning("Rich presentation failed, falling back to basic generator")
                success = create_pptx_from_json(data, tmp.name)
            
            if not success:
                return {"error": "Failed to create PowerPoint presentation", "status": "error"}
            
            # Return file directly if requested
            if return_file:
                filename = f"{search_phrase.replace(' ', '_')}_Presentation.pptx"
                logger.info(f"Returning presentation file directly: {filename}")
                return FileResponse(
                    path=tmp.name,
                    filename=filename,
                    media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            
            # Store file persistently for download URL
            persistent_path = f"/tmp/slides_pptx_{tmp.name.split('/')[-1]}"
            shutil.move(tmp.name, persistent_path)
            
            download_url = f"{DOWNLOAD_BASE_URL}/download/{persistent_path.split('/')[-1]}"
            logger.info(f"Generated presentation: {download_url}")
            
            return {
                "download_url": download_url,
                "filename": f"{search_phrase.replace(' ', '_')}_Presentation.pptx",
                "status": "success",
                "presentation_type": "General",
                "slides_generated": len(data.get("slides", [])),
                "search_phrase": search_phrase
            }
            
    except Exception as e:
        logger.error(f"Error creating presentation: {str(e)}")
        return {"error": f"Failed to create presentation: {str(e)}", "status": "error"}

@app.post("/process-ai-content")
async def process_ai_content(ai_data: dict):
    """Process AI-generated content and create general business presentation"""
    try:
        logger.info(f"Processing AI-generated content: {type(ai_data)}")
        
        # Always process as general slides with rich visuals
        logger.info("Processing general slides content with rich presentation generator")
        
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
            # Try to use the new rich presentation generator first
            presentation = create_general_presentation(ai_data, ai_data.get('search_phrase', 'Analysis'))
            if presentation:
                presentation.save(tmp.name)
                success = True
            else:
                # Fallback to basic generator if rich generator fails
                logger.warning("Rich presentation failed, falling back to basic generator")
                success = create_pptx_from_json(ai_data, tmp.name)
            
            if not success:
                return {"error": "Failed to create PowerPoint presentation", "status": "error"}
            
            # Store file persistently
            persistent_path = f"/tmp/slides_pptx_{tmp.name.split('/')[-1]}"
            shutil.move(tmp.name, persistent_path)
            
            download_url = f"{DOWNLOAD_BASE_URL}/download/{persistent_path.split('/')[-1]}"
            logger.info(f"Generated general presentation: {download_url}")
            
            return {
                "download_url": download_url,
                "filename": f"Presentation.pptx",
                "status": "success",
                "presentation_type": "General",
                "slides_generated": len(ai_data.get("slides", []))
            }
            
    except Exception as e:
        logger.error(f"Error processing AI content: {str(e)}")
        return {"error": f"Failed to process AI content: {str(e)}", "status": "error"}

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
