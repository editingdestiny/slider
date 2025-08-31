from fastapi import FastAPI, UploadFile, File, Response, Request, Form
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
import uvicorn
import json
import tempfile
from pptx import Presentation
from typing import Dict
import requests
import os

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# HTML template for the form
HTML_FORM = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PowerPoint Slide Generator</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            max-width: 800px;
            margin: 0 auto;
            padding: 20px;
            background-color: #f5f5f5;
        }
        .container {
            background: white;
            padding: 30px;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        h1 {
            color: #333;
            text-align: center;
            margin-bottom: 30px;
        }
        .form-group {
            margin-bottom: 20px;
        }
        label {
            display: block;
            margin-bottom: 5px;
            font-weight: bold;
            color: #555;
        }
        input[type="text"], textarea {
            width: 100%;
            padding: 10px;
            border: 1px solid #ddd;
            border-radius: 5px;
            font-size: 16px;
            box-sizing: border-box;
        }
        textarea {
            height: 100px;
            resize: vertical;
        }
        .slide-input {
            border: 1px solid #e0e0e0;
            padding: 15px;
            margin-bottom: 15px;
            border-radius: 5px;
            background-color: #fafafa;
        }
        .slide-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 10px;
        }
        .slide-number {
            font-weight: bold;
            color: #007bff;
        }
        .remove-slide {
            background: #dc3545;
            color: white;
            border: none;
            padding: 5px 10px;
            border-radius: 3px;
            cursor: pointer;
        }
        button {
            background-color: #007bff;
            color: white;
            padding: 12px 30px;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            font-size: 16px;
            width: 100%;
            margin-top: 20px;
        }
        button:hover {
            background-color: #0056b3;
        }
        .add-slide-btn {
            background-color: #28a745;
            margin-top: 10px;
        }
        .add-slide-btn:hover {
            background-color: #218838;
        }
        .loading {
            display: none;
            text-align: center;
            color: #666;
            margin-top: 20px;
        }
        .result {
            margin-top: 20px;
            padding: 15px;
            border-radius: 5px;
            display: none;
        }
        .success {
            background-color: #d4edda;
            border: 1px solid #c3e6cb;
            color: #155724;
        }
        .error {
            background-color: #f8d7da;
            border: 1px solid #f5c6cb;
            color: #721c24;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>🎯 PowerPoint Slide Generator</h1>
        <p style="text-align: center; color: #666; margin-bottom: 30px;">
            Create professional PowerPoint presentations with AI assistance
        </p>

        <form id="slideForm">
            <div id="slidesContainer">
                <div class="slide-input" data-slide="1">
                    <div class="slide-header">
                        <span class="slide-number">Slide 1</span>
                        <button type="button" class="remove-slide" onclick="removeSlide(1)" style="display: none;">Remove</button>
                    </div>
                    <div class="form-group">
                        <label for="title1">Slide Title:</label>
                        <input type="text" id="title1" name="title1" placeholder="Enter slide title..." required>
                    </div>
                    <div class="form-group">
                        <label for="content1">Slide Content:</label>
                        <textarea id="content1" name="content1" placeholder="Enter slide content..." required></textarea>
                    </div>
                </div>
            </div>

            <button type="button" class="add-slide-btn" onclick="addSlide()">+ Add Another Slide</button>
            <button type="submit">🚀 Generate PowerPoint Presentation</button>
        </form>

        <div class="loading" id="loading">
            <p>🎨 Generating your PowerPoint presentation...</p>
            <p>This may take a few moments...</p>
        </div>

        <div class="result" id="result"></div>
    </div>

    <script>
        let slideCount = 1;

        function addSlide() {
            slideCount++;
            const container = document.getElementById('slidesContainer');

            const slideDiv = document.createElement('div');
            slideDiv.className = 'slide-input';
            slideDiv.setAttribute('data-slide', slideCount);

            slideDiv.innerHTML = `
                <div class="slide-header">
                    <span class="slide-number">Slide ${slideCount}</span>
                    <button type="button" class="remove-slide" onclick="removeSlide(${slideCount})">Remove</button>
                </div>
                <div class="form-group">
                    <label for="title${slideCount}">Slide Title:</label>
                    <input type="text" id="title${slideCount}" name="title${slideCount}" placeholder="Enter slide title..." required>
                </div>
                <div class="form-group">
                    <label for="content${slideCount}">Slide Content:</label>
                    <textarea id="content${slideCount}" name="content${slideCount}" placeholder="Enter slide content..." required></textarea>
                </div>
            `;

            container.appendChild(slideDiv);

            // Show remove button for first slide if we have more than one
            if (slideCount > 1) {
                document.querySelector('.remove-slide').style.display = 'block';
            }
        }

        function removeSlide(slideNum) {
            if (slideCount <= 1) return;

            const slideToRemove = document.querySelector(`[data-slide="${slideNum}"]`);
            slideToRemove.remove();
            slideCount--;

            // Hide remove button if only one slide left
            if (slideCount === 1) {
                const remainingRemoveBtn = document.querySelector('.remove-slide');
                if (remainingRemoveBtn) {
                    remainingRemoveBtn.style.display = 'none';
                }
            }

            // Renumber remaining slides
            const slides = document.querySelectorAll('.slide-input');
            slides.forEach((slide, index) => {
                const newNum = index + 1;
                slide.setAttribute('data-slide', newNum);
                slide.querySelector('.slide-number').textContent = `Slide ${newNum}`;
                slide.querySelector('.remove-slide').setAttribute('onclick', `removeSlide(${newNum})`);
                slide.querySelector('input').id = `title${newNum}`;
                slide.querySelector('input').name = `title${newNum}`;
                slide.querySelector('textarea').id = `content${newNum}`;
                slide.querySelector('textarea').name = `content${newNum}`;
                slide.querySelector('label[for*="title"]').setAttribute('for', `title${newNum}`);
                slide.querySelector('label[for*="content"]').setAttribute('for', `content${newNum}`);
            });
        }

        document.getElementById('slideForm').addEventListener('submit', async function(e) {
            e.preventDefault();

            const loading = document.getElementById('loading');
            const result = document.getElementById('result');
            const submitBtn = this.querySelector('button[type="submit"]');

            loading.style.display = 'block';
            result.style.display = 'none';
            submitBtn.disabled = true;
            submitBtn.textContent = 'Generating...';

            try {
                // Collect form data
                const formData = new FormData(this);
                const slides = [];

                for (let i = 1; i <= slideCount; i++) {
                    const title = formData.get(`title${i}`);
                    const content = formData.get(`content${i}`);

                    if (title && content) {
                        slides.push({
                            title: title,
                            content: content
                        });
                    }
                }

                // Send to N8N webhook
                const response = await fetch('https://sd-n8n.duckdns.org/webhook-test/slider', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                    },
                    body: JSON.stringify({
                        slides: slides
                    })
                });

                if (response.ok) {
                    const data = await response.json();
                    result.className = 'result success';
                    result.innerHTML = `
                        <h3>✅ Success!</h3>
                        <p>Your PowerPoint presentation has been generated and enhanced with AI!</p>
                        <p><strong>Download:</strong> <a href="${data.download_url}" target="_blank">Click here to download your presentation</a></p>
                        <p><em>Your content has been professionally enhanced with better formatting and bullet points.</em></p>
                    `;
                } else {
                    throw new Error('Failed to generate presentation');
                }

            } catch (error) {
                console.error('Error:', error);
                result.className = 'result error';
                result.innerHTML = `
                    <h3>❌ Error</h3>
                    <p>Sorry, there was an error generating your presentation. Please try again.</p>
                    <p>Error details: ${error.message}</p>
                `;
            } finally {
                loading.style.display = 'none';
                submitBtn.disabled = false;
                submitBtn.textContent = '🚀 Generate PowerPoint Presentation';
                result.style.display = 'block';
            }
        });
    </script>
</body>
</html>
"""

@app.get("/", response_class=HTMLResponse)
async def root():
    return HTML_FORM

def create_pptx_from_json(slides_json: Dict, output_path: str):
    prs = Presentation()
    for slide_data in slides_json.get("slides", []):
        slide_layout = prs.slide_layouts[0]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        content = slide.placeholders[1]
        title.text = slide_data.get("title", "")
        content.text = slide_data.get("content", "")
    prs.save(output_path)

@app.post("/generate-pptx")
async def generate_pptx(request: Request):
    # Accept raw JSON instead of file upload
    data = await request.json()
    slides_json = data  # Data is already parsed JSON
    
    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
        create_pptx_from_json(slides_json, tmp.name)
        # Return JSON response with download URL instead of direct file
        download_url = f"https://slider.sd-ai.co.uk/download/{tmp.name.split('/')[-1]}"
        return {"download_url": download_url, "filename": "slides.pptx"}

@app.get("/download/{filename}")
async def download_file(filename: str):
    # In a production environment, you'd want to store files more securely
    # For now, we'll serve from temp directory
    file_path = f"/tmp/{filename}"
    if not os.path.exists(file_path):
        return {"error": "File not found"}
    return FileResponse(file_path, filename="slides.pptx", media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation")

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8010)
