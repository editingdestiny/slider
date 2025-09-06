let slideCount = 1;

// Test server connectivity on page load
async function testConnectivity() {
    try {
        console.log('Testing server connectivity...');
        const response = await fetch('https://slider.sd-ai.co.uk/health');
        if (response.ok) {
            console.log('✅ Server connectivity test passed');
        } else {
            console.warn('⚠️ Server connectivity test failed:', response.status);
        }
    } catch (error) {
        console.error('❌ Server connectivity test failed:', error);
    }
}

// Run connectivity test when page loads
document.addEventListener('DOMContentLoaded', function() {
    console.log('🔄 Page loaded, setting up event listeners...');
    testConnectivity();
    
    // Add event listener for quick form ONLY
    const quickForm = document.getElementById('quickForm');
    if (quickForm) {
        quickForm.addEventListener('submit', handleQuickGeneration);
        console.log('✅ Single event listener added to quick form');
    } else {
        console.error('❌ Quick form not found! Check HTML structure.');
    }
});

// Handle quick generation form submission
async function handleQuickGeneration(e) {
    console.log('🚀 Generate button clicked! Form submitted.');
    e.preventDefault();
    e.stopPropagation();

    const loading = document.getElementById('loading');
    const result = document.getElementById('result');
    const submitBtn = e.target.querySelector('button[type="submit"]');
    const formData = new FormData(e.target);

    const searchPhrase = formData.get('searchPhrase');
    const numberOfSlides = parseInt(formData.get('numberOfSlides'));

    // Prevent double submissions
    if (submitBtn.disabled) {
        console.log('Button already disabled, ignoring submission');
        return;
    }

    loading.style.display = 'block';
    result.style.display = 'none';
    submitBtn.disabled = true;
    submitBtn.textContent = 'Generating...';

    try {
        console.log(`Generating presentation for: "${searchPhrase}" with ${numberOfSlides} slides`);

        const requestBody = {
            search_phrase: searchPhrase,
            number_of_slides: numberOfSlides
        };

        console.log('Sending request to: /generate-slides-from-search');
        console.log('Request body:', requestBody);

        const response = await fetch('/generate-slides-from-search', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(requestBody)
        });

        console.log('Response status:', response.status);
        console.log('Response content-type:', response.headers.get('content-type'));

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        // Check if we got a PowerPoint file
        const contentType = response.headers.get('content-type') || '';
        if (contentType.includes('application/vnd.openxmlformats-officedocument.presentationml.presentation')) {
            const blob = await response.blob();
            console.log('Blob received, size:', blob.size);
            
            if (blob.size === 0) {
                throw new Error('Received empty file');
            }
            
            // Create download
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            const filename = `${searchPhrase.replace(/\s+/g, '_')}_Presentation.pptx`;
            
            a.href = url;
            a.download = filename;
            a.style.display = 'none';
            document.body.appendChild(a);
            a.click();
            
            // Cleanup
            setTimeout(() => {
                window.URL.revokeObjectURL(url);
                if (document.body.contains(a)) {
                    document.body.removeChild(a);
                }
            }, 100);

            result.className = 'result success';
            result.innerHTML = `
                <h3>✅ Presentation Downloaded Successfully!</h3>
                <p><strong>Topic:</strong> ${searchPhrase}</p>
                <p><strong>Slides:</strong> ${numberOfSlides}</p>
                <p><strong>File:</strong> ${filename}</p>
                <p>The PowerPoint file has been automatically downloaded to your computer.</p>
            `;
        } else {
            // Handle JSON response (fallback)
            const data = await response.json();
            console.log('Response data:', data);

            result.className = 'result success';
            result.innerHTML = `
                <h3>✅ Success!</h3>
                <p>Your presentation has been generated!</p>
                <p><strong>Topic:</strong> ${data.search_phrase}</p>
                <p><strong>Slides Generated:</strong> ${data.slides_generated || numberOfSlides}</p>
                <p><strong>Download:</strong> <a href="${data.download_url}" target="_blank">Click here to download your presentation</a></p>
            `;
        }

    } catch (error) {
        console.error('Error details:', error);

        result.className = 'result error';
        result.innerHTML = `
            <h3>❌ Error</h3>
            <p>Sorry, there was an error generating your presentation. Please try again.</p>
            <p>Error details: ${error.message || 'Unknown error'}</p>
            <p><small>Check the browser console for more details.</small></p>
        `;
    } finally {
        loading.style.display = 'none';
        submitBtn.disabled = false;
        submitBtn.textContent = '🚀 Generate Presentation';
        result.style.display = 'block';
    }
}

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
            <label for="headline${slideCount}">Slide Headline:</label>
            <input type="text" id="headline${slideCount}" name="headline${slideCount}" placeholder="Enter slide headline..." required>
        </div>
        <div class="form-group">
            <label for="content${slideCount}">Slide Content/Data:</label>
            <textarea id="content${slideCount}" name="content${slideCount}" placeholder="Enter slide content or data points..." required></textarea>
        </div>
        <div class="form-group">
            <label for="backgroundColor${slideCount}">Background Color:</label>
            <select id="backgroundColor${slideCount}" name="backgroundColor${slideCount}">
                <!-- Whites and Grays -->
                <optgroup label="Whites & Grays">
                    <option value="#FFFFFF">White</option>
                    <option value="#F8F9FA">Light Gray</option>
                    <option value="#E9ECEF">Very Light Gray</option>
                    <option value="#DEE2E6">Light Gray</option>
                    <option value="#CED4DA">Medium Gray</option>
                    <option value="#ADB5BD">Gray</option>
                    <option value="#6C757D">Dark Gray</option>
                    <option value="#495057">Darker Gray</option>
                    <option value="#343A40">Very Dark Gray</option>
                    <option value="#212529">Almost Black</option>
                </optgroup>

                <!-- Blues -->
                <optgroup label="Blues">
                    <option value="#E3F2FD">Light Blue</option>
                    <option value="#BBDEFB">Very Light Blue</option>
                    <option value="#90CAF9">Light Sky Blue</option>
                    <option value="#64B5F6">Sky Blue</option>
                    <option value="#42A5F5">Blue</option>
                    <option value="#2196F3">Medium Blue</option>
                    <option value="#1E88E5">Blue</option>
                    <option value="#1976D2">Dark Blue</option>
                    <option value="#1565C0">Darker Blue</option>
                    <option value="#0D47A1">Very Dark Blue</option>
                </optgroup>

                <!-- Greens -->
                <optgroup label="Greens">
                    <option value="#E8F5E8">Light Green</option>
                    <option value="#C8E6C9">Very Light Green</option>
                    <option value="#A5D6A7">Light Mint</option>
                    <option value="#81C784">Mint Green</option>
                    <option value="#66BB6A">Green</option>
                    <option value="#4CAF50">Medium Green</option>
                    <option value="#43A047">Forest Green</option>
                    <option value="#388E3C">Dark Green</option>
                    <option value="#2E7D32">Darker Green</option>
                    <option value="#1B5E20">Very Dark Green</option>
                </optgroup>

                <!-- Purples -->
                <optgroup label="Purples">
                    <option value="#F3E5F5">Light Purple</option>
                    <option value="#E1BEE7">Very Light Purple</option>
                    <option value="#CE93D8">Light Lavender</option>
                    <option value="#BA68C8">Lavender</option>
                    <option value="#AB47BC">Purple</option>
                    <option value="#9C27B0">Medium Purple</option>
                    <option value="#8E24AA">Deep Purple</option>
                    <option value="#7B1FA2">Dark Purple</option>
                    <option value="#6A1B9A">Darker Purple</option>
                    <option value="#4A148C">Very Dark Purple</option>
                </optgroup>

                <!-- Reds and Pinks -->
                <optgroup label="Reds & Pinks">
                    <option value="#FFEBEE">Light Red</option>
                    <option value="#FFCDD2">Very Light Red</option>
                    <option value="#EF9A9A">Light Coral</option>
                    <option value="#E57373">Coral</option>
                    <option value="#EF5350">Red</option>
                    <option value="#F44336">Medium Red</option>
                    <option value="#E53935">Crimson</option>
                    <option value="#D32F2F">Dark Red</option>
                    <option value="#C62828">Darker Red</option>
                    <option value="#B71C1C">Very Dark Red</option>
                    <option value="#FCE4EC">Light Pink</option>
                    <option value="#F8BBD9">Pink</option>
                    <option value="#F48FB1">Hot Pink</option>
                    <option value="#EC407A">Deep Pink</option>
                </optgroup>

                <!-- Oranges and Yellows -->
                <optgroup label="Oranges & Yellows">
                    <option value="#FFF3E0">Light Orange</option>
                    <option value="#FFE0B2">Very Light Orange</option>
                    <option value="#FFCC80">Light Peach</option>
                    <option value="#FFB74D">Peach</option>
                    <option value="#FFA726">Orange</option>
                    <option value="#FF9800">Medium Orange</option>
                    <option value="#FB8C00">Dark Orange</option>
                    <option value="#F57C00">Darker Orange</option>
                    <option value="#EF6C00">Very Dark Orange</option>
                    <option value="#FFF9C4">Light Yellow</option>
                    <option value="#FFF59D">Pale Yellow</option>
                    <option value="#FFF176">Yellow</option>
                    <option value="#FFEB3B">Bright Yellow</option>
                    <option value="#FBC02D">Gold</option>
                </optgroup>

                <!-- Teals and Cyans -->
                <optgroup label="Teals & Cyans">
                    <option value="#E0F2F1">Light Teal</option>
                    <option value="#B2DFDB">Very Light Teal</option>
                    <option value="#80CBC4">Light Teal</option>
                    <option value="#4DB6AC">Teal</option>
                    <option value="#26A69A">Medium Teal</option>
                    <option value="#009688">Dark Teal</option>
                    <option value="#00897B">Darker Teal</option>
                    <option value="#00796B">Very Dark Teal</option>
                    <option value="#E0F7FA">Light Cyan</option>
                    <option value="#B2EBF2">Very Light Cyan</option>
                    <option value="#80DEEA">Light Cyan</option>
                    <option value="#4DD0E1">Cyan</option>
                    <option value="#26C6DA">Medium Cyan</option>
                    <option value="#00BCD4">Dark Cyan</option>
                </optgroup>

                <!-- Browns and Beiges -->
                <optgroup label="Browns & Beiges">
                    <option value="#EFEBE9">Light Brown</option>
                    <option value="#D7CCC8">Very Light Brown</option>
                    <option value="#BCAAA4">Light Taupe</option>
                    <option value="#A1887F">Taupe</option>
                    <option value="#8D6E63">Brown</option>
                    <option value="#795548">Medium Brown</option>
                    <option value="#6D4C41">Dark Brown</option>
                    <option value="#5D4037">Darker Brown</option>
                    <option value="#4E342E">Very Dark Brown</option>
                    <option value="#3E2723">Almost Black Brown</option>
                </optgroup>

                <!-- Professional Colors -->
                <optgroup label="Professional Colors">
                    <option value="#1565C0">Corporate Blue</option>
                    <option value="#2E7D32">Professional Green</option>
                    <option value="#6D4C41">Executive Brown</option>
                    <option value="#424242">Business Gray</option>
                    <option value="#263238">Slate</option>
                    <option value="#004D40">Deep Teal</option>
                    <option value="#1A237E">Navy Blue</option>
                    <option value="#880E4F">Burgundy</option>
                    <option value="#BF360C">Rust</option>
                    <option value="#827717">Olive</option>
                </optgroup>
            </select>
        </div>
        <div class="form-group">
            <label for="chartType${slideCount}">Chart Type (Optional):</label>
            <select id="chartType${slideCount}" name="chartType${slideCount}" onchange="toggleChartData(${slideCount})">
                <option value="">No Chart</option>
                <option value="bar">Bar Chart</option>
                <option value="pie">Pie Chart</option>
                <option value="line">Line Chart</option>
                <option value="column">Column Chart</option>
                <option value="area">Area Chart</option>
                <option value="doughnut">Doughnut Chart</option>
            </select>
        </div>
        <div class="form-group chart-data" id="chartData${slideCount}" style="display: none;">
            <label for="chartDataInput${slideCount}">Chart Data (JSON format):</label>
            <textarea id="chartDataInput${slideCount}" name="chartDataInput${slideCount}" placeholder='Example: {"labels": ["Q1", "Q2", "Q3"], "values": [100, 150, 200]}' rows="3"></textarea>
            <small style="color: #666;">Enter data as JSON with "labels" and "values" arrays</small>
        </div>
    `;

    container.appendChild(slideDiv);

    // Show remove button for first slide if we have more than one
    if (slideCount > 1) {
        document.querySelector('.remove-slide').style.display = 'block';
    }
}

function toggleChartData(slideNum) {
    const chartType = document.getElementById(`chartType${slideNum}`);
    const chartDataDiv = document.getElementById(`chartData${slideNum}`);

    if (chartType.value) {
        chartDataDiv.style.display = 'block';
    } else {
        chartDataDiv.style.display = 'none';
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

        // Update all form field IDs and names
        const fields = ['title', 'headline', 'content', 'chartType', 'chartDataInput', 'backgroundColor'];
        fields.forEach(field => {
            const element = slide.querySelector(`#${field}${slideNum}`);
            if (element) {
                element.id = `${field}${newNum}`;
                element.name = `${field}${newNum}`;
                const label = slide.querySelector(`label[for="${field}${slideNum}"]`);
                if (label) {
                    label.setAttribute('for', `${field}${newNum}`);
                }
            }
        });

        // Update chart type onchange
        const chartTypeSelect = slide.querySelector(`#chartType${newNum}`);
        if (chartTypeSelect) {
            chartTypeSelect.setAttribute('onchange', `toggleChartData(${newNum})`);
        }

        // Update chart data div ID
        const chartDataDiv = slide.querySelector(`#chartData${slideNum}`);
        if (chartDataDiv) {
            chartDataDiv.id = `chartData${newNum}`;
        }
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
            const headline = formData.get(`headline${i}`);
            const content = formData.get(`content${i}`);
            const chartType = formData.get(`chartType${i}`);
            const chartData = formData.get(`chartDataInput${i}`);
            const backgroundColor = formData.get(`backgroundColor${i}`) || '#FFFFFF';

            if (title && headline && content) {
                const slideData = {
                    title: title,
                    headline: headline,
                    content: content,
                    backgroundColor: backgroundColor
                };

                // Add chart data if chart type is selected
                if (chartType) {
                    slideData.chartType = chartType;
                    if (chartData) {
                        try {
                            slideData.chartData = JSON.parse(chartData);
                        } catch (e) {
                            slideData.chartData = chartData; // Keep as string if JSON parsing fails
                        }
                    }
                }

                slides.push(slideData);
            }
        }

        // Validate that we have slides
        if (slides.length === 0) {
            throw new Error('No valid slides found. Please fill in at least one slide.');
        }

        console.log('Sending slides data:', slides);

        // Send directly to our local PowerPoint generation endpoint
        console.log('Making fetch request to:', 'https://slider.sd-ai.co.uk/ai-generate-pptx');
        const response = await fetch('https://slider.sd-ai.co.uk/ai-generate-pptx', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                slides: slides
            })
        });

        console.log('Fetch response received:', response);
        console.log('Response status:', response.status);
        console.log('Response headers:', [...response.headers.entries()]);

        if (response.ok) {
            const data = await response.json();
            console.log('Response data:', data);

            result.className = 'result success';
            result.innerHTML = `
                <h3>✅ Success!</h3>
                <p>Your PowerPoint presentation has been generated!</p>
                <p><strong>Download:</strong> <a href="${data.download_url}" target="_blank">Click here to download your presentation</a></p>
                <div class="debug-info">
                    <strong>Debug Info:</strong><br>
                    Slides sent: ${slides.length}<br>
                    Response status: ${response.status}
                </div>
            `;
        } else {
            const errorText = await response.text();
            console.error('Response error:', response.status, errorText);
            throw new Error(`Server error: ${response.status} - ${errorText}`);
        }

    } catch (error) {
        console.error('Error details:', error);
        console.error('Error name:', error.name);
        console.error('Error message:', error.message);
        console.error('Error stack:', error.stack);

        let errorMessage = error.message;
        let debugInfo = `
                ${error.stack || 'No stack trace available'}
                Error Name: ${error.name}
                Error Message: ${error.message}`;

        // Provide more specific error messages
        if (error.name === 'TypeError' && error.message.includes('Failed to fetch')) {
            errorMessage = 'Network error: Unable to connect to the server. Please check your internet connection and try again.';
            debugInfo += '\nPossible causes:\n- Network connectivity issues\n- CORS policy blocking the request\n- Server temporarily unavailable';
        }

        result.className = 'result error';
        result.innerHTML = `
            <h3>❌ Error</h3>
            <p>Sorry, there was an error generating your presentation. Please try again.</p>
            <p>Error details: ${errorMessage}</p>
            <div class="debug-info">
                <strong>Debug Info:</strong><br>
                ${debugInfo}
            </div>
        `;
    } finally {
        loading.style.display = 'none';
        submitBtn.disabled = false;
        submitBtn.textContent = '🚀 Generate PowerPoint Presentation';
        result.style.display = 'block';
    }
});
