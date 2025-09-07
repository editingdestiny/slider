// Simple form handler for quick presentation generation
document.addEventListener('DOMContentLoaded', function() {
    const quickForm = document.getElementById('quickForm');
    if (quickForm) {
        quickForm.addEventListener('submit', handleQuickGeneration);
    }
});

async function handleQuickGeneration(e) {
    e.preventDefault();

    const loading = document.getElementById('loading');
    const loadingMessage = document.getElementById('loading-message');
    const result = document.getElementById('result');
    const submitBtn = e.target.querySelector('button[type="submit"]');
    const formData = new FormData(e.target);

    const searchPhrase = formData.get('searchPhrase');
    const numberOfSlides = parseInt(formData.get('numberOfSlides'));

    // Gather customization options
    const customization = {
        slide_bg_color: document.getElementById('slideBgColor').value,
        title_font_color: document.getElementById('titleFontColor').value,
        title_bg_color: document.getElementById('titleBgColor').value,
        body_text_color: document.getElementById('bodyTextColor').value,
        title_position: document.getElementById('titlePosition').value,
        font_size: parseInt(document.getElementById('fontSize').value)
    };

    // Prevent double submissions
    if (submitBtn.disabled) return;

    // Show loading state and start cycling messages
    loading.style.display = 'block';
    result.style.display = 'none';
    submitBtn.disabled = true;
    submitBtn.textContent = 'Generating...';

    const messages = [
        "Contacting AI assistant...",
        "Gathering information...",
        "Generating slides...",
        "Creating charts and tables...",
        "This can take a few minutes...",
        "Almost there...",
        "Finalizing your presentation...",
        "Preparing your download...",

    ];
    let messageIndex = 0;
    loadingMessage.textContent = messages[messageIndex];
    const messageInterval = setInterval(() => {
        messageIndex = (messageIndex + 1) % messages.length;
        loadingMessage.textContent = messages[messageIndex];
    }, 10000); // Change message every 10 seconds

    try {
        const response = await fetch('/generate-slides-from-search', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                search_phrase: searchPhrase,
                number_of_slides: numberOfSlides,
                customization: customization
            })
        });

        if (!response.ok) {
            throw new Error(`Server error: ${response.status}`);
        }

        // Handle response - either PowerPoint file or JSON error
        const contentType = response.headers.get('content-type') || '';
        
        if (contentType.includes('application/vnd.openxmlformats-officedocument.presentationml.presentation')) {
            // Handle PowerPoint file download
            const blob = await response.blob();
            
            if (blob.size === 0) {
                throw new Error('Received empty file');
            }
            
            // Download file
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
                document.body.removeChild(a);
            }, 100);

            // Show success message
            result.className = 'result success';
            result.innerHTML = `
                <h3>✅ Presentation Downloaded Successfully!</h3>
                <p><strong>Topic:</strong> ${searchPhrase}</p>
                <p><strong>Slides:</strong> ${numberOfSlides}</p>
                <p><strong>File:</strong> ${filename}</p>
                <p>The PowerPoint file has been downloaded to your computer.</p>
            `;
        } else if (contentType.includes('application/json')) {
            // Handle JSON error response (webhook failure case)
            const errorData = await response.json();
            const errorMessage = errorData.error || errorData.message || 'Unknown error occurred';
            throw new Error(errorMessage);
        } else {
            throw new Error('Unexpected response format from server');
        }

    } catch (error) {
        result.className = 'result error';
        result.innerHTML = `
            <h3>❌ Error</h3>
            <p>Sorry, there was an error generating your presentation.</p>
            <p>Error: ${error.message}</p>
        `;
    } finally {
        // Reset UI
        clearInterval(messageInterval); // Stop cycling messages
        loading.style.display = 'none';
        submitBtn.disabled = false;
        submitBtn.textContent = '🚀 Generate Presentation';
        result.style.display = 'block';
    }
}
