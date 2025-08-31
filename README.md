# 🎯 AI-Powered PowerPoint Slide Generator

An intelligent web application that creates professional PowerPoint presentations using AI content enhancement. Users can input basic slide content through a beautiful web interface, which gets enhanced by GPT-4 and converted into polished PowerPoint presentations.

## 🌟 Features

- **🎨 Beautiful Web Interface**: Modern, responsive form for slide creation
- **🤖 AI Content Enhancement**: GPT-4 powered content improvement
- **📊 Professional Presentations**: Auto-generated PowerPoint files
- **🔒 Secure Downloads**: Direct download links for generated files
- **🐳 Containerized**: Docker deployment with Traefik routing
- **⚡ FastAPI Backend**: High-performance Python service
- **🔄 N8N Workflow**: Automated content processing pipeline

## 🚀 Quick Start

### Prerequisites
- Docker & Docker Compose
- N8N instance (for workflow automation)
- OpenAI API key (for AI enhancement)

### Installation

1. **Clone the repository**
   ```bash
   git clone <your-repo-url>
   cd slider
   ```

2. **Start the services**
   ```bash
   docker compose up --build -d
   ```

3. **Access the application**
   - Web Interface: https://slider.sd-ai.co.uk
   - N8N Dashboard: https://your-n8n-instance.com

### Configuration

1. **Import N8N Workflow**
   - Open your N8N dashboard
   - Import `n8n-workflow.json`
   - Add your OpenAI API key to the OpenAI node
   - Update webhook URL in the workflow

2. **Environment Setup**
   - Ensure Traefik is configured for `slider.sd-ai.co.uk`
   - Verify Docker network connectivity

## 📁 Project Structure

```
slider/
├── docker-compose.yml          # Docker services configuration
├── n8n-workflow.json          # N8N automation workflow
├── .gitignore                 # Git ignore rules
└── python_service/
    ├── main.py               # FastAPI application
    ├── requirements.txt      # Python dependencies
    ├── Dockerfile           # Python service container
    └── README.md            # Service documentation
```

## 🔧 API Endpoints

### Web Interface
- `GET /` - Main application interface
- `POST /generate-pptx` - Generate PowerPoint from JSON
- `GET /download/{filename}` - Download generated files

### N8N Integration
- Webhook endpoint processes form submissions
- AI enhancement via OpenAI GPT-4
- Automated PowerPoint generation

## 🎨 AI Enhancement Features

The system uses advanced AI to transform basic user input into professional presentation content:

- **Smart Titles**: Concise, impactful slide titles
- **Bullet Points**: 3-5 optimized bullet points per slide
- **Professional Language**: Business-appropriate terminology
- **Content Structure**: Logical flow and organization
- **Presentation Ready**: Optimized for visual impact

### Example Transformation

**Input:**
```json
{
  "title": "Our Services",
  "content": "We provide good services to customers"
}
```

**AI Enhanced Output:**
```json
{
  "title": "Comprehensive Service Solutions",
  "content": "• Client-centric service delivery\n• Quality-driven solutions\n• 24/7 customer support\n• Measurable business results"
}
```

## 🐳 Docker Services

### Python Service
- **Image**: Custom Python 3.11 with FastAPI
- **Port**: 8010 (internal)
- **Features**: PowerPoint generation, file serving

### N8N (External)
- **Purpose**: Workflow automation and AI processing
- **Integration**: Webhook processing and API calls

### Traefik (External)
- **Purpose**: Reverse proxy and SSL termination
- **Routing**: HTTPS routing to services

## 🔐 Security Features

- HTTPS encryption via Traefik
- Secure file download links
- Input validation and sanitization
- Container isolation
- CORS protection

## 📊 Workflow Architecture

```
User Form → N8N Webhook → AI Enhancement → Python Service → PowerPoint → Download Link
    ↓           ↓              ↓              ↓              ↓           ↓
  HTML Form   JSON Data     GPT-4 Processing FastAPI       .pptx      Secure URL
```

## 🛠️ Development

### Local Development
```bash
# Start services
docker compose up --build

# View logs
docker logs slider

# Access API documentation
curl http://localhost:8010/docs
```

### Testing
```bash
# Test PowerPoint generation
curl -X POST "http://localhost:8010/generate-pptx" \
  -H "Content-Type: application/json" \
  -d '{"slides": [{"title": "Test", "content": "Test content"}]}'
```

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## 📝 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🙏 Acknowledgments

- **FastAPI**: High-performance web framework
- **python-pptx**: PowerPoint file generation
- **N8N**: Workflow automation platform
- **OpenAI**: AI content enhancement
- **Traefik**: Modern reverse proxy

## 📞 Support

For support and questions:
- Create an issue in the repository
- Check the documentation
- Review the N8N workflow configuration

---

**Made with ❤️ for creating better presentations, faster.**
