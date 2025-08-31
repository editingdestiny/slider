# 🎯 AI-Powered PowerPoint Slide Generator

An intelligent web application that creates **professional executive presentations** using AI content enhancement. Features dark blue backgrounds with white text for C-level presentations.

## 🌟 Features

- **🎨 Professional Design**: Dark blue (#1e3a8a) background with white text
- **🤖 AI Content Enhancement**: GPT-4 powered executive-level content creation
- **📊 Smart Charts**: Automatic chart generation with business color schemes
- **🔒 Secure Downloads**: Direct download links for generated files
- **🐳 Containerized**: Docker deployment with Traefik routing
- **⚡ FastAPI Backend**: High-performance Python service
- **🔄 N8N Workflow**: Complete AI automation pipeline

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
- `POST /ai-generate-pptx` - AI-enhanced PowerPoint generation
- `POST /test-professional-slide` - Test professional styling
- `GET /download/{filename}` - Download generated files
- `GET /health` - Service health check

### N8N Integration
- **Webhook URL**: `https://slider-n8n.your-domain.com/webhook/slider`
- **Method**: POST
- **Content-Type**: application/json

## 🎨 Professional Design System

- **Background**: Executive dark blue (#1e3a8a)
- **Title Text**: White, 32pt, Bold
- **Headline Text**: White, 24pt, Bold
- **Content Text**: White, 18pt, Professional
- **Chart Text**: White, 16pt, Clear
- **Layout**: Clean, executive-focused design
- **Color Scheme**: Business-appropriate chart colors

### Example Transformation

**Input:**
```json
{
  "title": "Sales Update",
  "content": "We did good this month"
}
```

**AI Enhanced Output:**
```json
{
  "title": "Executive Sales Performance Report",
  "headline": "Record-Breaking Monthly Results",
  "content": "• 150 new customers acquired\n• $500K revenue milestone achieved\n• 40% increase in customer acquisition\n• Team performance exceeded targets",
  "chartType": "bar",
  "chartData": {
    "labels": ["Q1", "Q2", "Q3", "Q4"],
    "values": [100, 150, 180, 220],
    "colors": ["#60A5FA", "#3B82F6", "#1D4ED8", "#1E3A8A"],
    "title": "Quarterly Revenue Performance"
  }
}
```

**Result**: Professional PowerPoint with dark blue background, white text, and business-appropriate charts.

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
Webhook Request → N8N AI Agent → Content Enhancement → Python Service → Professional PowerPoint → Response
     ↓              ↓              ↓                      ↓                      ↓              ↓
Raw Content     Prompt Analysis   Executive Content    Dark Blue Theme      .pptx File    Download URL
```

### AI Workflow Steps:
1. **Webhook**: Receives user content and requirements
2. **AI Agent 1**: Analyzes content and creates enhancement prompt
3. **AI Agent 2**: Generates professional executive content in JSON
4. **Python Service**: Creates PowerPoint with dark blue background
5. **Response**: Returns secure download URL

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
