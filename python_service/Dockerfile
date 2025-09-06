FROM python:3.11-slim

WORKDIR /app

# Install system dependencies for fonts
RUN apt-get update && apt-get install -y --no-install-recommends \
    fontconfig \
    && rm -rf /var/lib/apt/lists/*

# Copy and install fonts
COPY fonts /usr/share/fonts/truetype/
RUN fc-cache -f -v

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8010"]
