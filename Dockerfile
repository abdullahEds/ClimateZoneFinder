# Use Python 3.11 slim image
FROM python:3.11-slim

# Set working directory
WORKDIR /app

# Install system dependencies for Kaleido (Plotly static image export)
RUN apt-get update && apt-get install -y --no-install-recommends \
    wget \
    gnupg \
    libgl1 \
    libglib2.0-0 \
    libsm6 \
    libxrender1 \
    libxext6 \
    && wget -q -O - https://dl-ssl.google.com/linux/linux_signing_key.pub | apt-key add - \
    && echo "deb [arch=amd64] http://dl.google.com/linux/chrome/deb/ stable main" | tee /etc/apt/sources.list.d/google-chrome.list \
    && apt-get update && apt-get install -y --no-install-recommends google-chrome-stable \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements and install
COPY requirements_api.txt .
RUN pip install --no-cache-dir -r requirements_api.txt

# Copy the entire project
COPY . .

# Expose port 8080 (Cloud Run expects this)
EXPOSE 8080

# Run the app with Uvicorn
CMD ["uvicorn", "report_api:app", "--host", "0.0.0.0", "--port", "8080"]
