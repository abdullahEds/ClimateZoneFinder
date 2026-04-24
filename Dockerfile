# Use Python 3.11 slim image
FROM python:3.11-slim

# Prevent Python from writing .pyc files & enable logs immediately
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

# Set working directory
WORKDIR /app

# Install system dependencies (optional but useful for pandas, matplotlib, etc.)
RUN apt-get update && apt-get install -y \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements first (for better caching)
COPY requirements_api.txt .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements_api.txt

# Copy the entire project
COPY . .

# Cloud Run uses PORT env variable (default 8080)
ENV PORT=8080

# Expose port (informational)
EXPOSE 8080

# Run the FastAPI app
CMD ["sh", "-c", "uvicorn report_api:app --host 0.0.0.0 --port ${PORT}"]
