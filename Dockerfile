# Use Python 3.11 slim image
FROM python:3.11-slim

# Prevent Python from writing pyc files and enable unbuffered logs
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

# Set working directory
WORKDIR /app

# System libs required by matplotlib/Pillow on slim images
RUN apt-get update && apt-get install -y --no-install-recommends \
    libgl1 \
    libglib2.0-0 \
    libsm6 \
    libxrender1 \
    libxext6 \
    # Uncomment the next line to enable PDF export via the output_format=pdf parameter:
    # libreoffice \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements first for Docker layer caching
COPY requirements_api.txt .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements_api.txt

# Copy SolarGIS data file explicitly (required by Solar PV module)
COPY solargis_country_pv_data.xlsx .

# Copy project files
COPY . .

# Cloud Run uses port 8080
ENV PORT=8080
EXPOSE 8080

# Start FastAPI app
CMD ["sh", "-c", "cd /app && uvicorn report_api:app --host 0.0.0.0 --port ${PORT}"]
