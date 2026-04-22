# Use Python 3.11 slim image
FROM python:3.11-slim

# Set working directory
WORKDIR /app

# Copy requirements and install
COPY requirements_api.txt .
RUN pip install --no-cache-dir -r requirements_api.txt

# Copy the entire project
COPY . .

# Expose port 8080 (Cloud Run expects this)
EXPOSE 8080

# Run the app with Uvicorn
CMD ["uvicorn", "report_api:app", "--host", "0.0.0.0", "--port", "8080"]