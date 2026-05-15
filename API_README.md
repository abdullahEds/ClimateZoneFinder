# Climate Zone Finder - PPT Report API

This is a FastAPI-based REST API for generating PowerPoint climate analysis reports from EPW (EnergyPlus Weather) files without requiring UI interaction.

## Installation

### Prerequisites
- Python 3.8+
- Virtual environment (recommended)

### Setup

1. **Install dependencies:**
```bash
pip install -r requirements_api.txt
```

2. **Start the API server:**
```bash
python report_api.py
```

The server will start on `http://localhost:8000`

## API Endpoints

### 1. Health Check
**GET** `/api/health`

Check if the API is running.

**Response:**
```json
{
  "status": "ok",
  "message": "Climate Zone Finder PPT Report API is running"
}
```

**Example:**
```bash
curl http://localhost:8000/api/health
```

---

### 2. Climate Analysis Report
**POST** `/api/reports/climate-analysis`

Generate a comprehensive climate analysis PowerPoint report from an EPW file.

**Parameters:**
- `file` (required): EPW weather file
- `start_date` (optional): Start date in format `YYYY-MM-DD` (default: first day in file)
- `end_date` (optional): End date in format `YYYY-MM-DD` (default: last day in file)
- `start_hour` (optional): Start hour 0-23 (default: 0)
- `end_hour` (optional): End hour 0-23 (default: 23)

**Response:** PowerPoint file (`.pptx`)

**Examples:**

Basic usage:
```bash
curl -X POST "http://localhost:8000/api/reports/climate-analysis" \
  -F "file=@weather.epw" \
  -o climate_report.pptx
```

With date range and hour filter:
```bash
curl -X POST "http://localhost:8000/api/reports/climate-analysis" \
  -F "file=@weather.epw" \
  -d "start_date=2024-06-01" \
  -d "end_date=2024-08-31" \
  -d "start_hour=9" \
  -d "end_hour=18" \
  -o climate_report.pptx
```

---

### 3. Shading Analysis Report
**POST** `/api/reports/shading-analysis`

Generate a detailed shading analysis PowerPoint report from an EPW file.

**Parameters:**
- `file` (required): EPW weather file
- `temp_threshold` (optional): Temperature threshold in °C (default: 28.0)
- `rad_threshold` (optional): Radiation threshold in W/m² (default: 315.0)
- `design_cutoff_angle` (optional): Design cutoff angle in degrees (default: 45.0)

**Response:** PowerPoint file (`.pptx`)

**Examples:**

Basic usage:
```bash
curl -X POST "http://localhost:8000/api/reports/shading-analysis" \
  -F "file=@weather.epw" \
  -o shading_report.pptx
```

With custom thresholds:
```bash
curl -X POST "http://localhost:8000/api/reports/shading-analysis" \
  -F "file=@weather.epw" \
  -d "temp_threshold=30" \
  -d "rad_threshold=400" \
  -d "design_cutoff_angle=50" \
  -o shading_report.pptx
```

---

## Usage Examples by Language

### JavaScript (Node.js)

```javascript
const FormData = require('form-data');
const fs = require('fs');
const axios = require('axios');

async function generateClimateReport() {
  const form = new FormData();
  form.append('file', fs.createReadStream('weather.epw'));
  
  try {
    const response = await axios.post(
      'http://localhost:8000/api/reports/climate-analysis',
      form,
      {
        headers: form.getHeaders(),
        responseType: 'arraybuffer'
      }
    );
    
    fs.writeFileSync('climate_report.pptx', response.data);
    console.log('Climate Analysis Report generated!');
  } catch (error) {
    console.error('Error:', error.response?.data || error.message);
  }
}

generateClimateReport();
```

### PHP

```php
<?php
// Climate Analysis Report
$file = new CURLFile('weather.epw');
$post = array(
    'file' => $file,
    'start_date' => '2024-06-01',
    'end_date' => '2024-08-31'
);

$ch = curl_init();
curl_setopt($ch, CURLOPT_URL, 'http://localhost:8000/api/reports/climate-analysis');
curl_setopt($ch, CURLOPT_POST, 1);
curl_setopt($ch, CURLOPT_POSTFIELDS, $post);
curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);

$response = curl_exec($ch);
curl_close($ch);

file_put_contents('climate_report.pptx', $response);
echo "Climate Analysis Report generated!";
?>
```

### Java

```java
import java.io.*;
import java.net.http.*;
import java.nio.file.Files;
import java.nio.file.Paths;

public class ClimateReportClient {
    public static void main(String[] args) throws Exception {
        HttpClient client = HttpClient.newHttpClient();
        
        // Read EPW file
        byte[] fileBytes = Files.readAllBytes(Paths.get("weather.epw"));
        
        HttpRequest request = HttpRequest.newBuilder()
            .uri(new URI("http://localhost:8000/api/reports/climate-analysis"))
            .POST(HttpRequest.BodyPublishers.ofByteArray(fileBytes))
            .header("Content-Type", "multipart/form-data")
            .build();
        
        HttpResponse<byte[]> response = client.send(request, 
            HttpResponse.BodyHandlers.ofByteArray());
        
        // Save report
        Files.write(Paths.get("climate_report.pptx"), response.body());
        System.out.println("Report generated!");
    }
}
```

### Python

```python
import requests

def generate_climate_report(epw_file_path):
    """Generate climate analysis report from EPW file."""
    
    with open(epw_file_path, 'rb') as f:
        files = {'file': f}
        params = {
            'start_date': '2024-06-01',
            'end_date': '2024-08-31',
            'start_hour': 9,
            'end_hour': 18
        }
        
        response = requests.post(
            'http://localhost:8000/api/reports/climate-analysis',
            files=files,
            params=params
        )
    
    if response.status_code == 200:
        with open('climate_report.pptx', 'wb') as f:
            f.write(response.content)
        print('Report generated successfully!')
    else:
        print(f'Error: {response.text}')

generate_climate_report('weather.epw')
```

### cURL Examples

**Climate Analysis Report:**
```bash
# Basic
curl -X POST "http://localhost:8000/api/reports/climate-analysis" \
  -F "file=@weather.epw" \
  -o climate_report.pptx

# With date range and hourly filter
curl -X POST "http://localhost:8000/api/reports/climate-analysis" \
  -F "file=@weather.epw" \
  -G \
  -d "start_date=2024-06-01" \
  -d "end_date=2024-08-31" \
  -d "start_hour=9" \
  -d "end_hour=18" \
  -o climate_report.pptx
```

**Shading Analysis Report:**
```bash
# Basic
curl -X POST "http://localhost:8000/api/reports/shading-analysis" \
  -F "file=@weather.epw" \
  -o shading_report.pptx

# With custom thresholds
curl -X POST "http://localhost:8000/api/reports/shading-analysis" \
  -F "file=@weather.epw" \
  -G \
  -d "temp_threshold=30" \
  -d "rad_threshold=400" \
  -d "design_cutoff_angle=50" \
  -o shading_report.pptx
```

---

## Report Contents

### Climate Analysis Report Includes:
- Dry Bulb Temperature trends (annual and monthly)
- Relative Humidity analysis
- Sun Path diagram
- Shading strategy recommendations
- Design considerations
- Project information table

### Shading Analysis Report Includes:
- Thermal & Radiation matrix
- Sun Path diagram with shading analysis
- Orientation-specific shading analysis
- Shading mask diagrams for all orientations
- Design recommendations

---

## Error Handling

The API returns appropriate HTTP status codes:

- `200`: Success - Report generated
- `400`: Bad Request - Invalid parameters or malformed file
- `500`: Server Error - Report generation failed

**Error Response Example:**
```json
{
  "detail": "Error parsing EPW file: ..."
}
```

---

## Configuration

To modify the API settings, edit `report_api.py`:

```python
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000, reload=True)
```

- `host`: API listening address (default: `0.0.0.0`)
- `port`: API listening port (default: `8000`)
- `reload`: Auto-reload on code changes (default: `True`)

---

## Production Deployment

For production deployment, use Gunicorn:

```bash
pip install gunicorn
gunicorn -w 4 -b 0.0.0.0:8000 report_api:app
```

---

## Troubleshooting

### Port Already in Use
```bash
# Change port in the script or kill the process using port 8000
lsof -ti:8000 | xargs kill -9  # Linux/Mac
netstat -ano | findstr :8000    # Windows
```

### Module Import Error
Ensure the working directory is the project root:
```bash
cd Path/To/ClimateZoneFinder
python report_api.py
```

### EPW File Parsing Error
Ensure your EPW file is valid and uses the standard EPW format. The file should have:
- Header lines with location metadata
- Data starting from line 9
- 8760 hourly records

---

## License & Attribution

This API uses the Climate Zone Finder project. See main project documentation for license information.

---

## Support

For issues or feature requests, contact: info@edsglobal.com
