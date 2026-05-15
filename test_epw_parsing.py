"""Test the fixed EPW parsing."""

import requests
import time

print("Testing API with EPW file...")
print("=" * 60)

# Wait a moment for API to be ready
time.sleep(2)

# Test health endpoint
print("\n1. Health Check:")
try:
    response = requests.get("http://localhost:8001/api/health", timeout=5)
    print(f"✓ Status: {response.status_code}")
    print(f"  Response: {response.json()}")
except Exception as e:
    print(f"✗ Error: {e}")

# Test with a real EPW file (if available)
print("\n2. Testing Climate Analysis Report Generation:")
epw_file = "IND_DL_New.Delhi-Safdarjung.AP.421820_ISHRAE2014.epw"

try:
    with open(epw_file, 'rb') as f:
        files = {'file': f}
        response = requests.post(
            'http://localhost:8001/api/reports/climate-analysis',
            files=files,
            timeout=60
        )
    
    print(f"Status Code: {response.status_code}")
    
    if response.status_code == 200:
        print("✓ Report generated successfully!")
        # Save the file
        with open('test_climate_report.pptx', 'wb') as out:
            out.write(response.content)
        print(f"  File saved: test_climate_report.pptx ({len(response.content)} bytes)")
    else:
        print(f"✗ Error: {response.text}")
        
except FileNotFoundError:
    print(f"✗ EPW file not found: {epw_file}")
except Exception as e:
    print(f"✗ Error: {e}")

print("\n" + "=" * 60)
print("Test complete!")
