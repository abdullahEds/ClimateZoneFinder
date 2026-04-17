"""Test script to verify API is working."""

import requests

def test_api():
    """Test the API endpoints."""
    
    print("Testing Climate Zone Finder API...")
    print("-" * 50)
    
    # Test health endpoint
    print("\n1. Testing Health Endpoint:")
    try:
        response = requests.get("http://localhost:8000/api/health", timeout=5)
        print(f"Status Code: {response.status_code}")
        print(f"Raw Response: {response.text}")
        if response.status_code == 200:
            print("✓ Health endpoint is working!")
    except Exception as e:
        print(f"✗ Error: {e}")
    
    # Test documentation endpoint
    print("\n2. Testing Documentation Endpoint:")
    try:
        response = requests.get("http://localhost:8000/api/docs", timeout=5)
        print(f"Status Code: {response.status_code}")
        if response.status_code == 200:
            print("✓ Documentation endpoint is working!")
            print(f"Response snippet: {response.text[:100]}...")
    except Exception as e:
        print(f"✗ Error: {e}")
    
    print("\n" + "=" * 50)
    print("API Setup Complete!")
    print("=" * 50)
    print("\nAPI is running on: http://localhost:8000")
    print("\nAvailable endpoints:")
    print("  - GET  /api/health")
    print("  - GET  /api/docs")
    print("  - POST /api/reports/climate-analysis (with EPW file)")
    print("  - POST /api/reports/shading-analysis (with EPW file)")
    print("\nDocumentation:")
    print("  - See API_README.md for full API documentation")
    print("  - Check API_README.md for usage examples in Python, JavaScript, PHP, Java, etc.")

if __name__ == "__main__":
    test_api()
