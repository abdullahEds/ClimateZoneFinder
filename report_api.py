"""FastAPI server for generating PowerPoint climate analysis reports from EPW files."""

from fastapi import FastAPI, UploadFile, File, HTTPException, Query, Form
from fastapi.responses import StreamingResponse
import pandas as pd
import io
import os
from datetime import datetime
from typing import Optional
import sys
from fastapi.middleware.cors import CORSMiddleware



# Add the pages directory and modules path for imports
current_dir = os.path.dirname(os.path.abspath(__file__))
pages_dir = os.path.join(current_dir, 'pages')
modules_dir = os.path.join(pages_dir, 'modules')

sys.path.insert(0, pages_dir)
sys.path.insert(0, modules_dir)

from pages.modules.ppt_report import generate_pptx_report, generate_shading_pptx_report
from pages.modules.combined_report import generate_combined_pptx_report

app = FastAPI(
    title="Climate Zone Finder - PPT Report API",
    description="REST API for generating PowerPoint climate analysis reports from EPW files",
    version="1.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Allow all origins for production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def parse_epw_file(file_content: bytes) -> tuple[pd.DataFrame, dict]:
    """Parse EPW file and extract weather data and metadata."""
    try:
        lines = file_content.decode('utf-8').split('\n')
        
        # Extract header information
        header = lines[0].split(',')
        metadata = {
            "city": header[1].strip() if len(header) > 1 else "Unknown",
            "state": header[2].strip() if len(header) > 2 else "",
            "country": header[3].strip() if len(header) > 3 else "",
            "latitude": float(header[6]) if len(header) > 6 else 0.0,
            "longitude": float(header[7]) if len(header) > 7 else 0.0,
            "timezone": float(header[8]) if len(header) > 8 else 0,
            "elevation": float(header[9]) if len(header) > 9 else 0,
        }
        
        # Parse weather data (skip header lines)
        data_lines = [line for line in lines[8:] if line.strip()]
        
        # Standard EPW columns (34 columns)
        standard_columns = [
            'Year', 'Month', 'Day', 'Hour', 'Minute', 'Data Source Flag',
            'dry_bulb_temperature', 'dew_point_temperature', 'relative_humidity',
            'atmospheric_pressure', 'extraterrestrial_horizontal_radiation',
            'extraterrestrial_direct_normal_radiation', 'horizontal_infrared_radiation_intensity',
            'global_horizontal_irradiance', 'direct_normal_irradiance',
            'diffuse_horizontal_irradiance', 'global_horizontal_illuminance',
            'direct_normal_illuminance', 'diffuse_horizontal_illuminance',
            'zenith_luminance', 'wind_direction', 'wind_speed',
            'total_sky_cover', 'opaque_sky_cover', 'visibility',
            'ceiling_height', 'present_weather_observation', 'present_weather_codes',
            'precipitable_water', 'aerosol_optical_depth', 'snow_depth',
            'days_since_last_snowfall', 'albedo', 'liquid_precipitation_depth',
            'liquid_precipitation_quantity'
        ]
        
        # Detect actual number of columns in the first data line
        if data_lines:
            sample_line = data_lines[0].split(',')
            actual_columns = len(sample_line)
        else:
            actual_columns = len(standard_columns)
        
        # Use standard columns, or add extra columns if needed
        if actual_columns > len(standard_columns):
            columns = standard_columns + [f'extra_col_{i}' for i in range(actual_columns - len(standard_columns))]
        else:
            columns = standard_columns[:actual_columns]
        
        data = []
        for line in data_lines:
            values = line.split(',')
            # Ensure we have the right number of columns
            if len(values) >= len(columns):
                row_data = [v.strip() for v in values[:len(columns)]]
                data.append(row_data)
            elif len(values) >= 34:  # Minimum requirement
                row_data = [v.strip() for v in values[:34]]
                data.append(row_data)
        
        if not data:
            raise ValueError("No valid data rows found in EPW file")
        
        df = pd.DataFrame(data, columns=columns[:len(data[0])])
        
        # Convert to numeric types
        numeric_cols = [
            'Year', 'Month', 'Day', 'Hour', 'Minute',
            'dry_bulb_temperature', 'dew_point_temperature', 'relative_humidity',
            'atmospheric_pressure', 'global_horizontal_irradiance',
            'direct_normal_irradiance', 'diffuse_horizontal_irradiance',
            'wind_direction', 'wind_speed'
        ]
        
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')
        
        # Create datetime column
        df['datetime'] = pd.to_datetime(
            df[['Year', 'Month', 'Day', 'Hour', 'Minute']],
            errors='coerce'
        )
        
        # Add derived columns
        df['month'] = df['Month'].astype(int)
        df['hour'] = df['Hour'].astype(int)
        df['doy'] = df['datetime'].dt.dayofyear
        
        return df, metadata
        
    except Exception as e:
        raise ValueError(f"Error parsing EPW file: {str(e)}")


@app.get("/api/health")
def health_check():
    """Health check endpoint."""
    return {"status": "ok", "message": "Climate Zone Finder PPT Report API is running"}


@app.post("/api/reports/climate-analysis")
async def generate_climate_report(
    file: UploadFile = File(...),
    start_date: Optional[str] = Query(None, description="Start date (YYYY-MM-DD)"),
    end_date: Optional[str] = Query(None, description="End date (YYYY-MM-DD)"),
    start_hour: int = Query(0, description="Start hour (0-23)"),
    end_hour: int = Query(23, description="End hour (0-23)"),
):
    """
    Generate a climate analysis PowerPoint report from an EPW file.
    
    Parameters:
    - file: EPW weather file (required)
    - start_date: Start analysis date (default: first day in file)
    - end_date: End analysis date (default: last day in file)
    - start_hour: Start hour of day (default: 0)
    - end_hour: End hour of day (default: 23)
    
    Returns: PowerPoint report file
    """
    try:
        # Read uploaded file
        content = await file.read()
        
        # Parse EPW
        df, metadata = parse_epw_file(content)
        
        if df.empty:
            raise ValueError("EPW file is empty or could not be parsed")
        
        # Set date range
        if start_date is None:
            start_dt = df['datetime'].min().date()
        else:
            start_dt = datetime.strptime(start_date, "%Y-%m-%d").date()
        
        if end_date is None:
            end_dt = df['datetime'].max().date()
        else:
            end_dt = datetime.strptime(end_date, "%Y-%m-%d").date()
        
        # Generate report
        pptx_buffer = generate_pptx_report(
            df=df,
            start_date=start_dt,
            end_date=end_dt,
            start_hour=start_hour,
            end_hour=end_hour,
            selected_parameter="dry_bulb_temperature",
            metadata=metadata
        )
        
        city = metadata.get('city', 'Climate_Report')
        filename = f"Climate_Analysis_{city}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
        
        return StreamingResponse(
            iter([pptx_buffer.getvalue()]),
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
        
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Report generation failed: {str(e)}")


@app.post("/api/reports/shading-analysis")
async def generate_shading_report(
    file: UploadFile = File(...),
    temp_threshold: float = Query(28.0, description="Temperature threshold (°C)"),
    rad_threshold: float = Query(315.0, description="Radiation threshold (W/m²)"),
    design_cutoff_angle: float = Query(45.0, description="Design cutoff angle (°)"),
):
    """
    Generate a shading analysis PowerPoint report from an EPW file.
    
    Parameters:
    - file: EPW weather file (required)
    - temp_threshold: Temperature threshold for overheating (default: 28°C)
    - rad_threshold: Radiation threshold for overheating (default: 315 W/m²)
    - design_cutoff_angle: Vertical shading angle cutoff (default: 45°)
    
    Returns: PowerPoint report file
    """
    try:
        # Read uploaded file
        content = await file.read()
        
        # Parse EPW
        df, metadata = parse_epw_file(content)
        
        if df.empty:
            raise ValueError("EPW file is empty or could not be parsed")
        
        # Generate report
        pptx_buffer = generate_shading_pptx_report(
            df=df,
            metadata=metadata,
            temp_threshold=temp_threshold,
            rad_threshold=rad_threshold,
            lat=metadata.get('latitude'),
            lon=metadata.get('longitude'),
            tz_str=str(metadata.get('timezone', 'UTC')),
            design_cutoff_angle=design_cutoff_angle
        )
        
        city = metadata.get('city', 'Shading_Report')
        filename = f"Shading_Analysis_{city}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
        
        return StreamingResponse(
            iter([pptx_buffer.getvalue()]),
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
        
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Report generation failed: {str(e)}")


@app.post("/api/reports/combined-analysis")
async def generate_combined_report(
    file: UploadFile = File(...),
    start_date: Optional[str] = Form(None, description="Start date (YYYY-MM-DD)"),
    end_date: Optional[str] = Form(None, description="End date (YYYY-MM-DD)"),
    start_hour: int = Form(0, description="Start hour (0-23), default: 0 (full day)"),
    end_hour: int = Form(23, description="End hour (0-23), default: 23 (full day)"),
    temp_threshold: float = Form(28.0, description="Temperature threshold (°C), default: 28.0"),
    rad_threshold: float = Form(315.0, description="Radiation threshold (W/m²), default: 315.0"),
    design_cutoff_angle: float = Form(45.0, description="Design cutoff angle (°), default: 45.0"),
):
    """
    Generate a combined Climate & Shading Analysis PowerPoint report from an EPW file.
    
    This endpoint generates a comprehensive report with:
    - Cover slide with location info
    - Assumptions & Analysis Parameters slide (customizable thresholds)
    - Climate Analysis: Dry Bulb Temperature, Relative Humidity, Sun Path
    - Shading Analysis: Thermal/Radiation Matrix, Sun Path Shading, Orientation Analysis, Shading Masks
    - Annexure with disclaimer and acknowledgements
    
    Parameters:
    - file: EPW weather file (required)
    - start_date: Start analysis date in YYYY-MM-DD format (optional, default: first day in file)
    - end_date: End analysis date in YYYY-MM-DD format (optional, default: last day in file)
    - start_hour: Start hour of analysis (0-23, default: 0 - full day)
    - end_hour: End hour of analysis (0-23, default: 23 - full day)
    - temp_threshold: Temperature threshold for overheating detection (°C, default: 28.0)
    - rad_threshold: Solar radiation threshold for shading analysis (W/m², default: 315.0)
    - design_cutoff_angle: Vertical design angle for shading calculations (°, default: 45.0)
    
    Returns: PowerPoint report file with combined climate and shading analysis
    """
    try:
        # Read uploaded file
        content = await file.read()
        
        # Parse EPW
        df, metadata = parse_epw_file(content)
        
        if df.empty:
            raise ValueError("EPW file is empty or could not be parsed")
        
        # Set date range - default to full year (Jan 1 - Dec 31)
        _year = df["datetime"].dt.year.iloc[0] if not df.empty else 2024
        
        if start_date is None:
            start_dt = pd.to_datetime(f"{_year}-01-01").date()
        else:
            start_dt = datetime.strptime(start_date, "%Y-%m-%d").date()
        
        if end_date is None:
            end_dt = pd.to_datetime(f"{_year}-12-31").date()
        else:
            end_dt = datetime.strptime(end_date, "%Y-%m-%d").date()
        
        # Generate combined report
        pptx_buffer = generate_combined_pptx_report(
            df=df,
            start_date=start_dt,
            end_date=end_dt,
            start_hour=start_hour,
            end_hour=end_hour,
            selected_parameter="combined",
            metadata=metadata,
            temp_threshold=temp_threshold,
            rad_threshold=rad_threshold,
            design_cutoff_angle=design_cutoff_angle,
        )
        
        city = metadata.get('city', 'Combined_Report')
        filename = f"Climate_Shading_Analysis_{city}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
        
        return StreamingResponse(
            iter([pptx_buffer.getvalue()]),
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
        
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Report generation failed: {str(e)}")


@app.get("/api/docs")
def api_documentation():
    """API documentation."""
    return {
        "title": "Climate Zone Finder - PPT Report API",
        "version": "1.0.0",
        "endpoints": {
            "health_check": {
                "method": "GET",
                "path": "/api/health",
                "description": "Check if API is running"
            },
            "climate_analysis_report": {
                "method": "POST",
                "path": "/api/reports/climate-analysis",
                "description": "Generate climate analysis PowerPoint report from EPW file",
                "parameters": {
                    "file": "EPW weather file (required)",
                    "start_date": "Start date in YYYY-MM-DD format (optional)",
                    "end_date": "End date in YYYY-MM-DD format (optional)",
                    "start_hour": "Start hour 0-23 (default: 0)",
                    "end_hour": "End hour 0-23 (default: 23)"
                }
            },
            "shading_analysis_report": {
                "method": "POST",
                "path": "/api/reports/shading-analysis",
                "description": "Generate shading analysis PowerPoint report from EPW file",
                "parameters": {
                    "file": "EPW weather file (required)",
                    "temp_threshold": "Temperature threshold in °C (default: 28.0)",
                    "rad_threshold": "Radiation threshold in W/m² (default: 315.0)",
                    "design_cutoff_angle": "Design cutoff angle in degrees (default: 45.0)"
                }
            },
            "combined_analysis_report": {
                "method": "POST",
                "path": "/api/reports/combined-analysis",
                "description": "Generate combined Climate & Shading Analysis PowerPoint report from EPW file (Comprehensive Report)",
                "parameters": {
                    "file": "EPW weather file (required)",
                    "start_date": "Start date in YYYY-MM-DD format (optional, default: first day in file)",
                    "end_date": "End date in YYYY-MM-DD format (optional, default: last day in file)",
                    "start_hour": "Start hour 0-23 (default: 0 - full day)",
                    "end_hour": "End hour 0-23 (default: 23 - full day)",
                    "temp_threshold": "Temperature threshold in °C for overheating (default: 28.0)",
                    "rad_threshold": "Radiation threshold in W/m² (default: 315.0)",
                    "design_cutoff_angle": "Design cutoff angle in degrees (default: 45.0)"
                }
            }
        },
        "examples": {
            "climate_analysis": "curl -X POST 'http://localhost:8001/api/reports/climate-analysis' -F 'file=@weather.epw' -o report.pptx",
            "shading_analysis": "curl -X POST 'http://localhost:8001/api/reports/shading-analysis' -F 'file=@weather.epw' -o shading_report.pptx",
            "combined_analysis_default": "curl -X POST 'http://localhost:8001/api/reports/combined-analysis' -F 'file=@weather.epw' -o combined_report.pptx",
            "combined_analysis_custom": "curl -X POST 'http://localhost:8001/api/reports/combined-analysis' -F 'file=@weather.epw' -G -d 'temp_threshold=26' -d 'rad_threshold=300' -d 'design_cutoff_angle=50' -o combined_report.pptx"
        }
    }


if __name__ == "__main__":
    import uvicorn
    print("Starting Climate Zone Finder PPT Report API...")
    print("Documentation available at: http://localhost:8001/api/docs")
    uvicorn.run(app, host="0.0.0.0", port=8001)
