# Psychrometric Chart Addition to thermal_comfort_ppt.py

## Summary of Changes

The psychrometric chart has been successfully added to the **thermal_comfort_ppt.py** module for the PowerPoint report generation.

## Changes Made:

### 1. New Function: `plot_psychrometric_chart()`
- **Location**: Lines 293-374 in `thermal_comfort_ppt.py`
- **Purpose**: Generates a matplotlib-based psychrometric chart visualization
- **Features**:
  - Plots hourly climate data points colored by dry bulb temperature
  - Shows constant relative humidity curves (10-90% RH) as background reference lines
  - Displays the ASHRAE 55 static comfort zone
  - Overlays design strategy zones:
    - Natural ventilation zone (green)
    - Evaporative cooling zone (teal)
    - Heating zone (orange)
  - Includes color bar showing temperature scale
  - High-quality matplotlib visualization suitable for PowerPoint

### 2. New PPT Slide: Psychrometric Chart
- **Location**: Lines 585-602 in `thermal_comfort_ppt.py`
- **Title**: "Psychrometric Chart – Climate Data"
- **Integration**: Added as the second visualization slide (after Comfort Heatmap, before Design Strategies)
- **Error Handling**: Includes try-except block for robust error reporting

### 3. Updated Cover Slide
- **Location**: Line 556
- **Change**: Updated section listing to include "Psychrometric Chart" in the table of contents
- **Before**: "Sections: Comfort Heatmap | Design Strategies | Degree Hours | Adaptive Comfort | Performance Summary"
- **After**: "Sections: Comfort Heatmap | Psychrometric Chart | Design Strategies | Degree Hours | Adaptive Comfort | Performance Summary"

## Technical Details:

### Psychrometric Chart Calculation
- Uses Magnus-Tetens formula for saturation vapor pressure
- Converts relative humidity to humidity ratio (g/kg dry air)
- X-axis: Dry Bulb Temperature (°C), range: -5 to 50°C
- Y-axis: Humidity Ratio (g/kg), range: 0 to 28 g/kg

### Chart Elements
1. **Background RH curves**: 10%, 20%, 30%, ..., 90% RH lines in light gray
2. **ASHRAE 55 Comfort Zone**: Filled green region (DBT 20-26°C, RH 20-80%)
3. **Natural Ventilation Zone**: Lighter green overlay (DBT 22-32°C, RH 20-70%)
4. **Evaporative Cooling Zone**: Teal overlay (DBT 25-45°C, RH 0-35%)
5. **Heating Zone**: Light orange (Cold temperatures)
6. **Hourly Data Points**: Colored scatter plot (blue = cold, red = hot)

## Files Modified:
- `pages/modules/thermal_comfort_ppt.py`

## Dependencies:
All required dependencies are already in `requirements.txt`:
- numpy
- pandas
- matplotlib
- python-pptx

## Testing:
The changes maintain backward compatibility with existing code. The psychrometric chart function:
- Handles missing data gracefully
- Validates input data with pd.to_numeric(..., errors="coerce")
- Produces figures at high DPI (130) suitable for printing
- Automatically cleans up temporary image files
