# Sun Path Diagram - Quick Reference

## 🚀 Getting Started

### To use the Sun Path feature in your dashboard:

1. **Upload an EPW file** using the "📤 Upload EPW File" section
2. **Select "Sun Path"** from the "⚙️ Parameter" dropdown
3. **View the diagram** - automatically generated using pvlib solar calculations

---

## 📊 Diagram Components

### What Each Part Represents

| Component | Meaning |
|-----------|---------|
| **Center (0°)** | Zenith (sun directly overhead) |
| **Edge (90°)** | Horizon (sun at horizon) |
| **Colored dots** | Sun position for each hour throughout the year |
| **Color gradient** | Day of year (blue=early year, red=mid-year, blue again=late year) |
| **Hour labels** | Solar time (00-23 hours) |
| **Orange curve** | March 21 (Spring Equinox) |
| **Red curve** | June 21 (Summer Solstice - highest sun) |
| **Blue curve** | December 21 (Winter Solstice - lowest sun) |

---

## 🧮 Data Source

- **Location**: Automatically extracted from EPW file header
  - Latitude (decimal degrees)
  - Longitude (decimal degrees)
  - Timezone (UTC offset)
- **Calculations**: pvlib solar position model
- **Data**: Full year of hourly data from your EPW file

---

## ⚙️ Technical Implementation

### Files Modified
- `pages/analysis.py` - Main dashboard application
- `requirements.txt` - Dependencies (pvlib added)

### Key Functions
- `parse_epw()` - Parses EPW and extracts metadata
- `plot_sun_path()` - Generates the diagram

### Dependencies
- `pvlib` - Solar position calculations
- `matplotlib` - Polar plot rendering
- `pandas` - Data handling
- `streamlit` - Display

---

## 🎨 Interpretation Guide

### Sun Path Patterns

**Equatorial locations** (lat = 0°):
- Curves are nearly symmetric
- Sun reaches high zenith angles

**Northern Hemisphere** (lat > 0°):
- December curve (blue) is lowest
- June curve (red) is highest

**Southern Hemisphere** (lat < 0°):
- June curve (blue) is lowest
- December curve (red) is highest

**High latitude** (|lat| > 60°):
- Very narrow sun paths
- Sun never reaches high elevations in winter

---

## 🔍 How To Read The Diagram

1. **Check a specific date**: Find the analog curve (Mar 21, Jun 21, or Dec 21)
2. **Check sunrise/sunset**: Look where the curve meets the 90° edge
3. **Check sun height**: Count distance from edge (90°) toward center (0°)
   - Closer to center = higher sun in sky
   - Closer to edge = lower sun in sky (near horizon)
4. **Check solar azimuth**: Look at position around the circle
   - Top = North
   - Right = East
   - Bottom = South
   - Left = West

---

## 💡 Common Use Cases

### Building Design
- Determine window orientations for passive solar heating
- Plan shading strategies for cooling

### Solar Panel Installation
- Optimize panel orientation and tilt angle
- Visualize seasonal sun path for obstructions

### Climate Analysis
- Understand seasonal solar radiation variation
- Assess heating/cooling seasonal patterns

### Daylighting Design
- Plan natural light strategies
- Assess solar gains during different seasons

---

## ⚠️ Important Notes

- **No re-processing**: Uses same EPW data loaded for other analyses
- **No hardcoding**: Location data from file header, not hardcoded
- **Timezone aware**: Properly handles timezone conversions
- **Full year**: Shows annual variation automatically
- **No filters**: All daylight hours are displayed

---

## 🐛 Troubleshooting

| Issue | Solution |
|-------|----------|
| "Location information not found" | EPW file header may be malformed. Check first line has lat/lon. |
| "No daytime solar positions" | Rare: occurs only in extreme polar regions during winter. |
| Diagram not rendering | Check pvlib is installed: `pip install pvlib` |
| Wrong coordinates | Verify EPW file is for correct location |

---

## 📚 Further Learning

- **About Sun Paths**: https://en.wikipedia.org/wiki/Analemma
- **pvlib Documentation**: https://pvlib-python.readthedocs.io/
- **EPW Format**: Weather file standard for building simulation

---

**Version**: 1.0 (Production Ready)
**Last Updated**: March 2026
