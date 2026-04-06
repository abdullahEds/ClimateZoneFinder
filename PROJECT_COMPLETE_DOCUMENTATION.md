# Climate Zone Finder - Complete Project Documentation

## 1. Project Overview

Climate Zone Finder is a Streamlit-based climate analytics application for processing EPW (EnergyPlus Weather) files and generating climate intelligence for building design workflows.

The application provides:
- Temperature analysis (annual, monthly, diurnal, comfort, energy metrics)
- Humidity analysis (annual, monthly, diurnal, comfort, risk metrics)
- Sun path and solar-shading analysis (including orientation-specific shading logic)
- Wind analysis (wind rose, speed/direction heatmaps, distribution, climate bubble plot)
- PowerPoint report generation for both general climate and shading-focused studies

## 2. Technology Stack

- Python
- Streamlit
- Pandas
- NumPy
- Plotly
- Matplotlib
- pvlib
- python-pptx
- reportlab (present in project dependencies)

Main dependency list is in `requirements.txt`.

## 3. Repository Structure

Top-level files and folders:
- `app.py`: home/landing UI
- `analysis.py`: legacy/placeholder analysis page ("Coming Soon")
- `pages/analysis.py`: main production analysis orchestrator
- `pages/modules/epw_parser.py`: EPW parsing and metadata extraction
- `pages/modules/dbt_module.py`: dry bulb temperature analysis module
- `pages/modules/humidity_module.py`: humidity analysis module
- `pages/modules/wind_module.py`: wind analysis module
- `pages/modules/sun_path.py`: sun path and shading UI logic
- `pages/modules/shading_helpers.py`: shading calculations and helper utilities
- `pages/modules/ppt_report.py`: PowerPoint report generation
- `images/`: logos/footer assets used in UI
- `SUNPATH_GUIDE.md`: user-level sun path usage notes
- `SUNPATH_IMPLEMENTATION.md`: implementation notes for sun path
- `COMPLETION_SUMMARY.md`: sun path completion summary

## 4. Application Architecture

The production app follows a thin-orchestrator architecture:

1. `app.py` serves as landing page with branding and navigation.
2. User clicks `Analysis` and is routed to `pages/analysis.py`.
3. `pages/analysis.py` controls:
- file upload
- global filters and control panel
- module routing
- report download buttons
4. Domain-specific logic is delegated to modules in `pages/modules/`.

This keeps page-level code focused on orchestration, while modules encapsulate analytics and visualization logic.

## 5. End-to-End User Workflow

### 5.1 Launch Workflow

1. Activate environment (`.venv` recommended).
2. Run Streamlit:
   ```powershell
   python -m streamlit run app.py
   ```
3. Open browser app.
4. Click `Analysis` to enter dashboard.

### 5.2 Analysis Workflow

1. Upload an EPW file (`.epw`).
2. Choose one module from selector:
- Temperature
- Humidity
- Sun Path
- Wind
3. Configure left-panel controls (hour range, month range, thresholds, orientation options depending on module).
4. Review charts and KPI cards in right panel.
5. Export report:
- General climate report (`generate_pptx_report`)
- Shading report (`generate_shading_pptx_report`) when Sun Path is in `Shading` mode

## 6. Core Data Flow

### 6.1 EPW Parse Flow

Source: `pages/modules/epw_parser.py`

`parse_epw(epw_text)` workflow:
1. Split and sanitize EPW text lines.
2. Read metadata from EPW LOCATION header:
- city
- location token
- latitude
- longitude
- timezone (mapped by `convert_epw_timezone`)
3. Detect first weather-data row (first token is 4-digit year).
4. Parse rows into DataFrame.
5. Map EPW fields into typed columns:
- datetime components
- dry bulb temperature
- relative humidity
- DNI, DHI, GHI
- wind direction and wind speed
6. Convert EPW hour convention 1-24 to 0-23.
7. Build `datetime` column and drop invalid rows.
8. Return tuple:
- climate DataFrame
- metadata dictionary

Returned DataFrame columns:
- `datetime`
- `dry_bulb_temperature`
- `relative_humidity`
- `direct_normal_irradiance`
- `diffuse_horizontal_irradiance`
- `global_horizontal_irradiance`
- `wind_direction`
- `wind_speed`
- `hour`

### 6.2 Orchestrator Flow

Source: `pages/analysis.py`

After EPW parse:
1. Adds derived fields: `doy`, `day`, `month`, `month_name`.
2. Builds control panel state:
- module options
- date range (month start/end)
- hour range
- wind options
- shading thresholds and location inputs
3. Generates download button payloads:
- climate report
- shading report
4. Routes right panel rendering by module.

## 7. Module Workflows

## 7.1 Temperature Module Workflow

Source: `pages/modules/dbt_module.py`

Entry: `render(df, daily_stats, active_tab, start_date, end_date, start_hour, end_hour)`

Tabs and behavior:
1. Annual Trend
- plots min/max daily temperature band
- overlays ASHRAE adaptive comfort bands (80% and 90%)
- highlights selected date period
- displays KPI cards (min/max/avg/diurnal range, cooling/heating metrics)

2. Monthly Trend
- groups by month and shows min/avg/max bars/lines
- highlights selected month range
- includes monthly summary table

3. Diurnal Profile
- groups by hour and month, then hourly aggregate
- shows hourly min/max range and average profile
- applies selected hour-range highlight

4. Comfort Analysis
- uses adaptive comfort output from `calculate_ashrae_comfort`
- compares daily temperature range vs ASHRAE 90% acceptability band

5. Energy Metrics
- computes HDD18 and CDD24 annual and filtered-period values
- shows monthly degree-day distribution

Key helper:
- `calculate_ashrae_comfort(df)` computes rolling adaptive comfort bands from daily average temperature.

## 7.2 Humidity Module Workflow

Source: `pages/modules/humidity_module.py`

Entry: `render(df, daily_stats, active_tab, start_date, end_date, start_hour, end_hour)`

Tabs and behavior:
1. Annual Trend
- comfort band plotting (30-65% RH)
- daily RH min/max range and average curve
- KPI cards for comfort %, high RH, condensation risk, low RH, etc.

2. Monthly Trend
- monthly min/max/avg RH with comfort overlay
- selected-period highlighting
- monthly summary table

3. Diurnal Profile
- hourly RH profile, range, and comfort overlay
- selected hour window emphasis

4. Comfort Analysis
- comfort band focus for 40-60% RH
- daily RH range and average in selected period

5. Energy Metrics (humidity risk)
- monthly and period risk counters:
- high RH hours (>60%)
- condensation risk hours (>75%)
- low RH (<30%)

## 7.3 Wind Module Workflow

Source: `pages/modules/wind_module.py`

Entry: `render_wind_analysis(epw_df, months, n_sectors, exclude_calm)`

Workflow:
1. Validate required columns: `wind_speed`, `wind_direction`.
2. `prepare_wind_data(...)`:
- add time components
- sanitize values
- normalize direction to [0, 360)
- classify calm hours (`wind_speed < 0.5 m/s`)
- assign centered directional sectors
- bin speed values
3. `compute_wind_rose(...)`:
- compute sector x speed-bin frequencies
- optionally normalize excluding calm hours
4. Render full set of visuals:
- wind rose
- speed heatmap (day x hour)
- direction heatmap (vector/circular averaging)
- wind speed histogram
- climate bubble plot (temperature, RH, wind)
5. Show prevailing wind KPI cards:
- prevailing direction
- mean speed
- max speed
- calm %
- strongest direction

Special algorithm details:
- Circular averaging with vector decomposition avoids 0/360 discontinuity artifacts.
- Sector assignment is centered to keep cardinal direction bins intuitive.

## 7.4 Sun Path Module Workflow

Source: `pages/modules/sun_path.py`

Entry: `render_sun_path_section(df, metadata)`

Chart types:
- Sun Path
- Dry Bulb Temperature (colored sun path)
- Direct Normal Radiation
- Global Horizontal Radiation
- Shading

Core chart flow (`plot_sun_path`):
1. Resolve lat/lon/timezone from EPW metadata.
2. Build timezone-aware annual timestamp series.
3. Compute solar positions with `pvlib.solarposition.get_solarposition`.
4. Remove nighttime points (`apparent_elevation <= 0`).
5. Merge solar geometry with EPW weather columns.
6. Compute derived GHI (if needed in chart logic).
7. Build Plotly polar chart:
- analemma traces per hour
- selected color encoding by chart type
- seasonal key curves (Mar 21, Jun 21, Dec 21)
- compass orientation and radial annotations
8. Return shading metrics when in `Shading` mode.

Shading mode extensions:
1. KPI metrics:
- total sunshine hours
- required shading hours
2. Thermal and radiation matrix (`build_thermal_matrix`)
3. Orientation shading analysis table (`build_orientation_table`)
4. Shading mask mini diagrams per facade orientation (`make_shading_mask_chart`)

## 8. Shading Computation Workflow

Source: `pages/modules/shading_helpers.py`

Main calculation chain:
1. `get_overheating_hours` filters rows where:
- dry bulb temperature > threshold
- GHI > threshold
2. `compute_solar_angles` computes solar altitude/azimuth for filtered timestamps.
3. `compute_shading_geometry` computes for each facade orientation:
- relative azimuth
- VSA (Vertical Shadow Angle)
- HSA (Horizontal Shadow Angle)
- whether rays hit facade (`abs(relative_azimuth) < 90`)
4. `build_orientation_table` returns facade-level metrics:
- rays hitting
- min VSA
- max abs(HSA)
- D/H overhang ratio
- D/W fin ratio
- protection percentage at design cutoff angle
5. `make_shading_mask_chart` generates compact polar diagram for each orientation.

## 9. Report Generation Workflows

Source: `pages/modules/ppt_report.py`

## 9.1 General Climate Report (`generate_pptx_report`)

Inputs:
- DataFrame
- selected dates and hour range
- selected module name
- metadata

Output:
- in-memory `.pptx` bytes stream

Slide workflow:
1. Initialize template deck (if available) or fallback deck.
2. Cover slide.
3. Temperature slides:
- annual trend
- monthly summary
4. Humidity slides:
- annual trend
- monthly summary
5. Sun path slide.
6. Shading strategy narrative slide.
7. Annexure slide (about/disclaimer/acknowledgement).

## 9.2 Shading Report (`generate_shading_pptx_report`)

Inputs:
- DataFrame and metadata
- thresholds
- optional lat/lon/tz override
- design cutoff angle

Output:
- in-memory `.pptx` bytes stream

Slide workflow:
1. Cover slide.
2. Thermal and radiation matrix slide.
3. Sun path (shading mode) slide.
4. Orientation shading analysis slide.
5. Shading mask diagrams slide.
6. Annexure slide.

## 10. UI State and Session Behavior

Managed in Streamlit session state (`st.session_state`):
- active analysis tab
- selected module
- month range selections
- hour range
- sun chart type
- shading thresholds
- shading location and cutoff values
- wind options (`n_sectors`, exclude calm)

This state supports interactivity without requiring data reload per control change.

## 11. Error Handling Strategy

Implemented patterns:
- parse-level hard stop with user-facing error if EPW is invalid
- module-level warnings for missing/empty datasets
- report generation wrapped with exception handling and on-screen messages
- graceful fallback for timezone parsing and optional metadata fields

## 12. Setup and Run Instructions

1. Create/activate virtual environment.
2. Install dependencies:
   ```powershell
   pip install -r requirements.txt
   ```
3. Run application:
   ```powershell
   python -m streamlit run app.py
   ```
4. In browser:
- click `Analysis`
- upload `.epw`
- choose module
- adjust controls
- export report if needed

## 13. Known Conventions and Assumptions

- EPW headers follow standard LOCATION field ordering.
- EPW hour values are hour-ending (1-24) and converted to 0-23.
- Time-series normalization in sun-path and shading logic maps timestamps to year 2020 for consistent annual plotting.
- Calm wind threshold is 0.5 m/s.
- Shading default thresholds are:
- temperature: 28.0 C
- GHI: 315 W/m2
- design cutoff: 45 degrees

## 14. Maintenance Guide

When adding new modules:
1. Implement module in `pages/modules/`.
2. Add selector option in `pages/analysis.py`.
3. Add left-panel controls for module parameters.
4. Route right-panel rendering.
5. Add report integration in `ppt_report.py` if export required.
6. Update this document with workflow and formulas.

When modifying EPW schema usage:
1. Update `epw_parser.py` mappings.
2. Validate downstream modules requiring new/changed columns.
3. Update report generators and this document.

## 15. Quick Workflow Reference

A. Standard climate review
1. Upload EPW
2. Choose Temperature or Humidity
3. Set month/hour filters
4. Review tabs and KPIs
5. Download climate report

B. Solar/shading review
1. Upload EPW
2. Choose Sun Path
3. Switch chart type to `Shading`
4. Tune threshold and cutoff parameters
5. Review matrix, orientation table, mask diagrams
6. Download shading report

C. Wind diagnostics
1. Upload EPW
2. Choose Wind
3. Set sectors and calm normalization preference
4. Review rose, heatmaps, histogram, bubble chart
5. Use prevailing stats for design direction insights

## 16. File Reference Index

- `app.py`
- `analysis.py`
- `pages/analysis.py`
- `pages/modules/epw_parser.py`
- `pages/modules/dbt_module.py`
- `pages/modules/humidity_module.py`
- `pages/modules/wind_module.py`
- `pages/modules/sun_path.py`
- `pages/modules/shading_helpers.py`
- `pages/modules/ppt_report.py`
- `requirements.txt`
- `SUNPATH_GUIDE.md`
- `SUNPATH_IMPLEMENTATION.md`
- `COMPLETION_SUMMARY.md`

---

Document version: 1.0  
Generated on: March 17, 2026
