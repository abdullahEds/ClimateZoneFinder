# Refactoring & Technical Debt Resolution Prompt
# Climate Zone Finder — EDS Global
# Target model: Claude Sonnet 4.6

---

## CONTEXT

You are working on the **Climate Zone Finder** project — a Python monorepo at the root of the connected folder. The project has two entry points:

1. **Streamlit app** (`app.py` + `pages/analysis.py`) — interactive browser UI
2. **FastAPI API** (`report_api.py`) — REST API for programmatic PPT report generation

Shared analysis logic lives in `pages/modules/` (one file per climate topic). PPT generators are also in `pages/modules/` (`ppt_report.py`, `combined_report.py`, `thermal_comfort_ppt.py`, `rainfall_ppt.py`).

Your task is to resolve **all** of the following issues in a single pass. Work through them in the order listed. After completing each item, verify it before moving to the next.

---

## ISSUES TO RESOLVE

---

### ISSUE 1 — Unified EPW Parser [HIGH PRIORITY]

**Problem:**  
There are two independent EPW parsers with slightly different logic:
- `pages/modules/epw_parser.py` — used by the Streamlit app. Parses to a lean DataFrame with columns: `datetime, dry_bulb_temperature, relative_humidity, direct_normal_irradiance, diffuse_horizontal_irradiance, global_horizontal_irradiance, wind_direction, wind_speed, hour`.
- `parse_epw_file()` defined inline in `report_api.py` (starting around line 43) — used by the API. Returns a richer DataFrame with `Year, Month, Day, Hour, Minute, atmospheric_pressure, dew_point_temperature, liquid_precipitation_depth, doy, month` etc., plus a metadata dict.

**Fix:**  
Extend `pages/modules/epw_parser.py` so its `parse_epw()` function returns a DataFrame that satisfies **both** consumers. Specifically:
- Keep all existing columns from `epw_parser.py`.
- Add the extra columns the API uses: `dew_point_temperature` (EPW col 7), `atmospheric_pressure` (EPW col 9), `liquid_precipitation_depth` (EPW col 33), `doy`, `month`, `Year`, `Month`, `Day`, `Minute`.
- Extend the metadata dict to also include `state`, `country`, `elevation` (already in EPW header).
- Delete the inline `parse_epw_file()` function from `report_api.py`.
- In `report_api.py`, import and call `parse_epw()` from `pages.modules.epw_parser` instead. Because the API receives bytes (not a string), add a thin wrapper at the call site: `df, metadata = parse_epw(file_content.decode('utf-8', errors='replace'))`.
- Ensure all existing call sites in `pages/analysis.py` that call `cached_parse_epw()` still work — `cached_parse_epw()` calls `parse_epw()`, so this should be transparent.

---

### ISSUE 2 — Fix Fragile sys.path Manipulation [HIGH PRIORITY]

**Problem:**  
`report_api.py` lines 15–21 manually insert `pages/` and `pages/modules/` into `sys.path` at runtime using `sys.path.insert(0, ...)`. This is fragile — it relies on the working directory and breaks if the server is started from a different location.

**Fix:**
1. Verify that `pages/__init__.py` and `pages/modules/__init__.py` already exist (they do — confirmed). If either is missing, create it as an empty file.
2. Remove the `sys.path.insert` block (lines 15–21) from `report_api.py`.
3. Change all imports in `report_api.py` from `from pages.modules.X import Y` to `from pages.modules.X import Y` — these should already work once `sys.path` manipulation is removed, **provided** the server is launched from the project root. Update the Dockerfile `CMD` and the `if __name__ == "__main__"` block to always `cd` to the project root before launching:
   - Dockerfile: `CMD ["sh", "-c", "cd /app && uvicorn report_api:app --host 0.0.0.0 --port ${PORT}"]`
   - `if __name__ == "__main__"`: add `os.chdir(os.path.dirname(os.path.abspath(__file__)))` before `uvicorn.run(...)`.
4. Test that the import chain still resolves after removing `sys.path` manipulation by doing a dry-run: `python -c "from pages.modules.epw_parser import parse_epw; print('OK')"` from the project root.

---

### ISSUE 3 — Consistent API Error Responses [HIGH PRIORITY]

**Problem:**  
`report_api.py` catches exceptions with two patterns:
```python
except ValueError as e:
    raise HTTPException(status_code=400, detail=str(e))
except Exception as e:
    raise HTTPException(status_code=500, detail=f"Report generation failed: {str(e)}")
```
The `detail` is a plain string. There is also no input validation beyond ad-hoc `if` checks.

**Fix:**
1. Create a Pydantic error response model in `report_api.py`:
```python
from pydantic import BaseModel

class ErrorResponse(BaseModel):
    error: str
    detail: str
```
2. Update all `HTTPException` raises to pass a dict: `detail={"error": "validation_error", "detail": str(e)}` for 400s and `detail={"error": "generation_failed", "detail": str(e)}` for 500s.
3. Add `responses={400: {"model": ErrorResponse}, 500: {"model": ErrorResponse}}` to each endpoint decorator.
4. For the `combined-analysis` endpoint, extract the parameter validation block (the `if rainfall_station_name is not None:` block) into a dedicated `_validate_combined_params()` helper function that raises `ValueError` on bad input — keeping the handler body clean.
5. Similarly, extract the `if n_sectors not in [4, 8, 16]: n_sectors = 16` silently-corrects-invalid-input pattern into a validator that raises `HTTPException(status_code=400, ...)` instead of silently substituting a default.

---

### ISSUE 4 — Async NOAA Fetch [HIGH PRIORITY]

**Problem:**  
`pages/modules/rainfall_module.py` functions `_fetch_noaa()` and `_fetch_percentile_depth()` use `urllib.request.urlopen()` (synchronous). In `report_api.py`, the `generate_rainfall_report` and `generate_combined_report` endpoints are `async def` handlers. Calling synchronous HTTP inside them blocks the FastAPI event loop.

**Fix:**
1. Add `httpx` to `requirements_api.txt` (if not already present).
2. In `report_api.py`, add an async helper:
```python
import httpx

async def _fetch_noaa_async(url: str) -> bytes:
    async with httpx.AsyncClient(timeout=30.0) as client:
        resp = await client.get(url)
        resp.raise_for_status()
        return resp.content
```
3. In the `generate_rainfall_report` and `generate_combined_report` handlers, before calling `generate_rainfall_pptx_report()` / `generate_combined_pptx_report()`, pre-fetch the NOAA data asynchronously and pass the raw bytes (or parsed DataFrame) into the generator functions via a new optional parameter.
4. Alternatively (simpler): wrap the synchronous `generate_rainfall_pptx_report()` call in `asyncio.get_event_loop().run_in_executor(None, ...)` to run it in a thread pool, so the event loop is not blocked. This is the preferred approach if modifying the rainfall module internals is out of scope.

   **Recommended approach (least invasive):** In each async endpoint that calls a PPT generator, wrap the call:
```python
import asyncio
from functools import partial

loop = asyncio.get_event_loop()
pptx_buffer = await loop.run_in_executor(
    None,
    partial(generate_rainfall_pptx_report, station_name=..., ...)
)
```
   Apply this pattern to **all** report endpoints (not just rainfall) since Matplotlib/python-pptx operations are also CPU-bound and block the event loop.

---

### ISSUE 5 — Result Caching for Report Endpoints [HIGH PRIORITY]

**Problem:**  
Every API call regenerates the PPTX from scratch. The combined report takes multiple seconds.

**Fix:**
1. Add an in-memory LRU cache using `functools.lru_cache` (or `cachetools.TTLCache` — add `cachetools` to `requirements_api.txt`).
2. Create a cache key function in `report_api.py`:
```python
import hashlib

def _make_cache_key(file_content: bytes, **params) -> str:
    content_hash = hashlib.sha256(file_content).hexdigest()[:16]
    param_str = "&".join(f"{k}={v}" for k, v in sorted(params.items()))
    return f"{content_hash}:{param_str}"
```
3. Add a module-level `TTLCache(maxsize=64, ttl=1800)` (30 min TTL). Before calling the PPT generator, compute the cache key and check the cache. On a hit, stream the cached bytes directly. On a miss, generate, store in cache, then stream.
4. Apply to all POST report endpoints.

---

### ISSUE 6 — Expand Rainfall Station Coverage [MEDIUM PRIORITY]

**Problem:**  
`STATIONS` in `pages/modules/rainfall_module.py` is a hardcoded dict of 12 entries.

**Fix:**
1. Keep the existing hardcoded dict as a `_BUILTIN_STATIONS` fallback.
2. Add a `get_stations()` function that attempts to load additional stations from a JSON file `noaa_stations.json` at the project root if it exists, merging with the built-in list. This makes the list extensible without code changes:
```python
import json, pathlib

def get_stations() -> dict:
    stations = dict(_BUILTIN_STATIONS)
    override_path = pathlib.Path(__file__).parents[3] / "noaa_stations.json"
    if override_path.exists():
        try:
            extra = json.loads(override_path.read_text())
            stations.update(extra)
        except Exception:
            pass
    return stations

STATIONS = get_stations()
```
3. Create a template `noaa_stations.json` at the project root with a comment header explaining the format:
```json
{
  "_comment": "Add custom NOAA stations here. Format: { 'Station Name': 'NOAA_STATION_ID' }",
  "Dubai (DXB)": "AE000041196",
  "Abu Dhabi": "AE000041217"
}
```

---

### ISSUE 7 — Report Branding / White-Labelling [MEDIUM PRIORITY]

**Problem:**  
The `combined-analysis` endpoint has no way to customise the report's project name, client name, or report date shown on the cover slide.

**Fix:**
1. Add optional form parameters to `generate_combined_report` in `report_api.py`:
```python
project_name: Optional[str] = Form(None, description="Project name shown on cover slide"),
client_name: Optional[str] = Form(None, description="Client name shown on cover slide"),
report_date: Optional[str] = Form(None, description="Report date shown on cover (default: today, format: DD Month YYYY)"),
```
2. Pass these as a `branding` dict to `generate_combined_pptx_report()`.
3. In `pages/modules/combined_report.py`, update `_make_cover_slide()` to accept and render the `branding` dict values (project name, client name, date) on the cover slide. Fall back to current defaults if not provided.
4. Add `branding` parameter documentation to the `/api/docs` endpoint response.

---

### ISSUE 8 — PDF Export Option [MEDIUM PRIORITY]

**Problem:**  
Clients must convert PPTX to PDF themselves.

**Fix:**
1. Add an optional query parameter to all POST report endpoints:
```python
output_format: str = Query("pptx", description="Output format: 'pptx' (default) or 'pdf'")
```
2. Create a helper function in `report_api.py`:
```python
import subprocess, tempfile

def _pptx_to_pdf(pptx_buffer: io.BytesIO) -> io.BytesIO:
    with tempfile.TemporaryDirectory() as tmpdir:
        pptx_path = os.path.join(tmpdir, "report.pptx")
        with open(pptx_path, "wb") as f:
            f.write(pptx_buffer.getvalue())
        result = subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", tmpdir, pptx_path],
            capture_output=True, timeout=120
        )
        if result.returncode != 0:
            raise RuntimeError(f"LibreOffice conversion failed: {result.stderr.decode()}")
        pdf_path = pptx_path.replace(".pptx", ".pdf")
        buf = io.BytesIO(open(pdf_path, "rb").read())
        buf.seek(0)
        return buf
```
3. After generating `pptx_buffer`, if `output_format == "pdf"`, call `_pptx_to_pdf(pptx_buffer)` and return with `media_type="application/pdf"` and a `.pdf` filename. Add a note in the endpoint docstring: "PDF conversion requires LibreOffice to be installed on the server. Not available in the base Docker image without adding the libreoffice package."
4. Update the Dockerfile to optionally include LibreOffice: add a comment showing the `apt-get install -y libreoffice` line (commented out) with an explanation.

---

### ISSUE 9 — Complete URL Auto-Load EPW [MEDIUM PRIORITY]

**Problem:**  
`app.py` parses `?location=` but does not fetch the EPW file from `INDIA-WeatherMapping.xlsx` automatically.

**Fix:**
1. In `app.py`, after parsing `url_location_params`, add a lookup against the India weather mapping DataFrame (already loaded in `app.py` as part of the climate zone lookup). If a match is found and the row contains an EPW URL column, fetch the file automatically:
```python
# After url_location_params is parsed and location is matched:
if url_location_params.get("location") and matched_row is not None:
    epw_url = matched_row.get("EPW_URL")  # adjust column name to match actual Excel column
    if epw_url and isinstance(epw_url, str) and epw_url.startswith("http"):
        try:
            import urllib.request
            with urllib.request.urlopen(epw_url, timeout=15) as r:
                st.session_state["epw_auto_loaded"] = r.read()
                st.session_state["epw_auto_filename"] = epw_url.split("/")[-1]
        except Exception:
            pass
```
2. In the file uploader section of `app.py`, if `st.session_state.get("epw_auto_loaded")` is set, pre-populate the EPW data and skip the upload prompt, showing an info banner: "EPW file auto-loaded for [location name]. You can upload a different file to override."
3. Note: First read the actual column names in `INDIA-WeatherMapping.xlsx` with `pd.read_excel(...).columns.tolist()` to identify the correct EPW URL column name before implementing.

---

### ISSUE 10 — Full Type Hints [LOW PRIORITY]

**Problem:**  
Module functions lack return type annotations.

**Fix:**  
Add complete type hints to the public API of every module file in `pages/modules/`. At minimum annotate:
- All `def` function signatures with parameter types and return types.
- Use `pd.DataFrame`, `dict`, `tuple`, `str`, `int`, `float`, `bool`, `Optional[X]`, `list[X]` appropriately.
- For Plotly figures use `go.Figure`. For Matplotlib use `plt.Figure` / `matplotlib.figure.Figure`.
- For BytesIO buffers use `io.BytesIO`.
- Do NOT change any logic — only add annotations.

Files to annotate:
- `pages/modules/epw_parser.py`
- `pages/modules/dbt_module.py` (public functions only: `calculate_ashrae_comfort`, `render`)
- `pages/modules/humidity_module.py` (public: `render`)
- `pages/modules/wind_module.py` (public: `prepare_wind_data`, `compute_wind_rose`, `render_wind_analysis`)
- `pages/modules/ventilation_module.py` (public: `render`)
- `pages/modules/thermal_comfort_module.py` (public: `compute_psychrometric_data`, `compute_adaptive_comfort`, `classify_comfort`, `map_strategies`, `compute_degree_hours`, `generate_design_summary`, `render`)
- `pages/modules/sun_path.py` (public: `plot_sun_path`, `render_sun_path_section`)
- `pages/modules/rainfall_module.py` (public: `render`, `get_stations`)
- `pages/modules/ppt_report.py` (public: `generate_pptx_report`, `generate_shading_pptx_report`, `generate_wind_pptx_report`)
- `pages/modules/combined_report.py` (public: `generate_combined_pptx_report`)
- `pages/modules/thermal_comfort_ppt.py` (public: `generate_thermal_comfort_pptx_report`)
- `pages/modules/rainfall_ppt.py` (public: `generate_rainfall_pptx_report`)

---

### ISSUE 11 — Expand Test Coverage [LOW PRIORITY]

**Problem:**  
`test_api.py` only tests health and docs endpoints. `test_epw_parsing.py` has minimal coverage.

**Fix:**  
Rewrite `test_epw_parsing.py` and expand `test_api.py` using `pytest`. Add `pytest` and `pytest-asyncio` to `requirements_api.txt`.

**test_epw_parsing.py — required test cases:**
1. `test_parse_epw_basic` — parse the bundled `IND_DL_New.Delhi-Safdarjung.AP.421820_ISHRAE2014.epw` file; assert DataFrame has 8760 rows, all expected columns present, no NaT in datetime column.
2. `test_parse_epw_hour_normalisation` — assert `hour` column values are all in range [0, 23].
3. `test_parse_epw_metadata` — assert metadata contains `latitude`, `longitude`, `city`, `timezone` with non-None values; latitude is a float between -90 and 90.
4. `test_parse_epw_temperature_range` — assert dry_bulb_temperature values are all between -50 and 60 °C (no absurd outliers from parse errors).
5. `test_parse_epw_missing_columns_raises` — pass a truncated EPW string with fewer than 22 columns per row; assert `ValueError` is raised.

**test_api.py — required test cases** (use `httpx.AsyncClient` with FastAPI `TestClient`):
1. `test_health` — GET /api/health returns 200 and `{"status": "ok"}`.
2. `test_rainfall_stations` — GET /api/rainfall/stations returns 200 and response contains a `stations` list with at least 10 entries.
3. `test_climate_report_valid_epw` — POST /api/reports/climate-analysis with the bundled EPW file; assert 200, content-type is PPTX MIME type, Content-Disposition contains `.pptx`.
4. `test_climate_report_invalid_file` — POST /api/reports/climate-analysis with a text file instead of EPW; assert 400.
5. `test_combined_report_valid` — POST /api/reports/combined-analysis with the bundled EPW; assert 200 and PPTX response.
6. `test_combined_report_invalid_rainfall_station` — POST /api/reports/combined-analysis with `rainfall_station_name=INVALID_STATION` and `rainfall_year=2023`; assert 400 with JSON error body.
7. `test_wind_report_invalid_sectors` — POST /api/reports/wind-analysis with `n_sectors=7`; assert 400 (after implementing the validator from Issue 3).
8. `test_error_response_schema` — for any 400 response, assert the body is JSON with keys `error` and `detail`.

---

### ISSUE 12 — Centralise Configuration Constants [LOW PRIORITY]

**Problem:**  
Magic numbers are scattered across module files.

**Fix:**
1. Create `pages/modules/config.py` with all shared constants:
```python
"""Shared configuration constants for Climate Zone Finder analysis modules."""

# ── ASHRAE 55 Adaptive Comfort ─────────────────────────────────────────────────
ASHRAE_ALPHA = 0.9              # Exponential running mean coefficient
ASHRAE_T_PMA_MIN = 10.0         # Minimum applicable prevailing mean temp (°C)
ASHRAE_T_PMA_MAX = 33.5         # Maximum applicable prevailing mean temp (°C)
ASHRAE_COMFORT_NEUTRAL_A = 0.31 # Coefficient: T_comf = A * T_pma + B
ASHRAE_COMFORT_NEUTRAL_B = 17.8

# ── Comfort Bands ──────────────────────────────────────────────────────────────
COMFORT_BAND_80_PCT = 3.5       # ± °C around neutral for 80% acceptability
COMFORT_BAND_90_PCT = 2.5       # ± °C around neutral for 90% acceptability

# ── Degree Hours ───────────────────────────────────────────────────────────────
CDH_BASE_TEMP = 24.0            # Cooling degree-hours base temperature (°C)
HDH_BASE_TEMP = 18.0            # Heating degree-hours base temperature (°C)

# ── Thermal Comfort Strategies ─────────────────────────────────────────────────
NV_MIN_WIND_SPEED = 1.0         # Min wind speed for natural ventilation (m/s)
NV_COOL_DBT_THRESHOLD = 24.0    # DBT above which NV is counted (°C)
NIGHT_FLUSH_DIURNAL_MIN = 8.0   # Min diurnal range for night flushing (°C)
MECH_COOLING_RH_THRESHOLD = 60.0 # RH above which mech. cooling preferred over evap (%)

# ── Wind Analysis ──────────────────────────────────────────────────────────────
CALM_WIND_THRESHOLD = 0.5       # Calm wind cutoff (m/s, WMO convention)
WIND_SPEED_BINS = [0, 2, 4, 6, 8, 10, 15, 100]
WIND_SPEED_LABELS = ["0-2", "2-4", "4-6", "6-8", "8-10", "10-15", "15+"]

# ── Shading / Radiation ────────────────────────────────────────────────────────
DEFAULT_TEMP_THRESHOLD = 28.0   # Default overheating temperature (°C)
DEFAULT_RAD_THRESHOLD = 315.0   # Default radiation threshold (W/m²)
DEFAULT_CUTOFF_ANGLE = 45.0     # Default design cutoff angle (degrees)

# ── Humidity Comfort ───────────────────────────────────────────────────────────
RH_COMFORT_MIN = 30.0           # Lower bound of comfort RH band (%)
RH_COMFORT_MAX = 65.0           # Upper bound of comfort RH band (%)

# ── Psychrometrics ─────────────────────────────────────────────────────────────
P_ATM = 101_325.0               # Standard atmospheric pressure (Pa)

# ── Rainfall ───────────────────────────────────────────────────────────────────
DEFAULT_HEAVY_RAIN_THRESHOLD = 50.0   # mm/day
RUNOFF_COEFF_ROOF   = 0.90
RUNOFF_COEFF_PAVED  = 0.90
RUNOFF_COEFF_GREEN  = 0.10
RUNOFF_COEFF_WATER  = 0.90
VALID_GI_PERCENTILES = [85, 90, 95, 98]
```

2. Replace every hardcoded instance of these values in all module files with imports from `config.py`. Use your search tools to find all occurrences before editing — search for the numeric values (e.g. `0.31`, `17.8`, `3.5`, `2.5`, `0.5`, `101_325`, `50.0`, `0.90`, `0.10`) across `pages/modules/`.

3. Do the same replacements in `report_api.py` for the default parameter values in endpoint signatures.

---

### ISSUE 13 — Dockerise Streamlit App [LOW PRIORITY]

**Problem:**  
No Docker configuration exists for the Streamlit app.

**Fix:**
1. Create `Dockerfile.streamlit` in the project root:
```dockerfile
FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

WORKDIR /app

RUN apt-get update && apt-get install -y --no-install-recommends \
    libgl1 libglib2.0-0 libsm6 libxrender1 libxext6 \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

ENV PORT=8501
EXPOSE 8501

CMD ["sh", "-c", "streamlit run app.py --server.port=${PORT} --server.address=0.0.0.0 --server.headless=true"]
```

2. Create `docker-compose.yml` at the project root:
```yaml
version: "3.9"
services:
  api:
    build:
      context: .
      dockerfile: Dockerfile
    ports:
      - "8001:8080"
    environment:
      - PORT=8080

  streamlit:
    build:
      context: .
      dockerfile: Dockerfile.streamlit
    ports:
      - "8501:8501"
    environment:
      - PORT=8501
    depends_on:
      - api
```

3. Add a `.dockerignore` entry for `__pycache__`, `*.pyc`, `.git`, `node_modules`, `*.epw` (large test files), `*.pptx` (large template — note: the branded template IS needed, so exclude only non-template pptx), `*.xlsx` lock files (`~$*.xlsx`).

---

## EXECUTION INSTRUCTIONS

1. **Read before writing.** For each issue, read the relevant source files first to understand the current implementation before making changes.

2. **One issue at a time.** Complete and verify each issue before starting the next. Do not batch edits across multiple issues in a single step.

3. **Never break existing functionality.** After every change, check that the modified file is syntactically valid (run `python -c "import <module>"` for Python files, `node --check` for JS if applicable).

4. **Minimal diffs.** Make the smallest change that resolves each issue. Do not reformat unrelated code, rename existing variables, or reorganise file structure beyond what is explicitly required.

5. **Update requirements files.** Whenever a new package is needed (e.g. `httpx`, `cachetools`, `pytest`, `pytest-asyncio`), add it to the appropriate requirements file (`requirements.txt` for Streamlit, `requirements_api.txt` for API).

6. **After all issues are resolved**, run a final verification:
   - `python -c "from pages.modules.epw_parser import parse_epw; print('EPW parser OK')"` from the project root.
   - `python -c "import report_api; print('API imports OK')"` from the project root.
   - `python -m pytest test_epw_parsing.py test_api.py -v --tb=short` (if the server is not running, skip API endpoint tests with `-k "not test_climate_report and not test_combined and not test_wind and not test_rainfall_report"`).
   - Confirm all new/modified files are saved.

7. **Do not delete any existing files** unless explicitly instructed above. Do not touch `pages/modules/combined_report.py`, `pages/modules/ppt_report.py`, `pages/modules/thermal_comfort_ppt.py`, or `pages/modules/rainfall_ppt.py` beyond the specific changes described (type hints and branding parameter in combined_report.py).

---

## FILES YOU WILL NEED TO READ FIRST

Before starting, read these files in full to understand the current state:
- `report_api.py`
- `pages/modules/epw_parser.py`
- `pages/modules/rainfall_module.py` (the `_fetch_noaa` and `_fetch_percentile_depth` functions)
- `app.py` (for Issue 9 — look for the `url_location_params` usage and the India mapping DataFrame load)
- `INDIA-WeatherMapping.xlsx` column names (use `python -c "import pandas as pd; print(pd.read_excel('INDIA-WeatherMapping.xlsx').columns.tolist())"`)
- `test_api.py` and `test_epw_parsing.py`

