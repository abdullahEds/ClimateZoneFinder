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
