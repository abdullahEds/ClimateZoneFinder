"""EPW file parsing utilities."""

import io
import re
from typing import Optional, Union
import pandas as pd


def convert_epw_timezone(tz_offset: Union[str, float, int]) -> str:
    """Convert EPW numeric timezone to valid pytz timezone string."""
    tz_map = {
        5.5: "Asia/Kolkata",
        0: "UTC",
        -5: "Etc/GMT+5",
        -6: "Etc/GMT+6",
        -7: "Etc/GMT+7",
        -8: "Etc/GMT+8",
        1: "Europe/London",
        2: "Europe/Paris",
    }
    try:
        tz_float = float(tz_offset)
        if tz_float in tz_map:
            return tz_map[tz_float]
    except (ValueError, TypeError):
        pass
    return "UTC"


def parse_epw(epw_text: str) -> tuple[pd.DataFrame, dict]:
    """Parse EPW formatted text and return a tuple of (DataFrame, metadata).

    Returns a DataFrame with all columns needed by both the Streamlit app and
    the FastAPI report endpoints:
        datetime, dry_bulb_temperature, relative_humidity,
        direct_normal_irradiance, diffuse_horizontal_irradiance,
        global_horizontal_irradiance, wind_direction, wind_speed, hour,
        dew_point_temperature, atmospheric_pressure, liquid_precipitation_depth,
        doy, month, Year, Month, Day, Minute.

    Metadata dict contains: latitude, longitude, timezone, city, location,
        state, country, elevation.
    """
    lines = [ln.strip() for ln in epw_text.splitlines() if ln.strip() != ""]

    # Extract metadata from header (first line)
    # EPW header: LOCATION,CITY,STATE,COUNTRY,DATA SOURCE,WMO #,LAT,LON,TZ,ELEV
    metadata = {
        "latitude": None,
        "longitude": None,
        "timezone": "UTC",
        "city": None,
        "location": None,
        "state": None,
        "country": None,
        "elevation": None,
    }
    if len(lines) > 0:
        header = lines[0].split(",")
        try:
            if len(header) >= 2:
                metadata["location"] = header[0].strip()
                metadata["city"] = header[1].strip()
            if len(header) >= 3:
                metadata["state"] = header[2].strip()
            if len(header) >= 4:
                metadata["country"] = header[3].strip()
            if len(header) >= 8:
                metadata["latitude"] = float(header[6].strip())
                metadata["longitude"] = float(header[7].strip())
            if len(header) >= 9:
                metadata["timezone"] = convert_epw_timezone(header[8].strip())
            if len(header) >= 10:
                metadata["elevation"] = float(header[9].strip())
        except (ValueError, IndexError, TypeError):
            pass

    data_start = None
    for i, ln in enumerate(lines):
        toks = ln.split(",")
        if len(toks) > 1 and re.fullmatch(r"\d{4}", toks[0].strip()):
            data_start = i
            break

    if data_start is None:
        raise ValueError("Could not locate EPW data rows")

    data_str = "\n".join(lines[data_start:])
    df_raw = pd.read_csv(io.StringIO(data_str), header=None)

    # EPW standard column indices (0-based):
    # 0=year, 1=month, 2=day, 3=hour, 4=minute, 5=data source,
    # 6=dry_bulb (°C), 7=dew_point (°C), 8=relative_humidity (%),
    # 9=atmospheric_pressure (Pa),
    # 13=global_horizontal_irradiance (Wh/m²),
    # 14=direct_normal_irradiance (Wh/m²),
    # 15=diffuse_horizontal_irradiance (Wh/m²),
    # 20=wind_direction (°), 21=wind_speed (m/s),
    # 33=liquid_precipitation_depth (mm)
    col_map = {
        "year": 0,
        "month": 1,
        "day": 2,
        "hour": 3,
        "minute": 4,
        "dry_bulb_temperature": 6,
        "dew_point_temperature": 7,
        "relative_humidity": 8,
        "atmospheric_pressure": 9,
        "global_horizontal_irradiance": 13,
        "direct_normal_irradiance": 14,
        "diffuse_horizontal_irradiance": 15,
        "wind_direction": 20,
        "wind_speed": 21,
        "liquid_precipitation_depth": 33,
    }

    max_needed = max(col_map.values())
    if df_raw.shape[1] <= max(col_map["wind_speed"], col_map["wind_direction"]):
        raise ValueError("EPW data appears to have insufficient columns")

    df = pd.DataFrame()
    df["year"] = pd.to_numeric(df_raw.iloc[:, col_map["year"]], errors="coerce").astype("Int64")
    df["month"] = pd.to_numeric(df_raw.iloc[:, col_map["month"]], errors="coerce").astype("Int64")
    df["day"] = pd.to_numeric(df_raw.iloc[:, col_map["day"]], errors="coerce").astype("Int64")
    df["hour_raw"] = pd.to_numeric(df_raw.iloc[:, col_map["hour"]], errors="coerce").astype("Int64")
    df["minute"] = pd.to_numeric(df_raw.iloc[:, col_map["minute"]], errors="coerce").astype("Int64")

    # EPW hours are 1-24 (hour ending); convert to 0-23
    df["hour"] = (df["hour_raw"].fillna(1).astype(int) - 1) % 24

    df["dry_bulb_temperature"] = pd.to_numeric(
        df_raw.iloc[:, col_map["dry_bulb_temperature"]], errors="coerce"
    )
    df["dew_point_temperature"] = pd.to_numeric(
        df_raw.iloc[:, col_map["dew_point_temperature"]], errors="coerce"
    )
    df["relative_humidity"] = pd.to_numeric(
        df_raw.iloc[:, col_map["relative_humidity"]], errors="coerce"
    )
    df["atmospheric_pressure"] = pd.to_numeric(
        df_raw.iloc[:, col_map["atmospheric_pressure"]], errors="coerce"
    )
    df["direct_normal_irradiance"] = pd.to_numeric(
        df_raw.iloc[:, col_map["direct_normal_irradiance"]], errors="coerce"
    )
    df["diffuse_horizontal_irradiance"] = pd.to_numeric(
        df_raw.iloc[:, col_map["diffuse_horizontal_irradiance"]], errors="coerce"
    ).fillna(0)
    df["global_horizontal_irradiance"] = pd.to_numeric(
        df_raw.iloc[:, col_map["global_horizontal_irradiance"]], errors="coerce"
    ).fillna(0)
    df["wind_direction"] = pd.to_numeric(
        df_raw.iloc[:, col_map["wind_direction"]], errors="coerce"
    ).fillna(0.0)
    df["wind_speed"] = pd.to_numeric(
        df_raw.iloc[:, col_map["wind_speed"]], errors="coerce"
    ).fillna(0.0)

    # liquid_precipitation_depth is column 33 — present in most EPW files
    if df_raw.shape[1] > col_map["liquid_precipitation_depth"]:
        df["liquid_precipitation_depth"] = pd.to_numeric(
            df_raw.iloc[:, col_map["liquid_precipitation_depth"]], errors="coerce"
        ).fillna(0.0)
    else:
        df["liquid_precipitation_depth"] = 0.0

    df["datetime"] = pd.to_datetime(
        dict(
            year=df["year"],
            month=df["month"],
            day=df["day"],
            hour=df["hour"],
            minute=df["minute"],
        ),
        errors="coerce",
    )

    df = df.dropna(subset=["datetime"]).reset_index(drop=True)

    # Derived columns used by the API parser
    df["doy"] = df["datetime"].dt.dayofyear
    # Aliased uppercase columns for API compatibility
    df["Year"] = df["year"].astype(int)
    df["Month"] = df["month"].astype(int)
    df["Day"] = df["day"].astype(int)
    df["Minute"] = df["minute"].astype(int)

    return (
        df[[
            "datetime",
            "dry_bulb_temperature",
            "dew_point_temperature",
            "relative_humidity",
            "atmospheric_pressure",
            "direct_normal_irradiance",
            "diffuse_horizontal_irradiance",
            "global_horizontal_irradiance",
            "wind_direction",
            "wind_speed",
            "liquid_precipitation_depth",
            "hour",
            "doy",
            "Year",
            "Month",
            "Day",
            "Minute",
            # lowercase aliases kept for Streamlit consumers
            "month",
        ]],
        metadata,
    )
