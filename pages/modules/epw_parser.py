"""EPW file parsing utilities."""

import io
import re
import pandas as pd


def convert_epw_timezone(tz_offset):
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


def parse_epw(epw_text: str) -> tuple:
    """Parse EPW formatted text and return a tuple of (DataFrame, metadata).

    Returns:
        tuple: (df, metadata) where df has datetime, dry_bulb_temperature,
               relative_humidity, hour and metadata contains latitude,
               longitude, timezone.
    """
    lines = [ln.strip() for ln in epw_text.splitlines() if ln.strip() != ""]

    # Extract metadata from header (first line)
    metadata = {
        "latitude": None,
        "longitude": None,
        "timezone": "UTC",
        "city": None,
        "location": None,
    }
    if len(lines) > 0:
        header = lines[0].split(",")
        try:
            # EPW header format:
            # LOCATION,CITY,STATE,COUNTRY,DATA SOURCE,WMO #,LAT,LON,TZ,ELEV
            if len(header) >= 2:
                metadata["location"] = header[0].strip()
                metadata["city"] = header[1].strip()
            if len(header) >= 8:
                metadata["latitude"] = float(header[6].strip())
                metadata["longitude"] = float(header[7].strip())
            if len(header) >= 9:
                metadata["timezone"] = convert_epw_timezone(header[8].strip())
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
    # 6=dry bulb (°C), 7=dew point, 8=relative humidity (%),
    # 13=global_horizontal_irradiance (Wh/m²),
    # 14=direct_normal_irradiance (Wh/m²),
    # 15=diffuse_horizontal_irradiance (Wh/m²)
    col_map = {
        "year": 0,
        "month": 1,
        "day": 2,
        "hour": 3,
        "minute": 4,
        "dry_bulb_temperature": 6,
        "relative_humidity": 8,
        "direct_normal_irradiance": 14,
        "diffuse_horizontal_irradiance": 15,
    }

    max_needed = max(col_map.values())
    if df_raw.shape[1] <= max_needed:
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
    df["relative_humidity"] = pd.to_numeric(
        df_raw.iloc[:, col_map["relative_humidity"]], errors="coerce"
    )
    df["direct_normal_irradiance"] = pd.to_numeric(
        df_raw.iloc[:, col_map["direct_normal_irradiance"]], errors="coerce"
    )
    df["diffuse_horizontal_irradiance"] = pd.to_numeric(
        df_raw.iloc[:, col_map["diffuse_horizontal_irradiance"]], errors="coerce"
    ).fillna(0)
    df["global_horizontal_irradiance"] = pd.to_numeric(
        df_raw.iloc[:, 13], errors="coerce"
    ).fillna(0)

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
    return (
        df[[
            "datetime",
            "dry_bulb_temperature",
            "relative_humidity",
            "direct_normal_irradiance",
            "diffuse_horizontal_irradiance",
            "global_horizontal_irradiance",
            "hour",
        ]],
        metadata,
    )
