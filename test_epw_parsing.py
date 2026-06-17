"""Pytest test suite for EPW parsing (pages/modules/epw_parser.py)."""

import pathlib
import pytest
import pandas as pd

from pages.modules.epw_parser import parse_epw

_EPW_PATH = pathlib.Path(__file__).parent / "IND_DL_New.Delhi-Safdarjung.AP.421820_ISHRAE2014.epw"
_BUNDLED_EPW_TEXT = _EPW_PATH.read_text(encoding="utf-8", errors="replace") if _EPW_PATH.exists() else None

_SKIP_NO_FILE = pytest.mark.skipif(
    _BUNDLED_EPW_TEXT is None,
    reason="Bundled EPW file not found",
)

EXPECTED_COLUMNS = {
    "datetime",
    "dry_bulb_temperature",
    "relative_humidity",
    "direct_normal_irradiance",
    "diffuse_horizontal_irradiance",
    "global_horizontal_irradiance",
    "wind_direction",
    "wind_speed",
    "hour",
    "dew_point_temperature",
    "atmospheric_pressure",
    "liquid_precipitation_depth",
    "doy",
    "month",
    "Year",
    "Month",
    "Day",
    "Minute",
}


@_SKIP_NO_FILE
def test_parse_epw_basic():
    df, metadata = parse_epw(_BUNDLED_EPW_TEXT)
    assert len(df) == 8760, f"Expected 8760 rows, got {len(df)}"
    missing = EXPECTED_COLUMNS - set(df.columns)
    assert not missing, f"Missing columns: {missing}"
    assert df["datetime"].isna().sum() == 0, "NaT values found in datetime column"


@_SKIP_NO_FILE
def test_parse_epw_hour_normalisation():
    df, _ = parse_epw(_BUNDLED_EPW_TEXT)
    assert df["hour"].min() >= 0
    assert df["hour"].max() <= 23


@_SKIP_NO_FILE
def test_parse_epw_metadata():
    _, metadata = parse_epw(_BUNDLED_EPW_TEXT)
    for key in ("latitude", "longitude", "city", "timezone"):
        assert metadata.get(key) is not None, f"metadata['{key}'] is None"
    lat = metadata["latitude"]
    assert isinstance(lat, float), "latitude must be a float"
    assert -90.0 <= lat <= 90.0, f"latitude out of range: {lat}"


@_SKIP_NO_FILE
def test_parse_epw_temperature_range():
    df, _ = parse_epw(_BUNDLED_EPW_TEXT)
    t = df["dry_bulb_temperature"].dropna()
    assert t.min() >= -50, f"Temperature too low: {t.min()}"
    assert t.max() <= 60, f"Temperature too high: {t.max()}"


def test_parse_epw_missing_columns_raises():
    # Construct a minimal EPW with only 21 columns per data row (< 22 needed)
    header = "LOCATION,TestCity,TestState,TestCountry,Source,999999,28.6,77.2,5.5,216"
    # 8 header lines, then data lines with only 21 columns
    data_row = ",".join(["2001", "1", "1", "1", "0"] + ["0"] * 16)  # 21 columns total
    lines = [header] + [""] * 7 + [data_row] * 10
    truncated_epw = "\n".join(lines)
    with pytest.raises(ValueError):
        parse_epw(truncated_epw)
