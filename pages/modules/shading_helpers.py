"""Shading analysis helper functions and constants."""

import numpy as np
import pandas as pd
import pytz
import plotly.graph_objects as go


_MONTH_SHORT = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

_ORIENTATIONS = {
    "North (0°)": 0,
    "North-East (45°)": 45,
    "East (90°)": 90,
    "South-East (135°)": 135,
    "South (180°)": 180,
    "South-West (225°)": 225,
    "West (270°)": 270,
    "North-West (315°)": 315,
}


def build_thermal_matrix(epw_df: pd.DataFrame, temp_threshold: float, rad_threshold: float):
    """Build a 24×12 matrix of mean temperature and GHI per (hour, month).

    Returns:
        temp_matrix:   24×12 DataFrame of mean dry bulb temperature
        rad_matrix:    24×12 DataFrame of mean GHI
        overheat_mask: bool 24×12 DataFrame where both thresholds are exceeded
    """
    d = epw_df.copy()
    d["_month"] = d["datetime"].dt.month
    d["_hour"] = d["hour"].astype(int)

    temp_pivot = (
        d.groupby(["_hour", "_month"])["dry_bulb_temperature"]
        .mean()
        .unstack("_month")
    )
    rad_pivot = (
        d.groupby(["_hour", "_month"])["global_horizontal_irradiance"]
        .mean()
        .unstack("_month")
    )

    temp_pivot.columns = _MONTH_SHORT
    rad_pivot.columns = _MONTH_SHORT
    temp_pivot.index.name = "Hour"
    rad_pivot.index.name = "Hour"

    overheat_mask = (temp_pivot > temp_threshold) & (rad_pivot > rad_threshold)
    return temp_pivot, rad_pivot, overheat_mask


def get_overheating_hours(epw_df: pd.DataFrame, temp_threshold: float, rad_threshold: float) -> pd.DataFrame:
    """Return EPW rows where temperature > temp_threshold AND GHI > rad_threshold."""
    mask = (
        (epw_df["dry_bulb_temperature"] > temp_threshold) &
        (epw_df["global_horizontal_irradiance"] > rad_threshold)
    )
    return epw_df[mask].copy().reset_index(drop=True)


def compute_solar_angles(overheat_df: pd.DataFrame, lat: float, lon: float, tz_str: str) -> pd.DataFrame:
    """Compute solar altitude and azimuth for each overheating timestamp using pvlib.

    Returns overheat_df rows above the horizon with solar_altitude and solar_azimuth columns.
    """
    from pvlib import solarposition

    try:
        tz = pytz.timezone(tz_str)
    except Exception:
        tz = pytz.UTC

    times = pd.DatetimeIndex(overheat_df["datetime"].values)
    if times.tz is None:
        times = times.tz_localize(tz)

    # Normalise to year 2020 for consistency with plot_sun_path
    times_2020 = times.map(lambda t: t.replace(year=2020))
    solpos = solarposition.get_solarposition(times_2020, lat, lon)

    result = overheat_df.copy()
    result["solar_altitude"] = solpos["apparent_elevation"].values
    result["solar_azimuth"] = solpos["azimuth"].values

    result = result[result["solar_altitude"] > 0].copy().reset_index(drop=True)
    return result


def compute_shading_geometry(solar_positions: pd.DataFrame, facade_azimuth: float) -> pd.DataFrame:
    """Compute VSA, HSA and facade-hit flag for a given facade azimuth.

    Formulas:
        relative_azimuth = solar_azimuth − facade_azimuth
        VSA = arctan( tan(altitude) / cos(relative_azimuth) )
        HSA = arctan( sin(relative_azimuth) / tan(altitude) )
    """
    alt_rad = np.radians(solar_positions["solar_altitude"].values)
    rel_az = solar_positions["solar_azimuth"].values - facade_azimuth
    rel_az = ((rel_az + 180) % 360) - 180   # normalise to −180…+180
    rel_az_rad = np.radians(rel_az)

    hits_facade = np.abs(rel_az) < 90

    tan_alt = np.tan(alt_rad)
    cos_rel = np.cos(rel_az_rad)
    sin_rel = np.sin(rel_az_rad)

    cos_rel_safe = np.where(np.abs(cos_rel) < 1e-9, np.sign(cos_rel + 1e-18) * 1e-9, cos_rel)
    tan_alt_safe = np.where(np.abs(tan_alt) < 1e-9, 1e-9, tan_alt)

    vsa_deg = np.degrees(np.arctan(tan_alt / cos_rel_safe))
    hsa_deg = np.degrees(np.arctan(sin_rel / tan_alt_safe))

    result = solar_positions.copy()
    result["relative_azimuth"] = rel_az
    result["VSA"] = vsa_deg
    result["HSA"] = hsa_deg
    result["hits_facade"] = hits_facade
    return result


def build_orientation_table(solar_positions: pd.DataFrame, design_cutoff_angle: float) -> pd.DataFrame:
    """Build the 8-orientation shading analysis table."""
    rows = []
    for orient_name, facade_az in _ORIENTATIONS.items():
        geom = compute_shading_geometry(solar_positions, facade_az)
        facing = geom[geom["hits_facade"]]

        if facing.empty:
            rows.append({
                "Orientation": orient_name,
                "Rays Hitting": 0,
                "Min VSA (°)": None,
                "Max |HSA| (°)": None,
                "D/H (Overhang)": None,
                "D/W (Fin)": None,
                "Protection (%)": 100.0,
            })
            continue

        min_vsa = float(facing["VSA"].min())
        max_abs_hsa = float(facing["HSA"].abs().max())

        tan_min_vsa = np.tan(np.radians(min_vsa))
        dh = round(min(1.0, float(1.0 / tan_min_vsa)) if abs(tan_min_vsa) > 1e-6 else 1.0, 3)
        dw = round(min(1.0, float(np.tan(np.radians(max_abs_hsa)))), 3)

        blocked = int((facing["VSA"] >= design_cutoff_angle).sum())
        protection_pct = (blocked / len(facing)) * 100

        rows.append({
            "Orientation": orient_name,
            "Rays Hitting": int(len(facing)),
            "Min VSA (°)": round(min_vsa, 1),
            "Max |HSA| (°)": round(max_abs_hsa, 1),
            "D/H (Overhang)": dh,
            "D/W (Fin)": dw,
            "Protection (%)": round(protection_pct, 1),
        })

    return pd.DataFrame(rows)


def make_shading_mask_chart(solar_positions: pd.DataFrame, facade_azimuth: float,
                             design_cutoff_angle: float):
    """Create a small polar chart showing overheating sun positions and shading cutoff arc."""
    geom = compute_shading_geometry(solar_positions, facade_azimuth)
    facing = geom[geom["hits_facade"]]
    other = geom[~geom["hits_facade"]]

    fig = go.Figure()

    if not other.empty:
        fig.add_trace(go.Scatterpolar(
            r=(90 - other["solar_altitude"]).tolist(),
            theta=other["solar_azimuth"].tolist(),
            mode="markers",
            marker=dict(size=3, color="lightgrey", opacity=0.5),
            showlegend=False,
            hoverinfo="skip",
        ))

    if not facing.empty:
        fig.add_trace(go.Scatterpolar(
            r=(90 - facing["solar_altitude"]).tolist(),
            theta=facing["solar_azimuth"].tolist(),
            mode="markers",
            marker=dict(size=4, color="#E65100", opacity=0.75),
            showlegend=False,
            hovertemplate="Alt: %{text}<br>Az: %{theta:.0f}°<extra></extra>",
            text=[f"{a:.1f}°" for a in facing["solar_altitude"]],
        ))

    # Shading cutoff arc
    rel_az_range = np.linspace(-89, 89, 179)
    tan_cutoff = np.tan(np.radians(design_cutoff_angle))
    cutoff_alt = np.degrees(np.arctan(tan_cutoff * np.cos(np.radians(rel_az_range))))
    cutoff_az_abs = facade_azimuth + rel_az_range
    valid = cutoff_alt > 0
    if valid.any():
        fig.add_trace(go.Scatterpolar(
            r=(90 - cutoff_alt[valid]).tolist(),
            theta=cutoff_az_abs[valid].tolist(),
            mode="lines",
            line=dict(color="#1565C0", width=2, dash="dash"),
            showlegend=False,
            hoverinfo="skip",
        ))

    # Facade direction indicator
    fig.add_trace(go.Scatterpolar(
        r=[0, 85],
        theta=[facade_azimuth, facade_azimuth],
        mode="lines",
        line=dict(color="#388E3C", width=2),
        showlegend=False,
        hoverinfo="skip",
    ))

    fig.update_layout(
        polar=dict(
            bgcolor="rgba(240, 248, 255, 0.6)",
            radialaxis=dict(visible=False, range=[0, 90]),
            angularaxis=dict(
                tickfont=dict(size=7),
                rotation=90,
                direction="clockwise",
                tickvals=[0, 90, 180, 270],
                ticktext=["N", "E", "S", "W"],
            ),
        ),
        showlegend=False,
        height=180,
        margin=dict(l=10, r=10, t=10, b=10),
        paper_bgcolor="rgba(0,0,0,0)",
    )
    return fig
