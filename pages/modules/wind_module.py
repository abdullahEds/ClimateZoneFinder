"""Wind Analysis module – replicates the Wind tab of Berkeley CBE Clima tool.

Exposes:
    render_wind_analysis(epw_df)  ← called from pages/analysis.py

Meteorological algorithm notes
--------------------------------
• Calm winds (< 0.5 m/s, WMO/ASHRAE convention) are excluded from all
  directional calculations.  Calm % is computed over the full period and
  displayed as a center annotation on the wind rose.

• Wind direction is normalised with ``% 360`` so the 0/360° boundary
  never causes wrap-around artefacts.

• Direction sectors are CENTRED on cardinal/intercardinal points:
    sector i is centred at  i * (360 / n_sectors) degrees.
  Assignment algorithm:
    shifted = (direction + sector_width / 2) % 360
    sector_idx = floor(shifted / sector_width) % n_sectors
  This places North (0°) in the centre of sector 0.

• Wind rose frequencies are expressed as **% of total hours** so the
  sum of all bars + calm% ≈ 100 %.  When the "exclude calm" toggle is
  on, bars are renormalised by non-calm hours only (shows directional
  distribution among wind hours).

• The direction heatmap uses vector (circular) averaging via atan2 to
  handle the 0/360° discontinuity correctly:
    u = cos(θ),  v = sin(θ)
    θ_mean = atan2(mean_v, mean_u)  →  convert to [0, 360).
"""

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
# import streamlit as st


# ─── Module-level constants ───────────────────────────────────────────────────

# Standard compass labels
_DIR_16 = [
    "N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE",
    "S", "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW",
]
_DIR_8 = ["N", "NE", "E", "SE", "S", "SW", "W", "NW"]
_DIR_4 = ["N", "E", "S", "W"]

# Speed bin edges and display labels (m/s)
_SPEED_BINS   = [0, 2, 4, 6, 8, 10, 15, 100]
_SPEED_LABELS = ["0–2", "2–4", "4–6", "6–8", "8–10", "10–15", "15+"]

# Diverging colour scale matching CBE Clima palette
_SPEED_COLORS = [
    "#313695", "#4575b4", "#74add1", "#abd9e9",
    "#fdae61", "#f46d43", "#d73027",
]

_MONTH_NAMES = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
]


# ─── Data preparation ─────────────────────────────────────────────────────────

def prepare_wind_data(
    df: pd.DataFrame,
    months: list = None,
    n_sectors: int = 16,
) -> pd.DataFrame:
    """Prepare and filter EPW wind data for visualisation.

    Steps
    -----
    1. Ensure time-component columns (month, hour, dayofyear) are present.
    2. Sanitise wind_speed and wind_direction (clamp, normalise to [0, 360)).
    3. Flag calm hours (wind_speed < 0.5 m/s).
    4. Apply month filter.
    5. Assign direction sectors using centered-bin algorithm.
    6. Assign speed bins via pd.cut.

    Parameters
    ----------
    df        : EPW dataframe.  Must contain wind_speed and wind_direction.
    months    : list of month numbers (1–12) to keep; None keeps all.
    n_sectors : Number of compass sectors (default 16).

    Returns
    -------
    Filtered DataFrame with added columns:
        month, hour, dayofyear, is_calm,
        sector_idx, direction_label, speed_bin
    """
    wdf = df.copy()

    # ── Time components ───────────────────────────────────────────────────────
    # Prefer precomputed columns (added by analysis.py) over recomputing them.
    if "datetime" in wdf.columns:
        dt = pd.to_datetime(wdf["datetime"])
        if "month" not in wdf.columns:
            wdf["month"] = dt.dt.month
        if "hour" not in wdf.columns:
            wdf["hour"] = dt.dt.hour
        # analysis.py stores day-of-year as "doy"; expose it as "dayofyear"
        if "dayofyear" not in wdf.columns:
            wdf["dayofyear"] = wdf["doy"] if "doy" in wdf.columns else dt.dt.dayofyear

    # ── Wind data sanitisation ────────────────────────────────────────────────
    # EPW uses 999 for missing values; clip to physically realistic range.
    wdf["wind_speed"] = (
        pd.to_numeric(wdf["wind_speed"], errors="coerce")
        .fillna(0.0)
        .clip(lower=0.0, upper=50.0)
    )
    # Normalise direction to [0, 360) – this eliminates 0/360 boundary errors.
    wdf["wind_direction"] = (
        pd.to_numeric(wdf["wind_direction"], errors="coerce")
        .fillna(0.0)
        % 360.0
    )

    # Calm threshold: WMO / ASHRAE convention
    wdf["is_calm"] = wdf["wind_speed"] < 0.5

    # ── Month filter ──────────────────────────────────────────────────────────
    if months:
        wdf = wdf[wdf["month"].isin(months)].copy()

    if wdf.empty:
        return wdf

    # ── Direction sector assignment (centered bins) ───────────────────────────
    # Each sector is centred on its midpoint angle.
    # Shifting by half a sector width before flooring ensures that 0° falls in
    # the middle of sector 0 (North) rather than on a sector boundary.
    sector_width = 360.0 / n_sectors
    shifted = (wdf["wind_direction"] + sector_width / 2.0) % 360.0
    wdf["sector_idx"] = (shifted / sector_width).astype(int) % n_sectors

    # Compass label lookup
    if n_sectors == 16:
        label_list = _DIR_16
    elif n_sectors == 8:
        label_list = _DIR_8
    elif n_sectors == 4:
        label_list = _DIR_4
    else:
        # Generic degree labels for non-standard sector counts
        angles = np.arange(0, 360, sector_width)
        label_list = [f"{int(a)}°" for a in angles]

    idx_to_label = {i: label_list[i] for i in range(n_sectors)}
    wdf["direction_label"] = wdf["sector_idx"].map(idx_to_label)

    # ── Speed bins ────────────────────────────────────────────────────────────
    wdf["speed_bin"] = pd.cut(
        wdf["wind_speed"],
        bins=_SPEED_BINS,
        labels=_SPEED_LABELS,
        right=False,
        include_lowest=True,
    )

    return wdf.reset_index(drop=True)


# ─── Wind Rose computation ────────────────────────────────────────────────────

def compute_wind_rose(
    wdf: pd.DataFrame,
    n_sectors: int = 16,
    exclude_calm: bool = False,
) -> tuple:
    """Build sector × speed-bin frequency table for the wind rose.

    Calm winds are always excluded from directional bars.
    The ``exclude_calm`` flag controls the denominator:
      - False  : normalise by total hours  → bars + calm% ≈ 100 %
  - True   : normalise by non-calm hours → shows directional share
                                           among wind-only hours

    Returns
    -------
    (rose_df, calm_percent)
    rose_df columns : direction_label, speed_bin, frequency_pct
    """
    total = len(wdf)
    if total == 0:
        return pd.DataFrame(), 0.0

    calm_count = int(wdf["is_calm"].sum())
    calm_pct   = calm_count / total * 100.0

    active = wdf[~wdf["is_calm"]].copy()
    if active.empty:
        return pd.DataFrame(), calm_pct

    denominator = len(active) if exclude_calm else total

    grouped = (
        active
        .groupby(["direction_label", "speed_bin"], observed=True)
        .size()
        .reset_index(name="count")
    )
    grouped["frequency_pct"] = grouped["count"] / denominator * 100.0

    return grouped, calm_pct


# ─── Plot: Wind Rose ──────────────────────────────────────────────────────────

def plot_wind_rose(
    rose_df: pd.DataFrame,
    calm_pct: float,
    n_sectors: int = 16,
) -> go.Figure:
    """Polar bar chart (Barpolar) wind rose.

    - Bars stacked by speed tier
    - Radial axis = % of hours
    - North at top, clockwise rotation (meteorological convention)
    - Calm % displayed as a centre annotation
    """
    sector_width  = 360.0 / n_sectors
    sector_angles = [i * sector_width for i in range(n_sectors)]

    if n_sectors == 16:
        label_list = _DIR_16
    elif n_sectors == 8:
        label_list = _DIR_8
    elif n_sectors == 4:
        label_list = _DIR_4
    else:
        label_list = [f"{int(a)}°" for a in sector_angles]

    label_to_angle = dict(zip(label_list, sector_angles))

    if rose_df.empty:
        fig = go.Figure()
        fig.add_annotation(
            text="No directional wind data available",
            xref="paper", yref="paper", x=0.5, y=0.5,
            showarrow=False, font_size=14,
        )
        fig.update_layout(height=520)
        return fig

    fig = go.Figure()

    for i, spd_lbl in enumerate(_SPEED_LABELS):
        subset   = rose_df[rose_df["speed_bin"] == spd_lbl]
        freq_map = dict(zip(subset["direction_label"], subset["frequency_pct"]))

        # Build complete arrays for all sectors (fill missing with 0.0)
        angles = [label_to_angle[lbl] for lbl in label_list]
        freqs  = [freq_map.get(lbl, 0.0)  for lbl in label_list]

        fig.add_trace(go.Barpolar(
            r                 = freqs,
            theta             = angles,
            name              = f"{spd_lbl} m/s",
            marker_color      = _SPEED_COLORS[i % len(_SPEED_COLORS)],
            marker_line_color = "white",
            marker_line_width = 0.5,
            opacity           = 0.9,
            hovertemplate     = (
                "<b>%{theta:.0f}°</b><br>%{r:.2f}%"
                "<extra>" + spd_lbl + " m/s</extra>"
            ),
        ))

    fig.update_layout(
        title=dict(text="Wind Rose",  font_size=16, font_color="#2c3e50"),
        polar=dict(
            radialaxis=dict(
                visible        = True,
                ticksuffix     = "%",
                gridcolor      = "rgba(128,128,128,0.3)",
                linecolor      = "rgba(128,128,128,0.3)",
                tickfont_size  = 10,
            ),
            angularaxis=dict(
                # rotation=90  → 0° at the top = North
                # direction="clockwise" → standard meteorological convention
                rotation       = 90,
                direction      = "clockwise",
                tickmode       = "array",
                tickvals       = sector_angles,
                ticktext       = label_list,
                tickfont_size  = 11,
                gridcolor      = "rgba(128,128,128,0.3)",
                linecolor      = "rgba(128,128,128,0.4)",
            ),
        ),
        legend=dict(
            title       = "Wind Speed",
            orientation = "v",
            x=1.05, y=0.5,
            font_size   = 11,
        ),
        showlegend = True,
        height     = 520,
        template   = "plotly_white",
        annotations=[dict(
            text      = f"Calm<br>{calm_pct:.1f}%",
            x=0.5, y=0.5,
            xref="paper", yref="paper",
            showarrow = False,
            font      = dict(size=11, color="#555"),
            align     = "center",
        )],
    )
    return fig


# ─── Plot: Speed Heatmap ──────────────────────────────────────────────────────

def plot_speed_heatmap(wdf: pd.DataFrame) -> go.Figure:
    """Wind speed heatmap: day-of-year (rows) × hour-of-day (columns).

    Each cell = mean wind speed for that day × hour combination.
    Color scale: Viridis.
    """
    pivot = wdf.pivot_table(
        values  = "wind_speed",
        index   = "dayofyear",
        columns = "hour",
        aggfunc = "mean",
    )

    fig = px.imshow(
        pivot,
        labels = dict(x="Hour of Day", y="Day of Year", color="m/s"),
        title  = "Wind Speed – Day × Hour",
        color_continuous_scale = "Viridis",
        aspect = "auto",
        origin = "lower",
    )
    fig.update_layout(
        height   = 400,
        template = "plotly_white",
        xaxis    = dict(title="Hour of Day", tickmode="linear", dtick=3),
        yaxis    = dict(title="Day of Year"),
        coloraxis_colorbar = dict(title="m/s"),
        # title    = dict(x=0.5),
    )
    return fig


# ─── Plot: Direction Heatmap ──────────────────────────────────────────────────

def plot_direction_heatmap(wdf: pd.DataFrame) -> go.Figure:
    """Wind direction heatmap: day-of-year (rows) × hour-of-day (columns).

    Uses VECTOR (circular) averaging to handle the 0/360° discontinuity.

    Algorithm
    ---------
    1. Convert each direction θ to unit vector:  u = cos(θ),  v = sin(θ).
    2. Compute mean u and mean v for each cell.
    3. Recover mean angle:  θ_mean = degrees(atan2(mean_v, mean_u)) % 360.

    Color scale: twilight (cyclic, perceptually uniform for angular data).
    """
    tmp = wdf.copy()
    rad      = np.deg2rad(tmp["wind_direction"])
    tmp["_u"] = np.cos(rad)
    tmp["_v"] = np.sin(rad)

    u_pivot = tmp.pivot_table(values="_u", index="dayofyear", columns="hour", aggfunc="mean")
    v_pivot = tmp.pivot_table(values="_v", index="dayofyear", columns="hour", aggfunc="mean")

    # Align on shared index / columns (important when month-filtered)
    u_pivot, v_pivot = u_pivot.align(v_pivot, join="inner")

    dir_deg   = np.degrees(np.arctan2(v_pivot.values, u_pivot.values)) % 360
    dir_pivot = pd.DataFrame(dir_deg, index=u_pivot.index, columns=u_pivot.columns)

    fig = px.imshow(
        dir_pivot,
        labels      = dict(x="Hour of Day", y="Day of Year", color="Direction°"),
        title       = "Wind Direction – Day × Hour",
        color_continuous_scale = "twilight",
        range_color = [0, 360],
        aspect      = "auto",
        origin      = "lower",
    )
    fig.update_layout(
        height   = 400,
        template = "plotly_white",
        xaxis    = dict(title="Hour of Day", tickmode="linear", dtick=3),
        yaxis    = dict(title="Day of Year"),
        coloraxis_colorbar=dict(
            title    = "Direction",
            tickvals = [0, 90, 180, 270, 360],
            ticktext = ["N 0°", "E 90°", "S 180°", "W 270°", "N 360°"],
        ),
        # title = dict(x=0.5),
    )
    return fig


# ─── Plot: Speed Histogram ────────────────────────────────────────────────────

def plot_speed_histogram(wdf: pd.DataFrame) -> go.Figure:
    """Wind speed frequency histogram.

    Bins match the wind rose speed tiers so the two charts are consistent.
    Y-axis shows % of total hours.
    """
    total = len(wdf)
    labels, pcts = [], []

    for i in range(len(_SPEED_BINS) - 1):
        lo, hi  = _SPEED_BINS[i], _SPEED_BINS[i + 1]
        count   = int(((wdf["wind_speed"] >= lo) & (wdf["wind_speed"] < hi)).sum())
        labels.append(_SPEED_LABELS[i])
        pcts.append(count / total * 100.0 if total > 0 else 0.0)

    fig = go.Figure(go.Bar(
        x            = labels,
        y            = pcts,
        marker_color = _SPEED_COLORS[: len(labels)],
        text         = [f"{p:.1f}%" for p in pcts],
        textposition = "outside",
    ))
    fig.update_layout(
        title       = dict(text="Wind Speed Distribution", font_size=16),
        xaxis_title = "Wind Speed Bin (m/s)",
        yaxis       = dict(title="Frequency (%)", ticksuffix="%"),
        height      = 420,
        template    = "plotly_white",
        showlegend  = False,
        bargap      = 0.15,
    )
    return fig


# ─── Plot: Climate Bubble Chart ─────────────────────────────────────────────

# 12 visually distinct colours, one per month
_MONTH_COLORS = [
    "#e6194b", "#f58231", "#ffe119", "#bfef45",
    "#3cb44b", "#42d4f4", "#4363d8", "#911eb4",
    "#f032e6", "#a9a9a9", "#9a6324", "#469990",
]


def plot_climate_bubble(wdf: pd.DataFrame) -> go.Figure:
    """Temperature–Humidity–Wind bubble chart.

    Each point = one EPW hour.

    Axes / visual encoding
    ----------------------
    X   : Dry Bulb Temperature (°C)
    Y   : Relative Humidity (%)
    Size: Wind Speed (m/s) – larger bubble = stronger wind
    Color: Month – seasonal context

    Interpretation guide
    --------------------
    Large bubble at high T & high RH  → hot & humid but breezy
                                        (natural ventilation potential)
    Small bubble at high T & high RH  → hot, humid, stagnant
                                        (most uncomfortable)
    Medium/large bubble near comfort   → near-comfortable with breeze

    Algorithm notes
    ---------------
    • A floor of +0.3 is added to wind_speed before sizing so that calm
      winds still render as tiny visible dots.
    • sizemode="area" keeps perception proportional (area ∝ speed).
    • sizeref is computed from the actual data maximum so the largest
      bubble never exceeds ~28 px in diameter.
    """
    needed = {"dry_bulb_temperature", "relative_humidity", "wind_speed", "month"}
    if not needed.issubset(wdf.columns):
        missing = needed - set(wdf.columns)
        fig = go.Figure()
        fig.add_annotation(
            text=f"Missing columns for bubble chart: {missing}",
            xref="paper", yref="paper", x=0.5, y=0.5,
            showarrow=False, font_size=13,
        )
        fig.update_layout(height=500)
        return fig

    tmp = wdf.dropna(
        subset=["dry_bulb_temperature", "relative_humidity", "wind_speed"]
    ).copy()

    # Size floor: calm winds show as tiny dots, not invisible ones
    tmp["_bubble_size"] = tmp["wind_speed"] + 0.3
    # sizeref so the very largest bubble ≈ 28 px diameter (area mode:
    # displayed_diameter ∝ sqrt(value/sizeref)  →  sizeref = value / (px/2)²)
    max_bubble = float(tmp["_bubble_size"].max())
    sizeref    = 2.0 * max_bubble / (28.0 ** 2) if max_bubble > 0 else 0.01

    fig = go.Figure()

    for m in range(1, 13):
        mdata = tmp[tmp["month"] == m]
        if mdata.empty:
            continue
        mname = _MONTH_NAMES[m - 1]
        fig.add_trace(go.Scatter(
            x          = mdata["dry_bulb_temperature"],
            y          = mdata["relative_humidity"],
            mode       = "markers",
            name       = mname,
            customdata = mdata["wind_speed"].values,
            marker     = dict(
                size      = mdata["_bubble_size"].values,
                sizemode  = "area",
                sizeref   = sizeref,
                sizemin   = 2,
                color     = _MONTH_COLORS[(m - 1) % len(_MONTH_COLORS)],
                opacity   = 0.55,
                line      = dict(width=0),
            ),
            hovertemplate=(
                f"<b>{mname}</b><br>"
                "Temp: %{x:.1f}°C<br>"
                "RH: %{y:.0f}%<br>"
                "Wind: %{customdata:.1f} m/s"
                "<extra></extra>"
            ),
        ))

    fig.update_layout(
        title  = dict(
            text      = "Temperature – Humidity – Wind Speed",
            # x         = 0.5,
            font_size = 16,
            font_color= "#2c3e50",
        ),
        xaxis  = dict(
            title     = "Dry Bulb Temperature (°C)",
            gridcolor = "rgba(200,200,200,0.4)",
        ),
        yaxis  = dict(
            title     = "Relative Humidity (%)",
            range     = [0, 105],
            gridcolor = "rgba(200,200,200,0.4)",
        ),
        legend = dict(
            title       = "Month",
            orientation = "v",
            font_size   = 11,
        ),
        height   = 560,
        template = "plotly_white",
        annotations=[dict(
            text      = "Bubble size = wind speed (m/s)",
            xref="paper", yref="paper",
            x=0.01, y=0.99,
            showarrow = False,
            font      = dict(size=11, color="#888"),
            align     = "left",
        )],
    )
    return fig


# ─── Wind statistics ──────────────────────────────────────────────────────────

def compute_wind_statistics(wdf: pd.DataFrame) -> dict:
    """Compute prevailing-wind summary statistics.

    Returns
    -------
    dict with keys:
        prevailing_direction  – most frequent non-calm direction label
        mean_speed            – mean wind speed (m/s) over all hours
        max_speed             – peak wind speed (m/s)
        calm_percent          – % of calm hours
        strongest_direction   – direction label for the peak-speed hour
    """
    total = len(wdf)
    if total == 0:
        return {}

    calm_pct = float(wdf["is_calm"].sum()) / total * 100.0
    active   = wdf[~wdf["is_calm"]]

    if active.empty:
        return dict(
            prevailing_direction = "N/A",
            mean_speed           = 0.0,
            max_speed            = 0.0,
            calm_percent         = calm_pct,
            strongest_direction  = "N/A",
        )

    prevailing    = active["direction_label"].value_counts().idxmax()
    max_idx       = wdf["wind_speed"].idxmax()
    strongest_dir = (
        wdf.at[max_idx, "direction_label"]
        if "direction_label" in wdf.columns
        else "N/A"
    )

    return dict(
        prevailing_direction = prevailing,
        mean_speed           = float(wdf["wind_speed"].mean()),
        max_speed            = float(wdf["wind_speed"].max()),
        calm_percent         = calm_pct,
        strongest_direction  = strongest_dir,
    )


# ─── KPI card helper ──────────────────────────────────────────────────────────

def _kpi_card(label: str, value: str, color: str) -> str:
    return (
        f'<div style="background:white;padding:14px;border-radius:8px;'
        f'border-left:4px solid {color};'
        f'box-shadow:0 2px 4px rgba(0,0,0,0.08);text-align:center;">'
        f'<div style="font-size:11px;font-weight:700;color:{color};'
        f'text-transform:uppercase;letter-spacing:0.5px;margin-bottom:6px;">{label}</div>'
        f'<div style="font-size:22px;font-weight:700;color:#2c3e50;">{value}</div>'
        f'</div>'
    )


# ─── Main entry point ─────────────────────────────────────────────────────────

def render_wind_analysis(
    epw_df: pd.DataFrame,
    months: list = None,
    n_sectors: int = 16,
    exclude_calm: bool = False,
) -> None:
    """Render the Wind Analysis dashboard.

    Called from pages/analysis.py inside ``col_right``.
    All controls (month filter via date range, direction sectors, options toggle)
    live in the left panel of analysis.py and are passed in as parameters.

    Parameters
    ----------
    epw_df       : Full parsed EPW DataFrame (8760 rows).
    months       : Month numbers to include (1–12); None = all 12.
    n_sectors    : Compass sector count for the wind rose.
    exclude_calm : When True, normalise frequencies by non-calm hours only.
    """
    st.markdown(
        '<h3>Wind Analysis</h3>',
        unsafe_allow_html=True,
    )

    # ── Validate required columns ────────────────────────────────────────────
    required = {"wind_speed", "wind_direction"}
    missing  = required - set(epw_df.columns)
    if missing:
        st.error(
            f"EPW dataframe is missing columns required for wind analysis: "
            f"{', '.join(sorted(missing))}. "
            "Please ensure epw_parser.py extracts "
            "wind_speed (EPW field 21) and wind_direction (EPW field 20)."
        )
        return

    # ── Normalise months ─────────────────────────────────────────────────────
    if not months:
        months = list(range(1, 13))

    # ════════ COMPUTE ═════════════════════════════════════════════════════════
    with st.spinner("Computing wind statistics…"):
        wdf = prepare_wind_data(epw_df, months=months, n_sectors=n_sectors)

    if wdf.empty:
        st.warning("No wind data available for the selected date range.")
        return

    rose_df, calm_pct = compute_wind_rose(wdf, n_sectors=n_sectors, exclude_calm=exclude_calm)
    stats             = compute_wind_statistics(wdf)

    # ════════ CHARTS (full width of col_right) ════════════════════════════════
    # ── 1. Wind Rose ──────────────────────────────────────────────────────────
    st.plotly_chart(
        plot_wind_rose(rose_df, calm_pct, n_sectors),
        use_container_width=True,
    )

    # ── 2. Wind Speed Heatmap (full width) ────────────────────────────────────
    st.plotly_chart(plot_speed_heatmap(wdf), use_container_width=True)

    # ── 3. Wind Direction Heatmap (full width) ────────────────────────────────
    st.plotly_chart(plot_direction_heatmap(wdf), use_container_width=True)

    # ── 4. Speed Distribution Histogram ──────────────────────────────────────
    st.plotly_chart(plot_speed_histogram(wdf), use_container_width=True)

    # ── 5. Temperature – Humidity – Wind Bubble Chart ─────────────────────────
    st.plotly_chart(plot_climate_bubble(wdf), use_container_width=True)

    # ── 6. Prevailing Wind Statistics ─────────────────────────────────────────
    st.markdown(
        '<div style="font-size:16px;font-weight:700;padding-bottom:6px;margin:20px 0 12px;">'
        "Prevailing Wind Statistics</div>",
        unsafe_allow_html=True,
    )

    if stats:
        sc1, sc2, sc3, sc4, sc5 = st.columns(5)
        with sc1:
            st.markdown(
                _kpi_card("Prevailing Dir.", stats["prevailing_direction"], "#3b82f6"),
                unsafe_allow_html=True,
            )
        with sc2:
            st.markdown(
                _kpi_card("Mean Speed", f"{stats['mean_speed']:.2f} m/s", "#8b5cf6"),
                unsafe_allow_html=True,
            )
        with sc3:
            st.markdown(
                _kpi_card("Max Speed", f"{stats['max_speed']:.2f} m/s", "#ef4444"),
                unsafe_allow_html=True,
            )
        with sc4:
            st.markdown(
                _kpi_card("Calm Hours", f"{stats['calm_percent']:.1f}%", "#f59e0b"),
                unsafe_allow_html=True,
            )
        with sc5:
            st.markdown(
                _kpi_card("Strongest Dir.", stats["strongest_direction"], "#06b6d4"),
                unsafe_allow_html=True,
            )
