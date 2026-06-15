"""Solar PV Potential module — SolarGIS country-level monthly data.

Loads the World Bank / SolarGIS PVOUT Level 1 dataset
(solargis_country_pv_data.xlsx, "Monthly data" sheet) at startup and
converts the daily averages (kWh/kWp/day) to monthly totals (kWh/kWp/month).

No EPW data is required — epw_df and metadata are accepted for API
compatibility with pages/analysis.py but are not used.

Exposes:
    render(epw_df, metadata)  ← called from pages/analysis.py
"""

import pathlib
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ─── Colour palette ───────────────────────────────────────────────────────────
_C_PRIMARY = "#f59e0b"
_C_HIGH    = "#10b981"
_C_LOW     = "#ef4444"
_C_BORDER  = "#f97316"
_C_TEXT    = "#2c3e50"

_MONTHS     = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
               "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

_EXCEL_PATH = pathlib.Path(__file__).parents[2] / "solargis_country_pv_data.xlsx"


# ─── Data loading ─────────────────────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def _load_data() -> pd.DataFrame:
    """Load and clean the SolarGIS monthly sheet.

    Returns a DataFrame with columns:
        country, iso_a3, region, yearly_daily,
        jan … dec  (monthly totals in kWh/kWp/month)
    """
    raw = pd.read_excel(
        _EXCEL_PATH,
        sheet_name="Monthly data",
        header=None,
    )

    # Row 1 contains the column names; rows 2+ are data
    raw.columns = raw.iloc[1]
    raw = raw.iloc[2:].reset_index(drop=True)

    month_cols = ["January", "February", "March", "April", "May", "June",
                  "July", "August", "September", "October", "November", "December"]

    df = pd.DataFrame()
    df["country"]      = raw["Country or region"].astype(str).str.strip()
    df["iso_a3"]       = raw["ISO_A3"].astype(str).str.strip()
    df["region"]       = raw["World Bank Region"].fillna("Other").astype(str).str.strip()
    df["yearly_daily"] = pd.to_numeric(raw["Yearly"], errors="coerce")

    # Keep raw daily averages (kWh/kWp/day) per month
    for col in month_cols:
        df[col.lower()[:3]] = pd.to_numeric(raw[col], errors="coerce")

    df = df.dropna(subset=["yearly_daily"]).reset_index(drop=True)
    return df


# ─── Helpers ──────────────────────────────────────────────────────────────────

def _kpi(label: str, value: str, sub: str, color: str) -> str:
    return (
        f'<div style="background:white;padding:16px 12px;border-radius:8px;'
        f'border-left:4px solid {color};'
        f'box-shadow:0 2px 6px rgba(0,0,0,0.08);text-align:center;">'
        f'<div style="font-size:11px;font-weight:700;color:{color};'
        f'text-transform:uppercase;letter-spacing:0.5px;margin-bottom:6px;">{label}</div>'
        f'<div style="font-size:24px;font-weight:700;color:{_C_TEXT};">{value}</div>'
        f'<div style="font-size:11px;color:#64748b;margin-top:4px;">{sub}</div>'
        f'</div>'
    )


def _build_chart(values: list, peak_idx: int, low_idx: int, country: str) -> go.Figure:
    colors = [
        _C_HIGH if i == peak_idx else _C_LOW if i == low_idx else _C_PRIMARY
        for i in range(12)
    ]

    fig = go.Figure(go.Bar(
        x=_MONTHS,
        y=values,
        marker_color=colors,
        marker_line_width=0,
        text=[f"{v:.2f}" for v in values],
        textposition="outside",
        textfont=dict(size=11, color=_C_TEXT),
        hovertemplate="<b>%{x}</b><br>%{y:.2f} kWh/kWp.day<extra></extra>",
    ))

    fig.update_layout(
        title=dict(
            text=f"Monthly Solar PV Specific Yield — {country}",
            font=dict(size=16, color=_C_TEXT),
            x=0,
        ),
        xaxis=dict(title="Month", tickfont=dict(size=12)),
        yaxis=dict(
            title="Specific Yield (kWh / kWp / day)",
            range=[0, max(values) * 1.2],
            gridcolor="#f1f5f9",
        ),
        height=420,
        template="plotly_white",
        showlegend=False,
        plot_bgcolor="white",
        margin=dict(t=60, b=60, l=70, r=20),
    )

    for color, label, xp in [
        (_C_HIGH, "Highest Generation", 0.70),
        (_C_LOW,  "Lowest Generation",  0.86),
    ]:
        fig.add_annotation(
            xref="paper", yref="paper",
            x=xp, y=1.07,
            text=f"<span style='color:{color}'>■</span> {label}",
            showarrow=False,
            font=dict(size=11, color=_C_TEXT),
            align="left",
        )

    return fig


# ─── Main entry point ─────────────────────────────────────────────────────────

def render() -> None:
    """Render the Solar PV Potential dashboard (no EPW data required)."""
    st.markdown(
        f'<h3 style="color:{_C_TEXT};margin-bottom:4px;">Solar PV Potential</h3>',
        unsafe_allow_html=True,
    )

    # ── Load data ─────────────────────────────────────────────────────────────
    try:
        df = _load_data()
    except Exception as exc:
        st.error(f"Could not load SolarGIS data: {exc}")
        return

    countries = sorted(df["country"].tolist())
    default   = "United Arab Emirates" if "United Arab Emirates" in countries else countries[0]

    selected = st.selectbox(
        "Select Country",
        options=countries,
        index=countries.index(default),
        help="SolarGIS PVOUT Level 1 — long-term monthly average specific yield.",
    )

    row    = df[df["country"] == selected].iloc[0]
    values = [float(row[m]) for m in ["jan", "feb", "mar", "apr", "may", "jun",
                                       "jul", "aug", "sep", "oct", "nov", "dec"]]

    annual_daily = row["yearly_daily"]
    peak_idx     = values.index(max(values))
    low_idx      = values.index(min(values))

    # ── KPI row ───────────────────────────────────────────────────────────────
    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)
    k1, k2, k3 = st.columns(3)

    with k1:
        st.markdown(
            _kpi("Annual Yield", f"{annual_daily:.2f}", "kWh / kWp.day (long-term avg)", _C_BORDER),
            unsafe_allow_html=True,
        )
    with k2:
        st.markdown(
            _kpi("Highest Generation", _MONTHS[peak_idx],
                 f"{values[peak_idx]:.2f} kWh/kWp.day", _C_HIGH),
            unsafe_allow_html=True,
        )
    with k3:
        st.markdown(
            _kpi("Lowest Generation", _MONTHS[low_idx],
                 f"{values[low_idx]:.2f} kWh/kWp.day", _C_LOW),
            unsafe_allow_html=True,
        )

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

    # ── Bar chart ─────────────────────────────────────────────────────────────
    st.plotly_chart(
        _build_chart(values, peak_idx, low_idx, selected),
        use_container_width=True,
    )

    # ── Data source note ──────────────────────────────────────────────────────
    iso    = row["iso_a3"]
    region = row["region"]

    st.markdown(
        f"""
        <div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:6px;
                    padding:12px 16px;margin-top:8px;font-size:12px;color:#64748b;">
        <strong>Data Source:</strong> World Bank Group — Global Solar Atlas 2.0,
        powered by <strong>SolarGIS</strong>. PVOUT Level 1 long-term average
        practical photovoltaic power output (kWh/kWp.day) for
        <strong>{selected}</strong> (ISO&nbsp;{iso}, {region} region).
        <br><br>
        <em>Monthly bars show the long-term average daily specific yield for each
        month. Actual yield varies with system tilt, shading, inverter losses, and
        local microclimate.</em>
        </div>
        """,
        unsafe_allow_html=True,
    )
