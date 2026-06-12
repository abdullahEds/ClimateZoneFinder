"""Solar PV Potential Module — site analysis + Solargis country benchmarking.

Data sources
------------
1. EPW file (site-specific, hourly):
   global_horizontal_irradiance (GHI, Wh/m²) — used to compute actual site PVOUT.

2. Solargis Global PV Potential Country Rankings 2020 (bundled Excel):
   Average practical potential (PVOUT Level 1, kWh/kWp/day), long-term monthly
   data for 209 countries + country indicators (GHI, LCOE, Seasonality, Installed
   capacity, electricity tariffs).

Key formulae
------------
Daily GHI (kWh/m²/day)    = Σ hourly_GHI_Wh / 1000  (per day)
Monthly avg GHI            = mean of daily GHI values for the month
Peak Sun Hours (PSH/day)   = monthly_avg_GHI  (numerically equal at 1 kW/m² STC)
Site PVOUT (kWh/kWp/day)   = PSH × PR   (PR = performance ratio, default 0.80)
Monthly yield (kWh)        = site_PVOUT × system_kWp × days_in_month
Annual specific yield      = Σ monthly_yield / system_kWp   (kWh/kWp/year)
System size (kWp)          = daily_demand_kWh / site_PVOUT_annual
Roof area (m²)             = required_kWp × 6.5  (≈400 W panels at ~15% efficiency)
Simple payback (years)     = (kWp × cost_per_kWp_USD) / (annual_kWh × tariff_USD)

Performance Ratio (PR) reference values
----------------------------------------
Hot desert climate (>5.0 kWh/m²/day):  0.77  (high temp losses)
Sunny warm climate (4.0–5.0):           0.80
Temperate climate  (3.0–4.0):           0.83
Cool / cloudy      (<3.0):              0.86

Exposes
-------
    render(epw_df, metadata)  ← called from pages/analysis.py
"""

import calendar
import os
import pathlib

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import streamlit as st

# ─── Constants ────────────────────────────────────────────────────────────────

_MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
           "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

_PANEL_AREA_PER_KWP = 6.5          # m² per kWp (standard 400 W panels ~15% eff.)
_DEFAULT_PR        = 0.80           # Performance Ratio default
_DEFAULT_COST_USD  = 700            # USD per kWp installed cost default
_DEFAULT_DEMAND    = 50.0           # kWh/day default daily demand
_CO2_FACTORS = {                    # kgCO2/kWh grid emission factors
    "IND": 0.82, "ARE": 0.48, "SAU": 0.65, "QAT": 0.49,
    "KWT": 0.65, "OMN": 0.57, "PAK": 0.42, "AUS": 0.73,
    "DEU": 0.37, "GBR": 0.23, "USA": 0.39, "CHN": 0.62,
    "default": 0.50,
}

# World Bank region full names
_REGION_NAMES = {
    "AFR": "Sub-Saharan Africa",
    "EAP": "East Asia & Pacific",
    "ECA": "Europe & Central Asia",
    "LCR": "Latin America & Caribbean",
    "MENA": "Middle East & North Africa",
    "SOA": "South Asia",
    "Other": "Other / High Income",
}

# EPW country abbreviation → ISO_A3 mapping (common EPW country strings)
_EPW_TO_ISO = {
    # India
    "IND": "IND", "India": "IND", "IN": "IND",
    # UAE
    "ARE": "ARE", "UAE": "ARE", "United Arab Emirates": "ARE",
    # Saudi Arabia
    "SAU": "SAU", "Saudi Arabia": "SAU",
    # Pakistan
    "PAK": "PAK", "Pakistan": "PAK",
    # Kuwait
    "KWT": "KWT", "Kuwait": "KWT",
    # Oman
    "OMN": "OMN", "Oman": "OMN",
    # Qatar
    "QAT": "QAT", "Qatar": "QAT",
    # Bahrain
    "BHR": "BHR", "Bahrain": "BHR",
    # Jordan
    "JOR": "JOR", "Jordan": "JOR",
    # Egypt
    "EGY": "EGY", "Egypt": "EGY",
    # Australia
    "AUS": "AUS", "Australia": "AUS",
    # Germany
    "DEU": "DEU", "Germany": "DEU",
    # USA
    "USA": "USA", "US": "USA", "United States": "USA",
    # UK
    "GBR": "GBR", "UK": "GBR", "United Kingdom": "GBR",
    # China
    "CHN": "CHN", "China": "CHN",
    # Singapore
    "SGP": "SGP", "Singapore": "SGP",
    # Bangladesh
    "BGD": "BGD", "Bangladesh": "BGD",
    # Sri Lanka
    "LKA": "LKA", "Sri Lanka": "LKA",
    # Nepal
    "NPL": "NPL", "Nepal": "NPL",
    # Morocco
    "MAR": "MAR", "Morocco": "MAR",
    # Nigeria
    "NGA": "NGA", "Nigeria": "NGA",
    # South Africa
    "ZAF": "ZAF", "South Africa": "ZAF",
}


# ─── Data loading ─────────────────────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def _load_solargis_data() -> tuple[pd.DataFrame, pd.DataFrame]:
    """Load and clean Solargis country PV data. Cached after first load."""
    base = pathlib.Path(__file__).parents[3]
    xlsx = base / "solargis_country_pv_data.xlsx"
    if not xlsx.exists():
        # Fallback: look in cwd
        xlsx = pathlib.Path("solargis_country_pv_data.xlsx")
    if not xlsx.exists():
        return pd.DataFrame(), pd.DataFrame()

    # ── Monthly PVOUT ────────────────────────────────────────────────────────
    raw_monthly = pd.read_excel(xlsx, sheet_name="Monthly data", header=None)
    monthly = raw_monthly.iloc[2:].copy().reset_index(drop=True)
    monthly.columns = ["ISO_A3", "Country", "Note", "Region",
                       "Yearly", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
                       "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    for col in ["Yearly", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
                "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]:
        monthly[col] = pd.to_numeric(monthly[col], errors="coerce")
    monthly = monthly.dropna(subset=["ISO_A3", "Yearly"]).reset_index(drop=True)

    # ── Country indicators ───────────────────────────────────────────────────
    raw_ci = pd.read_excel(xlsx, sheet_name="Country indicators", header=None)
    ci = raw_ci.iloc[2:].copy().reset_index(drop=True)
    ci.columns = [
        "ISO_A3", "Country", "Note", "Region",
        "Population", "Total_area_km2", "Evaluated_area_km2", "Level1_area_pct",
        "HDI", "GDP_per_capita",
        "GHI_avg", "PVOUT_avg", "LCOE_USD_kWh", "Seasonality_index",
        "PV_equiv_area_pct", "Installed_PV_MWp", "Installed_PV_Wp_capita",
        "Rural_electricity_pct", "Elec_consumption_kWh_capita",
        "Supply_reliability", "Electricity_tariff_USD_cent",
    ]
    numeric_ci = ["GHI_avg", "PVOUT_avg", "LCOE_USD_kWh", "Seasonality_index",
                  "PV_equiv_area_pct", "Installed_PV_MWp", "GDP_per_capita",
                  "Electricity_tariff_USD_cent", "HDI", "Population"]
    for col in numeric_ci:
        ci[col] = pd.to_numeric(ci[col], errors="coerce")
    ci = ci.dropna(subset=["ISO_A3", "PVOUT_avg"]).reset_index(drop=True)

    return monthly, ci


def _resolve_iso(metadata: dict) -> str | None:
    """Try to resolve ISO_A3 code from EPW metadata."""
    for key in ("country", "city", "location"):
        val = metadata.get(key, "") or ""
        iso = _EPW_TO_ISO.get(val.strip())
        if iso:
            return iso
        # Try 3-char prefix of city/country string
        iso = _EPW_TO_ISO.get(val.strip()[:3])
        if iso:
            return iso
    return None


# ─── Compute functions ────────────────────────────────────────────────────────

def compute_site_solar(df: pd.DataFrame, pr: float = _DEFAULT_PR) -> pd.DataFrame:
    """Compute monthly GHI, PSH, and PVOUT from EPW hourly data.

    Returns a DataFrame indexed by month (1-12) with columns:
        ghi_daily       — mean daily GHI (kWh/m²/day)
        psh             — peak sun hours (= ghi_daily at 1 kW/m² STC)
        pvout           — site PVOUT (kWh/kWp/day) = psh × pr
        days            — days in month
    """
    d = df.copy()
    if "month" not in d.columns:
        d["month"] = d["datetime"].dt.month
    if "global_horizontal_irradiance" not in d.columns:
        return pd.DataFrame()

    # Daily GHI sum (Wh/m²) → kWh/m²
    d["date"] = d["datetime"].dt.date
    daily = d.groupby("date")["global_horizontal_irradiance"].sum() / 1000.0
    daily_df = daily.reset_index()
    daily_df["month"] = pd.to_datetime(daily_df["date"]).dt.month

    monthly = daily_df.groupby("month")["global_horizontal_irradiance"].mean().reset_index()
    monthly.columns = ["month", "ghi_daily"]
    monthly["psh"] = monthly["ghi_daily"]
    monthly["pvout"] = monthly["psh"] * pr
    monthly["days"] = monthly["month"].apply(lambda m: calendar.monthrange(2024, m)[1])
    return monthly.set_index("month")


def compute_annual_yield(
    site: pd.DataFrame,
    system_kWp: float,
) -> dict:
    """Compute monthly and annual PV yield for a given system size."""
    site = site.copy()
    site["monthly_yield_kWh"] = site["pvout"] * system_kWp * site["days"]
    annual_yield = site["monthly_yield_kWh"].sum()
    specific_yield = annual_yield / system_kWp  # kWh/kWp/year
    pvout_annual = site["pvout"].mean()          # kWh/kWp/day avg
    capacity_factor = annual_yield / (system_kWp * 8760) * 100  # %
    return {
        "monthly": site[["monthly_yield_kWh"]].copy(),
        "annual_kWh": annual_yield,
        "specific_yield_kWh_kWp": specific_yield,
        "pvout_annual": pvout_annual,
        "capacity_factor_pct": capacity_factor,
    }


def compute_system_sizing(
    daily_demand_kWh: float,
    site: pd.DataFrame,
    pr: float = _DEFAULT_PR,
) -> dict:
    """Size a PV system to meet a given daily demand."""
    pvout_annual = site["pvout"].mean()
    if pvout_annual <= 0:
        return {}
    required_kWp = daily_demand_kWh / pvout_annual
    roof_area_m2 = required_kWp * _PANEL_AREA_PER_KWP
    annual_yield = required_kWp * pvout_annual * 365
    return {
        "required_kWp": required_kWp,
        "roof_area_m2": roof_area_m2,
        "annual_yield_kWh": annual_yield,
        "pvout_annual": pvout_annual,
    }


def compute_economics(
    system_kWp: float,
    annual_yield_kWh: float,
    cost_per_kWp_usd: float,
    tariff_usd_per_kwh: float,
    iso_code: str = None,
) -> dict:
    """Simple financial and CO2 metrics."""
    total_cost_usd = system_kWp * cost_per_kWp_usd
    annual_savings_usd = annual_yield_kWh * tariff_usd_per_kwh
    payback_years = total_cost_usd / annual_savings_usd if annual_savings_usd > 0 else None
    co2_factor = _CO2_FACTORS.get(iso_code, _CO2_FACTORS["default"])
    annual_co2_kg = annual_yield_kWh * co2_factor
    lifetime_co2_t = annual_co2_kg * 25 / 1000  # 25-year lifetime
    return {
        "total_cost_usd": total_cost_usd,
        "annual_savings_usd": annual_savings_usd,
        "payback_years": payback_years,
        "annual_co2_kg": annual_co2_kg,
        "lifetime_co2_t": lifetime_co2_t,
        "lcoe_usd": total_cost_usd / (annual_yield_kWh * 25) if annual_yield_kWh > 0 else None,
    }


def _pr_from_ghi(ghi_daily: float) -> float:
    """Suggest a performance ratio based on average daily GHI."""
    if ghi_daily >= 5.0:
        return 0.77
    elif ghi_daily >= 4.0:
        return 0.80
    elif ghi_daily >= 3.0:
        return 0.83
    else:
        return 0.86


# ─── KPI card helper ──────────────────────────────────────────────────────────

def _kpi(label: str, value: str, sub: str = "", color: str = "#A85C42") -> str:
    return f"""
    <div style="background:#fff;border-radius:8px;padding:14px 18px;border-left:4px solid {color};
                box-shadow:0 1px 4px rgba(0,0,0,0.08);min-height:80px;">
      <div style="font-size:11px;color:#888;text-transform:uppercase;letter-spacing:0.5px;">{label}</div>
      <div style="font-size:22px;font-weight:700;color:#1a1a1a;margin:4px 0;">{value}</div>
      <div style="font-size:11px;color:#999;">{sub}</div>
    </div>"""


def _pvout_color(pvout: float) -> str:
    if pvout >= 5.0:   return "#27ae60"
    elif pvout >= 4.5: return "#2ecc71"
    elif pvout >= 4.0: return "#f39c12"
    elif pvout >= 3.5: return "#e67e22"
    else:              return "#e74c3c"


def _pvout_label(pvout: float) -> str:
    if pvout >= 5.0:   return "Excellent"
    elif pvout >= 4.5: return "Very Good"
    elif pvout >= 4.0: return "Good"
    elif pvout >= 3.5: return "Moderate"
    else:              return "Low"


# ─── Tab renderers ────────────────────────────────────────────────────────────

def _render_solar_resource(site: pd.DataFrame, iso: str, monthly_sol: pd.DataFrame, ci: pd.DataFrame):
    """Tab 1: Site solar resource from EPW."""
    pvout_annual = site["pvout"].mean()
    ghi_annual = site["ghi_daily"].mean()
    psh_peak = site["psh"].max()
    psh_min = site["psh"].min()
    seasonality = round(psh_peak / psh_min, 2) if psh_min > 0 else None

    # KPIs
    kpi_cols = st.columns(4)
    kpis = [
        ("Annual Avg PVOUT", f"{pvout_annual:.2f} kWh/kWp/day",
         _pvout_label(pvout_annual), _pvout_color(pvout_annual)),
        ("Annual Avg GHI", f"{ghi_annual:.2f} kWh/m²/day",
         "Global Horizontal Irradiance", "#3498db"),
        ("Peak Sun Hours", f"{psh_peak:.2f} hrs/day",
         f"Best month", "#e67e22"),
        ("Seasonality Index", f"{seasonality:.2f}" if seasonality else "N/A",
         "Peak / trough ratio", "#9b59b6"),
    ]
    for col, (label, val, sub, color) in zip(kpi_cols, kpis):
        col.markdown(_kpi(label, val, sub, color), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Monthly GHI + PVOUT bar chart ────────────────────────────────────────
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(
        x=_MONTHS, y=[site.loc[m, "ghi_daily"] for m in range(1, 13)],
        name="GHI (kWh/m²/day)", marker_color="#F4B942", opacity=0.85,
        hovertemplate="<b>%{x}</b><br>GHI: %{y:.2f} kWh/m²/day<extra></extra>"
    ), secondary_y=False)
    fig.add_trace(go.Scatter(
        x=_MONTHS, y=[site.loc[m, "pvout"] for m in range(1, 13)],
        name="Site PVOUT (kWh/kWp/day)", mode="lines+markers",
        line=dict(color="#A85C42", width=2.5), marker=dict(size=7),
        hovertemplate="<b>%{x}</b><br>PVOUT: %{y:.2f} kWh/kWp/day<extra></extra>"
    ), secondary_y=True)

    # Solargis country benchmark overlay if available
    if iso and not monthly_sol.empty:
        row = monthly_sol[monthly_sol["ISO_A3"] == iso]
        if not row.empty:
            row = row.iloc[0]
            bench_vals = [row[m] for m in _MONTHS]
            fig.add_trace(go.Scatter(
                x=_MONTHS, y=bench_vals,
                name="Solargis Country Avg", mode="lines",
                line=dict(color="#95a5a6", width=1.5, dash="dot"),
                hovertemplate="<b>%{x}</b><br>Country avg: %{y:.2f} kWh/kWp/day<extra></extra>"
            ), secondary_y=True)

    fig.update_layout(
        title="Monthly Solar Resource — GHI & PVOUT",
        height=400, legend=dict(orientation="h", y=-0.15),
        plot_bgcolor="#fafafa", paper_bgcolor="white",
        margin=dict(t=50, b=60, l=50, r=50),
    )
    fig.update_yaxes(title_text="GHI (kWh/m²/day)", secondary_y=False, gridcolor="#eee")
    fig.update_yaxes(title_text="PVOUT (kWh/kWp/day)", secondary_y=True)
    st.plotly_chart(fig, use_container_width=True)

    # ── Irradiance heatmap (hour × month) ────────────────────────────────────
    st.markdown("#### Hourly Irradiance Pattern")
    d = site.reset_index()
    epw_ref = None
    # We need the original df here — handled via session_state trick below
    # The heatmap is rendered in render() where df is available


def _render_irradiance_heatmap(df: pd.DataFrame):
    """Render a 24×12 GHI heatmap (hour × month)."""
    d = df.copy()
    if "month" not in d.columns:
        d["month"] = d["datetime"].dt.month
    if "hour" not in d.columns:
        d["hour"] = d["datetime"].dt.hour

    pivot = d.groupby(["hour", "month"])["global_horizontal_irradiance"].mean().unstack("month")
    pivot.columns = _MONTHS
    pivot.index.name = "Hour"

    fig = go.Figure(go.Heatmap(
        z=pivot.values,
        x=_MONTHS,
        y=[f"{h:02d}:00" for h in range(24)],
        colorscale="YlOrRd",
        colorbar=dict(title="GHI (Wh/m²)"),
        hovertemplate="Hour: %{y}<br>Month: %{x}<br>GHI: %{z:.0f} Wh/m²<extra></extra>",
    ))
    fig.update_layout(
        title="Average Hourly GHI by Month",
        height=420, margin=dict(t=50, b=40, l=60, r=40),
        paper_bgcolor="white",
    )
    st.plotly_chart(fig, use_container_width=True)


def _render_yield(site: pd.DataFrame, iso: str):
    """Tab 2: System yield calculator."""
    st.markdown("#### System Yield Estimator")
    col1, col2, col3 = st.columns(3)
    system_kWp = col1.number_input("System Size (kWp)", min_value=0.5, max_value=50000.0,
                                    value=100.0, step=0.5, key="pv_sys_kwp")
    pr = col2.number_input("Performance Ratio", min_value=0.50, max_value=0.95,
                            value=_pr_from_ghi(site["ghi_daily"].mean()), step=0.01,
                            key="pv_pr", help="0.77–0.80 for hot climates, 0.83–0.86 for cool climates")
    tilt_note = col3.selectbox("Module Type", ["Crystalline Silicon (Standard)",
                                                "Thin Film (Amorphous Si)",
                                                "Bifacial (Mono-PERC)"],
                                key="pv_module_type")

    # Recompute site with adjusted PR
    site_adj = site.copy()
    site_adj["pvout"] = site_adj["psh"] * pr
    result = compute_annual_yield(site_adj, system_kWp)

    # KPIs
    kpi_cols = st.columns(4)
    kpi_data = [
        ("Annual Yield", f"{result['annual_kWh']:,.0f} kWh", f"for {system_kWp:.0f} kWp system", "#27ae60"),
        ("Specific Yield", f"{result['specific_yield_kWh_kWp']:,.0f} kWh/kWp", "per year", "#2980b9"),
        ("Avg Daily Output", f"{result['annual_kWh']/365:,.1f} kWh/day", "annual average", "#e67e22"),
        ("Capacity Factor", f"{result['capacity_factor_pct']:.1f}%", "annual", "#9b59b6"),
    ]
    for col, (label, val, sub, color) in zip(kpi_cols, kpi_data):
        col.markdown(_kpi(label, val, sub, color), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Monthly yield bar chart
    monthly_yield = result["monthly"].reset_index()
    monthly_yield["month_name"] = monthly_yield["month"].apply(lambda m: _MONTHS[m - 1])

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=monthly_yield["month_name"],
        y=monthly_yield["monthly_yield_kWh"],
        marker_color="#A85C42", opacity=0.85,
        text=[f"{v:,.0f}" for v in monthly_yield["monthly_yield_kWh"]],
        textposition="outside", textfont=dict(size=10),
        hovertemplate="<b>%{x}</b><br>Yield: %{y:,.0f} kWh<extra></extra>",
    ))
    fig.update_layout(
        title=f"Monthly PV Yield — {system_kWp:.0f} kWp System",
        yaxis_title="Energy Yield (kWh)",
        height=380, plot_bgcolor="#fafafa", paper_bgcolor="white",
        margin=dict(t=50, b=40),
    )
    st.plotly_chart(fig, use_container_width=True)

    # Cumulative yield line
    monthly_yield["cumulative_kWh"] = monthly_yield["monthly_yield_kWh"].cumsum()
    fig2 = go.Figure(go.Scatter(
        x=monthly_yield["month_name"],
        y=monthly_yield["cumulative_kWh"],
        fill="tozeroy", line=dict(color="#A85C42", width=2),
        hovertemplate="<b>%{x}</b><br>Cumulative: %{y:,.0f} kWh<extra></extra>",
    ))
    fig2.update_layout(
        title="Cumulative Annual Yield",
        yaxis_title="Cumulative Energy (kWh)",
        height=280, plot_bgcolor="#fafafa", paper_bgcolor="white",
        margin=dict(t=50, b=40),
    )
    st.plotly_chart(fig2, use_container_width=True)


def _render_sizing(site: pd.DataFrame):
    """Tab 3: System sizing from demand."""
    st.markdown("#### System Sizing Calculator")
    col1, col2, col3 = st.columns(3)
    daily_demand = col1.number_input("Daily Energy Demand (kWh/day)", min_value=1.0,
                                      max_value=100000.0, value=_DEFAULT_DEMAND, step=1.0,
                                      key="pv_demand")
    fraction = col2.slider("PV Coverage (%)", min_value=25, max_value=100, value=80,
                            key="pv_coverage",
                            help="Percentage of daily demand to be met by PV")
    pr_size = col3.number_input("Performance Ratio", min_value=0.50, max_value=0.95,
                                 value=_pr_from_ghi(site["ghi_daily"].mean()),
                                 step=0.01, key="pv_pr_size")

    target_demand = daily_demand * fraction / 100
    site_adj = site.copy()
    site_adj["pvout"] = site_adj["psh"] * pr_size
    result = compute_system_sizing(target_demand, site_adj, pr_size)

    if not result:
        st.error("Cannot size system — check EPW solar data.")
        return

    kWp = result["required_kWp"]
    area = result["roof_area_m2"]

    # KPI row
    kpi_cols = st.columns(4)
    kpi_data = [
        ("Required System Size", f"{kWp:.1f} kWp", f"Meeting {fraction}% of demand", "#A85C42"),
        ("Roof Area Needed", f"{area:,.0f} m²", "≈400 W panels @ 15% efficiency", "#e67e22"),
        ("Annual PV Generation", f"{result['annual_yield_kWh']:,.0f} kWh", "estimated", "#27ae60"),
        ("Site PVOUT", f"{result['pvout_annual']:.2f} kWh/kWp/day", "annual average", "#2980b9"),
    ]
    for col, (label, val, sub, color) in zip(kpi_cols, kpi_data):
        col.markdown(_kpi(label, val, sub, color), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Demand coverage breakdown chart
    pv_supply = result["annual_yield_kWh"]
    total_demand_annual = daily_demand * 365
    pv_covered = min(pv_supply, total_demand_annual)
    grid_import = max(0, total_demand_annual - pv_covered)

    fig = go.Figure()
    fig.add_trace(go.Bar(name="PV Generation", x=["Annual Energy"],
                          y=[pv_supply], marker_color="#F4B942"))
    fig.add_trace(go.Bar(name="PV Covers Demand", x=["Annual Energy"],
                          y=[pv_covered], marker_color="#27ae60"))
    fig.add_trace(go.Bar(name="Grid Import", x=["Annual Energy"],
                          y=[grid_import], marker_color="#e74c3c"))
    fig.update_layout(barmode="group", height=320,
                       title="Annual Energy Balance",
                       yaxis_title="Energy (kWh)",
                       paper_bgcolor="white", plot_bgcolor="#fafafa",
                       margin=dict(t=50, b=40))
    st.plotly_chart(fig, use_container_width=True)

    # System size sensitivity table
    st.markdown("#### Sensitivity: System Size vs Coverage")
    sizes = [daily_demand * f / 100 / result["pvout_annual"]
             for f in [25, 50, 75, 100, 125]]
    fracs = [25, 50, 75, 100, 125]
    tbl = pd.DataFrame({
        "Coverage (%)": fracs,
        "System Size (kWp)": [f"{s:.1f}" for s in sizes],
        "Roof Area (m²)": [f"{s * _PANEL_AREA_PER_KWP:,.0f}" for s in sizes],
        "Annual Yield (kWh)": [f"{s * result['pvout_annual'] * 365:,.0f}" for s in sizes],
    })
    st.dataframe(tbl, use_container_width=True, hide_index=True)


def _render_benchmark(iso: str, monthly_sol: pd.DataFrame, ci: pd.DataFrame,
                       site: pd.DataFrame):
    """Tab 4: Country and regional benchmark."""
    if ci.empty or monthly_sol.empty:
        st.warning("Solargis benchmark data not available.")
        return

    if site.empty or "pvout" not in site.columns:
        site_pvout = None
    else:
        site_pvout = site["pvout"].mean()

    # Get country row
    country_row = ci[ci["ISO_A3"] == iso] if iso else pd.DataFrame()
    country_name = country_row.iloc[0]["Country"] if not country_row.empty else "Your Site"
    country_pvout = float(country_row.iloc[0]["PVOUT_avg"]) if not country_row.empty else None
    country_ghi   = float(country_row.iloc[0]["GHI_avg"]) if not country_row.empty else None
    country_lcoe  = float(country_row.iloc[0]["LCOE_USD_kWh"]) if not country_row.empty else None
    country_region = country_row.iloc[0]["Region"] if not country_row.empty else None
    country_installed = float(country_row.iloc[0]["Installed_PV_MWp"]) if not country_row.empty else None
    country_tariff = float(country_row.iloc[0]["Electricity_tariff_USD_cent"]) if not country_row.empty else None

    # Global percentile rank of this country
    global_sorted = ci["PVOUT_avg"].dropna().sort_values(ascending=False).reset_index(drop=True)
    if country_pvout:
        rank = int((global_sorted > country_pvout).sum()) + 1
        total = len(global_sorted)
        percentile_rank = round((1 - rank / total) * 100, 1)
    else:
        rank, total, percentile_rank = None, len(global_sorted), None

    # KPIs
    kpi_cols = st.columns(4)
    kpi_data = [
        ("Site PVOUT (EPW)", f"{site_pvout:.2f} kWh/kWp/day" if site_pvout else "N/A (no EPW)",
         _pvout_label(site_pvout) if site_pvout else "Upload EPW for site data",
         _pvout_color(site_pvout) if site_pvout else "#aaa"),
        ("Country Avg PVOUT", f"{country_pvout:.2f}" if country_pvout else "N/A",
         f"{country_name}", "#2980b9"),
        ("Global Rank", f"#{rank} of {total}" if rank else "N/A",
         f"Top {100-percentile_rank:.0f}% worldwide" if percentile_rank else "", "#e67e22"),
        ("LCOE (Country)", f"${country_lcoe:.3f}/kWh" if country_lcoe else "N/A",
         "Levelised cost of energy", "#27ae60"),
    ]
    for col, (label, val, sub, color) in zip(kpi_cols, kpi_data):
        col.markdown(_kpi(label, val, sub, color), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Monthly comparison: Site EPW vs Solargis country ─────────────────────
    if not country_row.empty and not site.empty:
        country_monthly = monthly_sol[monthly_sol["ISO_A3"] == iso]
        if not country_monthly.empty:
            c_row = country_monthly.iloc[0]
            bench_vals = [float(c_row[m]) for m in _MONTHS]
            site_vals = [site.loc[m, "pvout"] for m in range(1, 13)]

            fig = go.Figure()
            fig.add_trace(go.Bar(x=_MONTHS, y=site_vals, name="Site (EPW-derived)",
                                  marker_color="#A85C42", opacity=0.85,
                                  hovertemplate="<b>%{x}</b><br>Site: %{y:.2f}<extra></extra>"))
            fig.add_trace(go.Bar(x=_MONTHS, y=bench_vals, name=f"Solargis Country Avg ({country_name})",
                                  marker_color="#5B9BD5", opacity=0.75,
                                  hovertemplate="<b>%{x}</b><br>Country avg: %{y:.2f}<extra></extra>"))
            fig.update_layout(
                barmode="group", title="Monthly PVOUT: Site vs Country Average",
                yaxis_title="PVOUT (kWh/kWp/day)", height=380,
                paper_bgcolor="white", plot_bgcolor="#fafafa",
                margin=dict(t=50, b=40), legend=dict(orientation="h", y=-0.15),
            )
            st.plotly_chart(fig, use_container_width=True)

    # ── Regional ranking bar chart ────────────────────────────────────────────
    if country_region:
        region_df = ci[ci["Region"] == country_region].dropna(subset=["PVOUT_avg"])
        region_df = region_df.sort_values("PVOUT_avg", ascending=True)
        colors = ["#A85C42" if r == iso else "#BDC3C7" for r in region_df["ISO_A3"]]

        fig2 = go.Figure(go.Bar(
            y=region_df["Country"].str[:20],
            x=region_df["PVOUT_avg"],
            orientation="h",
            marker_color=colors,
            hovertemplate="<b>%{y}</b><br>PVOUT: %{x:.2f} kWh/kWp/day<extra></extra>",
        ))
        fig2.update_layout(
            title=f"PVOUT Rankings — {_REGION_NAMES.get(country_region, country_region)}",
            xaxis_title="PVOUT (kWh/kWp/day)",
            height=max(350, len(region_df) * 22),
            margin=dict(t=50, b=40, l=160, r=40),
            paper_bgcolor="white", plot_bgcolor="#fafafa",
        )
        st.plotly_chart(fig2, use_container_width=True)

    # ── Global distribution ───────────────────────────────────────────────────
    st.markdown("#### Global PVOUT Distribution")
    fig3 = go.Figure()
    fig3.add_trace(go.Histogram(
        x=ci["PVOUT_avg"].dropna(), nbinsx=30,
        marker_color="#BDC3C7", name="All countries",
        hovertemplate="PVOUT: %{x:.2f}<br>Countries: %{y}<extra></extra>",
    ))
    if country_pvout:
        fig3.add_vline(x=country_pvout, line_color="#A85C42", line_width=2.5,
                       annotation_text=f"  {country_name} ({country_pvout:.2f})",
                       annotation_position="top right",
                       annotation_font_color="#A85C42")
    fig3.add_vline(x=ci["PVOUT_avg"].mean(), line_color="#2980b9", line_width=1.5,
                   line_dash="dot",
                   annotation_text=f"  Global avg ({ci['PVOUT_avg'].mean():.2f})",
                   annotation_position="top left")
    fig3.update_layout(
        title="Global Distribution of Country PVOUT Values",
        xaxis_title="PVOUT (kWh/kWp/day)",
        yaxis_title="Number of Countries",
        height=320, paper_bgcolor="white", plot_bgcolor="#fafafa",
        margin=dict(t=50, b=40),
    )
    st.plotly_chart(fig3, use_container_width=True)


def _render_economics(site: pd.DataFrame, iso: str, ci: pd.DataFrame):
    """Tab 5: Simple economic snapshot."""
    st.markdown("#### Economic Analysis")

    # Pull default tariff from Solargis data if available
    default_tariff = 0.10
    if iso and not ci.empty:
        row = ci[ci["ISO_A3"] == iso]
        if not row.empty:
            t = row.iloc[0]["Electricity_tariff_USD_cent"]
            if pd.notna(t) and t > 0:
                default_tariff = float(t) / 100  # convert cents to USD

    col1, col2, col3 = st.columns(3)
    system_kWp_eco = col1.number_input("System Size (kWp)", min_value=0.5,
                                        max_value=50000.0, value=100.0, step=0.5,
                                        key="pv_eco_kwp")
    cost_per_kWp = col2.number_input("Installed Cost (USD/kWp)", min_value=100.0,
                                      max_value=3000.0, value=float(_DEFAULT_COST_USD),
                                      step=50.0, key="pv_cost")
    tariff = col3.number_input("Electricity Tariff (USD/kWh)", min_value=0.01,
                                max_value=1.00, value=default_tariff, step=0.01,
                                key="pv_tariff",
                                help="Local retail electricity price")

    pr_eco = _pr_from_ghi(site["ghi_daily"].mean())
    site_adj = site.copy()
    site_adj["pvout"] = site_adj["psh"] * pr_eco
    yield_result = compute_annual_yield(site_adj, system_kWp_eco)
    econ = compute_economics(system_kWp_eco, yield_result["annual_kWh"],
                              cost_per_kWp, tariff, iso)

    # KPIs
    kpi_cols = st.columns(4)
    payback_str = f"{econ['payback_years']:.1f} years" if econ.get("payback_years") else "N/A"
    kpi_data = [
        ("Total System Cost", f"${econ['total_cost_usd']:,.0f}", f"{system_kWp_eco:.0f} kWp", "#e74c3c"),
        ("Annual Savings", f"${econ['annual_savings_usd']:,.0f}", "at given tariff", "#27ae60"),
        ("Simple Payback", payback_str, "without incentives", "#e67e22"),
        ("Project LCOE", f"${econ['lcoe_usd']:.4f}/kWh" if econ.get("lcoe_usd") else "N/A",
         "25-year lifetime", "#2980b9"),
    ]
    for col, (label, val, sub, color) in zip(kpi_cols, kpi_data):
        col.markdown(_kpi(label, val, sub, color), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # CO2 metrics
    co2_cols = st.columns(3)
    co2_cols[0].markdown(_kpi("Annual CO₂ Avoided",
                               f"{econ['annual_co2_kg']/1000:.1f} tonnes CO₂",
                               f"Grid factor: {_CO2_FACTORS.get(iso or '', _CO2_FACTORS['default'])} kgCO₂/kWh",
                               "#7f8c8d"), unsafe_allow_html=True)
    co2_cols[1].markdown(_kpi("25-Year CO₂ Offset",
                               f"{econ['lifetime_co2_t']:.0f} tonnes CO₂",
                               "Lifetime of system", "#7f8c8d"), unsafe_allow_html=True)
    equiv_cars = econ["annual_co2_kg"] / 1000 / 2.4  # avg car emits ~2.4 t CO2/year
    co2_cols[2].markdown(_kpi("Equivalent Cars Off Road",
                               f"{equiv_cars:.1f} cars/year",
                               "Annual equivalent", "#7f8c8d"), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Payback sensitivity chart
    st.markdown("#### Payback Period Sensitivity")
    tariff_range = np.linspace(0.05, 0.30, 20)
    paybacks = [(system_kWp_eco * cost_per_kWp) / (yield_result["annual_kWh"] * t)
                for t in tariff_range]
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=tariff_range, y=paybacks, mode="lines",
        fill="tozeroy", fillcolor="rgba(168,92,66,0.1)",
        line=dict(color="#A85C42", width=2),
        hovertemplate="Tariff: $%{x:.3f}/kWh<br>Payback: %{y:.1f} years<extra></extra>",
    ))
    fig.add_vline(x=tariff, line_dash="dot", line_color="#A85C42",
                  annotation_text=f"  Current tariff ${tariff:.3f}",
                  annotation_font_color="#A85C42")
    fig.add_hline(y=10, line_dash="dot", line_color="#27ae60",
                  annotation_text=" 10-yr target  ", annotation_position="right")
    fig.update_layout(
        xaxis_title="Electricity Tariff (USD/kWh)",
        yaxis_title="Simple Payback (years)",
        height=320, paper_bgcolor="white", plot_bgcolor="#fafafa",
        margin=dict(t=30, b=40),
    )
    st.plotly_chart(fig, use_container_width=True)


# ─── Main render ──────────────────────────────────────────────────────────────

def render(df: pd.DataFrame, metadata: dict) -> None:
    """Main entry point — renders the Solar PV module in Streamlit.

    Parameters
    ----------
    df : pd.DataFrame
        Parsed EPW DataFrame from epw_parser.parse_epw().
    metadata : dict
        Metadata dict from epw_parser.parse_epw() with keys:
        city, location, latitude, longitude, timezone.
    """
    # ── Load Solargis data ───────────────────────────────────────────────────
    monthly_sol, ci = _load_solargis_data()

    # ── Resolve country ISO code ─────────────────────────────────────────────
    iso = _resolve_iso(metadata)

    # Country selector override
    all_countries = ci[["ISO_A3", "Country"]].dropna().sort_values("Country")
    country_options = ["Auto-detect from EPW"] + [
        f"{row['Country']} ({row['ISO_A3']})" for _, row in all_countries.iterrows()
    ]
    default_idx = 0
    if iso:
        match = all_countries[all_countries["ISO_A3"] == iso]
        if not match.empty:
            label = f"{match.iloc[0]['Country']} ({iso})"
            if label in country_options:
                default_idx = country_options.index(label)

    with st.expander("🌍  Country / Benchmark Selection", expanded=(iso is None)):
        selected = st.selectbox(
            "Select country for Solargis benchmark",
            country_options, index=default_idx,
            key="pv_country_select",
            help="Auto-detect uses the country embedded in the EPW file header.",
        )
        if selected != "Auto-detect from EPW":
            iso = selected.split("(")[-1].rstrip(")")

    # ── Compute site solar resource ──────────────────────────────────────────
    pr_default = _pr_from_ghi(
        df.groupby(df["datetime"].dt.month)["global_horizontal_irradiance"]
        .mean().mean() / 1000 * 24 / 1000
        if "global_horizontal_irradiance" in df.columns else 4.0
    )
    epw_available = not df.empty and "global_horizontal_irradiance" in df.columns
    site = compute_site_solar(df, pr=pr_default) if epw_available else pd.DataFrame()
    if site.empty:
        # No EPW data — show only the benchmark tab
        st.info("☀️ No EPW file loaded. Showing country benchmark data only. "
                "Upload an EPW file for site-specific solar resource analysis.")
        tab_labels = ["🌍 Country Benchmark"]
        tabs = st.tabs(tab_labels)
        with tabs[0]:
            _render_benchmark(iso, monthly_sol, ci, pd.DataFrame())
        return

    # ── City / location banner ───────────────────────────────────────────────
    city = metadata.get("city") or metadata.get("location") or "Unknown Site"
    lat = metadata.get("latitude")
    lon = metadata.get("longitude")
    lat_str = f"{abs(lat):.2f}°{'N' if lat >= 0 else 'S'}" if lat else ""
    lon_str = f"{abs(lon):.2f}°{'E' if lon >= 0 else 'W'}" if lon else ""
    st.markdown(
        f"<div style='font-size:13px;color:#888;margin-bottom:8px;'>"
        f"📍 <b>{city}</b>  &nbsp;|&nbsp;  {lat_str}  {lon_str}</div>",
        unsafe_allow_html=True,
    )

    # ── Tabs ─────────────────────────────────────────────────────────────────
    tab_labels = ["☀️ Solar Resource", "⚡ System Yield", "📐 System Sizing",
                  "🌍 Country Benchmark", "💰 Economics"]
    tabs = st.tabs(tab_labels)

    with tabs[0]:
        _render_solar_resource(site, iso, monthly_sol, ci)
        _render_irradiance_heatmap(df)

    with tabs[1]:
        _render_yield(site, iso)

    with tabs[2]:
        _render_sizing(site)

    with tabs[3]:
        _render_benchmark(iso, monthly_sol, ci, site)

    with tabs[4]:
        _render_economics(site, iso, ci)
