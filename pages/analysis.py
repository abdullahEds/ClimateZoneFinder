"""Climate Analytics Dashboard – thin orchestrator.

All heavy lifting is delegated to pages/modules/:
  epw_parser      → parse_epw
  ppt_report      → generate_pptx_report, generate_shading_pptx_report
  sun_path        → render_sun_path_section
  dbt_module      → calculate_ashrae_comfort, render (Temperature tabs)
  humidity_module → render (Humidity tabs)
"""

import sys
import os

# Ensure pages/modules/ is importable when Streamlit runs this script
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import base64
import pandas as pd
import streamlit as st

from modules.epw_parser import parse_epw
from modules.ppt_report import generate_pptx_report, generate_shading_pptx_report
from modules.sun_path import render_sun_path_section
from modules.dbt_module import calculate_ashrae_comfort
from modules import dbt_module, humidity_module, wind_module

# ─── Page configuration ───────────────────────────────────────────────────────

st.set_page_config(
    page_title="Climate Analytics Dashboard",
    layout="wide",
)


def get_base64_image(image_path):
    try:
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    except Exception:
        try:
            with open(f"../{image_path}", "rb") as img_file:
                return base64.b64encode(img_file.read()).decode()
        except Exception:
            return ""


logo_base64 = get_base64_image("images/EDSlogo.jpg")

# ─── Header ───────────────────────────────────────────────────────────────────

col_logo, col_title, col_home = st.columns([1, 4, 1])

with col_logo:
    st.markdown(
        f'<img src="data:image/png;base64,{logo_base64}" style="height: 80px; margin-top: 45px;">',
        unsafe_allow_html=True,
    )

with col_title:
    st.markdown(
        '<h2 style="text-align: center; color: #a85c42; margin-top: 45px;">'
        "Climate Analytics Dashboard</h2>",
        unsafe_allow_html=True,
    )

st.markdown(
    """
    <style>
    div[data-testid="stHorizontalBlock"]:first-of-type {
        border-bottom: 1px solid #e6e6e6;
        padding-bottom: 20px;
        margin-bottom: 0px;
        background-color: white;
    }
    div[data-testid="stHorizontalBlock"]:first-of-type img:hover {
        transform: scale(1.05);
        opacity: 0.85;
        transition: all 0.3s ease;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ─── Session state defaults ───────────────────────────────────────────────────

if "active_tab" not in st.session_state:
    st.session_state.active_tab = "Annual Trend"

# ─── Global CSS ───────────────────────────────────────────────────────────────

st.markdown(
    """
    <style>
    .block-container { padding-top: 0rem !important; }
    section[data-testid="stSidebar"] { display: none !important; }
    button[kind="header"] { display: none !important; }
    .main .block-container {
        max-width: 100% !important;
        padding-left: 1.5rem !important;
        padding-right: 1.5rem !important;
    }
    .control-section-header {
        font-size: 15px; font-weight: 700; color: #2c3e50;
        text-transform: uppercase; letter-spacing: 0.5px;
        margin-bottom: 12px; margin-top: 16px;
        display: flex; align-items: center; gap: 8px; width: 200px;
    }
    .control-section-header:first-child { margin-top: 0; }
    [data-testid="fileUploadDropzone"] {
        border-radius: 6px !important; border-color: #cbd5e0 !important;
    }
    .stSlider   { margin-top: 12px; }
    .stDateInput { margin-top: 12px; }
    .stSelectbox { margin-top: 8px; }
    .section-title {
        font-size: 18px; font-weight: 700; color: #2c3e50;
        margin-bottom: 16px; border-bottom: 2px solid #3498db;
        padding-bottom: 8px; display: inline-block;
    }
    .kpi-card {
        background: white; padding: 16px; border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.08); text-align: center;
    }
    .kpi-label {
        font-size: 11px; font-weight: 700; color: #718096;
        text-transform: uppercase; letter-spacing: 0.5px;
        margin-bottom: 8px; opacity: 0.9;
    }
    .kpi-value { font-size: 26px; font-weight: 700; color: #2c3e50; margin: 8px 0; }
    .kpi-meta  { font-size: 11px; color: #718096; opacity: 0.85; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ─── Main layout ──────────────────────────────────────────────────────────────

col_left, col_right = st.columns([0.85, 2.15], gap="small")

with col_left:
    st.write("##### 📤 Upload EPW File")
    uploaded = st.file_uploader("", type=["epw"], label_visibility="collapsed", width=300)

    st.write("##### Module")
    selected_parameter = st.selectbox(
        "Select parameter",
        ["Temperature", "Humidity", "Sun Path", "Wind"],
        label_visibility="collapsed",
        key="parameter_selector",
        width=300,
    )

if uploaded is None:
    with col_left:
        st.info("Please upload an .epw file to analyze.", width=300)
    st.stop()

# ─── Parse EPW and render left-panel controls ─────────────────────────────────

try:
    raw = uploaded.getvalue().decode("utf-8", errors="replace")
    df, metadata = parse_epw(raw)

    # Derived date fields used by charts
    df["doy"]        = df["datetime"].dt.dayofyear
    df["day"]        = df["datetime"].dt.day
    df["month"]      = df["datetime"].dt.month
    df["month_name"] = df["datetime"].dt.strftime("%b")

    with col_left:
        if selected_parameter == "Sun Path":
            hour_range = (0, 23)

            st.markdown('<div class="control-section-header">🌡️ Overheating Thresholds</div>', unsafe_allow_html=True)
            st.number_input(
                "Temperature Threshold (°C)", value=28.0, step=0.5,
                key="temp_threshold",
                help="Hours with dry-bulb temperature above this value are flagged",
                width=300,
            )
            st.number_input(
                "Radiation Threshold (W/m²)", value=315.0, step=10.0,
                key="rad_threshold",
                help="Hours with Global Horizontal Irradiance above this value are flagged",
                width=300,
            )

            st.markdown('<div class="control-section-header">📍 Location</div>', unsafe_allow_html=True)
            st.number_input(
                "Latitude (°)",  value=float(metadata.get("latitude")  or 0.0),
                step=0.1, key="shading_lat", help="Auto-read from EPW metadata", width=300,
            )
            st.number_input(
                "Longitude (°)", value=float(metadata.get("longitude") or 0.0),
                step=0.1, key="shading_lon", help="Auto-read from EPW metadata", width=300,
            )

            st.markdown('<div class="control-section-header">📐 Design Parameters</div>', unsafe_allow_html=True)
            st.number_input(
                "Design Cutoff Angle (°)", value=45.0, step=1.0,
                min_value=5.0, max_value=89.0, key="design_cutoff_angle",
                help="Solar altitude cutoff used to assess shading device protection",
                width=300,
            )
        elif selected_parameter == "Wind":
            hour_range = (0, 23)

            st.markdown('<div class="control-section-header">🧭 Direction Sectors</div>', unsafe_allow_html=True)
            st.select_slider(
                "Sectors",
                options=[4, 8, 12, 16, 24, 32, 36],
                value=16,
                key="wind_n_sectors",
                label_visibility="collapsed",
                help="Number of compass sectors in the wind rose (8 or 16 recommended)",
                width=300,
            )

            st.markdown('<div class="control-section-header">⚙️ Wind Options</div>', unsafe_allow_html=True)
            st.toggle(
                "Renormalise by non-calm hours",
                value=False,
                key="wind_exclude_calm",
                help=(
                    "Off: bar frequencies = % of total hours. "
                    "On: % of non-calm hours only."
                ),
            )
        else:
            st.markdown('<div class="control-section-header">⏰ Time Range (Hours)</div>', unsafe_allow_html=True)
            hour_range = st.slider(
                "Select hours (start - end)", min_value=0, max_value=23,
                value=(8, 18), step=1, key="hour_range",
                label_visibility="collapsed", width=300,
            )

        # ── Date range ─────────────────────────────────────────────────────────
        st.markdown('<div class="control-section-header">📅 Date Range</div>', unsafe_allow_html=True)

        months_list = [
            "January","February","March","April","May","June",
            "July","August","September","October","November","December",
        ]
        if "start_month_idx" not in st.session_state:
            st.session_state.start_month_idx = 0
        if "end_month_idx" not in st.session_state:
            st.session_state.end_month_idx = 11

        month_col1, month_col2, _ = st.columns([1, 1, 0.5], gap="small")

        with month_col1:
            start_month = st.selectbox(
                "From", options=range(len(months_list)),
                format_func=lambda x: months_list[x],
                key="start_month_select", label_visibility="collapsed", width=150,
            )
            st.session_state.start_month_idx = start_month

        with month_col2:
            end_month_options = list(range(start_month, len(months_list)))
            if st.session_state.end_month_idx < start_month:
                st.session_state.end_month_idx = start_month
            end_month = st.selectbox(
                "To", options=end_month_options,
                format_func=lambda x: months_list[x],
                key="end_month_select",
                index=min(st.session_state.end_month_idx - start_month, len(end_month_options) - 1),
                label_visibility="collapsed", width=150,
            )
            st.session_state.end_month_idx = end_month

        # ── PowerPoint report download ─────────────────────────────────────────
        st.markdown('<div class="control-section-header">📊 Report (PowerPoint)</div>', unsafe_allow_html=True)

        _active_chart  = st.session_state.get("sun_chart_type", "Sun Path")
        _is_shading    = selected_parameter == "Sun Path" and _active_chart == "Shading"

        try:
            _year = df["datetime"].dt.year.iloc[0] if not df.empty else 2024
            _s_num = st.session_state.start_month_idx + 1
            _e_num = st.session_state.end_month_idx + 1
            _start_date = pd.to_datetime(f"{_year}-{_s_num}-01").date()
            _end_date   = (
                pd.to_datetime(f"{_year}-12-31").date()
                if _e_num == 12
                else (pd.to_datetime(f"{_year}-{_e_num+1}-01") - pd.Timedelta(days=1)).date()
            )
            _sh, _eh = st.session_state.get("hour_range", (8, 18))

            if _is_shading:
                _tz_str = metadata.get("timezone", "UTC")
                shading_bytes = generate_shading_pptx_report(
                    df, metadata,
                    temp_threshold=float(st.session_state.get("temp_threshold", 28.0)),
                    rad_threshold=float(st.session_state.get("rad_threshold", 315.0)),
                    lat=float(st.session_state.get("shading_lat",  metadata.get("latitude")  or 0.0)),
                    lon=float(st.session_state.get("shading_lon",  metadata.get("longitude") or 0.0)),
                    tz_str=_tz_str,
                    design_cutoff_angle=float(st.session_state.get("design_cutoff_angle", 45.0)),
                )
                st.download_button(
                    label="⬇️ Download Shading Report",
                    data=shading_bytes,
                    file_name="Shading_Analysis_Report.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    key="download_shading_report",
                    width=300,
                )
            else:
                report_bytes = generate_pptx_report(
                    df, _start_date, _end_date, _sh, _eh,
                    selected_parameter, metadata=metadata,
                )
                st.download_button(
                    label="⬇️ Download Climate Report",
                    data=report_bytes,
                    file_name=(
                        f"Climate_Analysis_Report_"
                        f"{_start_date.strftime('%Y%m%d')}_to_{_end_date.strftime('%Y%m%d')}.pptx"
                    ),
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    key="download_report",
                    width=300,
                )
        except Exception as _e:
            st.error(f"❌ Failed to generate report: {_e}")

except Exception as e:
    with col_left:
        st.error(f"❌ Failed to parse EPW: {e}")
    st.stop()

# ─── Right panel ──────────────────────────────────────────────────────────────

with col_right:
    st.markdown(
        """
        <style>
        .tab-container {
            display: flex; gap: 0; background-color: #f8f9fa; padding: 0;
            margin: -1rem -1rem 1.5rem -1rem; border-bottom: 2px solid #e9ecef;
        }
        .tab-button {
            padding: 12px 24px; cursor: pointer; font-size: 14px; font-weight: 600;
            color: #495057; background-color: #f8f9fa; border: none;
            border-bottom: 3px solid transparent; transition: all 0.3s ease;
        }
        .tab-button:hover { background-color: #e9ecef; }
        .tab-button.active {
            color: #2c3e50; border-bottom-color: #3498db; background-color: white;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    if selected_parameter == "Sun Path":
        render_sun_path_section(df, metadata)

    elif selected_parameter == "Wind":
        _s = st.session_state.start_month_idx + 1
        _e = st.session_state.end_month_idx + 1
        _wind_months    = list(range(_s, _e + 1))
        _wind_n_sectors = st.session_state.get("wind_n_sectors", 16)
        _wind_excl_calm = st.session_state.get("wind_exclude_calm", False)
        wind_module.render_wind_analysis(
            df,
            months       = _wind_months,
            n_sectors    = _wind_n_sectors,
            exclude_calm = _wind_excl_calm,
        )

    else:
        # Tab navigation buttons
        tc1, tc2, tc3, tc4, tc5 = st.columns(5, gap="small")
        with tc1:
            if st.button("Annual Trend",    key="tab_annual",   use_container_width=True):
                st.session_state.active_tab = "Annual Trend"
        with tc2:
            if st.button("Monthly Trend",   key="tab_monthly",  use_container_width=True):
                st.session_state.active_tab = "Monthly Trend"
        with tc3:
            if st.button("Diurnal Profile", key="tab_diurnal",  use_container_width=True):
                st.session_state.active_tab = "Diurnal Profile"
        with tc4:
            if st.button("Comfort Analysis",key="tab_comfort",  use_container_width=True):
                st.session_state.active_tab = "Comfort Analysis"
        with tc5:
            if st.button("Energy Metrics",  key="tab_energy",   use_container_width=True):
                st.session_state.active_tab = "Energy Metrics"

        # Resolve date / time range
        year = df["datetime"].dt.year.iloc[0] if not df.empty else 2024
        start_month_num = st.session_state.start_month_idx + 1
        end_month_num   = st.session_state.end_month_idx + 1

        start_date = pd.to_datetime(f"{year}-{start_month_num}-01").date()
        end_date   = (
            pd.to_datetime(f"{year}-12-31").date()
            if end_month_num == 12
            else (pd.to_datetime(f"{year}-{end_month_num+1}-01") - pd.Timedelta(days=1)).date()
        )
        start_hour, end_hour = st.session_state.get("hour_range", (8, 18))

        # Compute daily statistics (shared by both modules)
        daily_stats = df.groupby("doy").agg(
            temp_min=("dry_bulb_temperature", "min"),
            temp_max=("dry_bulb_temperature", "max"),
            temp_avg=("dry_bulb_temperature", "mean"),
            rh_min  =("relative_humidity",    "min"),
            rh_max  =("relative_humidity",    "max"),
            rh_avg  =("relative_humidity",    "mean"),
        ).reset_index()

        daily_stats["datetime"] = pd.to_datetime(
            daily_stats["doy"].astype(str) + f"-{year}", format="%j-%Y", errors="coerce"
        )
        daily_stats["datetime_display"] = daily_stats["datetime"].dt.strftime("%b %d")

        # Merge ASHRAE adaptive comfort bands
        c80lo, c80hi, c90lo, c90hi = calculate_ashrae_comfort(df)
        comfort_df = pd.DataFrame({
            "doy":             c80lo.index,
            "comfort_80_lower": c80lo.values,
            "comfort_80_upper": c80hi.values,
            "comfort_90_lower": c90lo.values,
            "comfort_90_upper": c90hi.values,
        })
        daily_stats = daily_stats.merge(comfort_df, on="doy", how="left")

        active_tab = st.session_state.active_tab

        if selected_parameter == "Temperature":
            dbt_module.render(
                df, daily_stats, active_tab,
                start_date, end_date, start_hour, end_hour,
            )
        elif selected_parameter == "Humidity":
            humidity_module.render(
                df, daily_stats, active_tab,
                start_date, end_date, start_hour, end_hour,
            )

# ─── Footer ───────────────────────────────────────────────────────────────────

st.markdown("<br><br>", unsafe_allow_html=True)
st.image("images/EDS-footer.png", width=2000)
st.markdown(
    """
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.4/css/all.min.css">
    <div style="text-align:center; font-size:14px;">
        Email: <a href="mailto:info@edsglobal.com">info@edsglobal.com</a>   |
        Phone: +91 . 11 . 4056 8633   |
        <a href="https://twitter.com/edsglobal?lang=en" target="_blank">
            <i class="fab fa-twitter" style="color:#1DA1F2; margin:0 6px;"></i></a>
        <a href="https://www.facebook.com/Environmental.Design.Solutions/" target="_blank">
            <i class="fab fa-facebook" style="color:#4267B2; margin:0 6px;"></i></a>
        <a href="https://www.instagram.com/eds_global/?hl=en" target="_blank">
            <i class="fab fa-instagram" style="color:#E1306C; margin:0 6px;"></i></a>
        <a href="https://www.linkedin.com/company/environmental-design-solutions/" target="_blank">
            <i class="fab fa-linkedin" style="color:#0077b5; margin:0 6px;"></i></a>
    </div>
    """,
    unsafe_allow_html=True,
)
