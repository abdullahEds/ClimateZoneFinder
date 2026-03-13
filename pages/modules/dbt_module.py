"""Dry Bulb Temperature tab rendering module."""

import numpy as np
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st


def calculate_ashrae_comfort(df: pd.DataFrame) -> tuple:
    """
    Calculate ASHRAE adaptive comfort bands.

    Returns:
        (comfort_80_lower, comfort_80_upper, comfort_90_lower, comfort_90_upper)
        as daily rolling averages indexed by doy.
    """
    daily_avg = df.groupby("doy")["dry_bulb_temperature"].mean()
    comfort_line = daily_avg.rolling(window=7, center=True).mean()
    return (
        comfort_line - 3.5,
        comfort_line + 3.5,
        comfort_line - 2.5,
        comfort_line + 2.5,
    )


def render(
    df: pd.DataFrame,
    daily_stats: pd.DataFrame,
    active_tab: str,
    start_date,
    end_date,
    start_hour: int,
    end_hour: int,
) -> None:
    """Dispatch rendering based on the active tab."""
    if active_tab == "Annual Trend":
        _render_annual_trend(df, daily_stats, start_date, end_date, start_hour, end_hour)
    elif active_tab == "Monthly Trend":
        _render_monthly_trend(df, start_date, end_date)
    elif active_tab == "Diurnal Profile":
        _render_diurnal_profile(df, start_hour, end_hour)
    elif active_tab == "Comfort Analysis":
        _render_comfort_analysis(df, daily_stats, start_date, end_date)
    elif active_tab == "Energy Metrics":
        _render_energy_metrics(df, start_date, end_date, start_hour, end_hour)


# ─────────────────────────────────────────────────────────────────────────────


def _render_annual_trend(df, daily_stats, start_date, end_date, start_hour, end_hour):
    start_month_num = start_date.month
    end_month_num   = end_date.month

    start_doy = pd.to_datetime(f"2024-{start_month_num:02d}-01").dayofyear
    end_doy   = (
        366
        if end_month_num == 12
        else pd.to_datetime(f"2024-{end_month_num + 1:02d}-01").dayofyear - 1
    )

    fig = go.Figure()

    # ── Greyed-out: before selected range ─────────────────────────────────────
    if start_doy > 1:
        before = daily_stats[daily_stats["doy"] < start_doy]
        fig.add_trace(go.Scatter(x=before["datetime_display"], y=before["temp_max"],
                                 fill=None, mode="lines", line_color="rgba(100,100,100,0)",
                                 showlegend=False, hoverinfo="skip"))
        fig.add_trace(go.Scatter(x=before["datetime_display"], y=before["temp_min"],
                                 fill="tonexty", mode="lines", line_color="rgba(100,100,100,0)",
                                 fillcolor="rgba(180,180,180,0.15)",
                                 name="Unselected Period", showlegend=True, hoverinfo="skip"))
        fig.add_trace(go.Scatter(x=before["datetime_display"], y=before["temp_avg"],
                                 mode="lines", line=dict(color="rgba(150,150,150,0.4)", width=1, dash="dot"),
                                 showlegend=False, hoverinfo="skip"))

    # ── Active range ──────────────────────────────────────────────────────────
    active = daily_stats[(daily_stats["doy"] >= start_doy) & (daily_stats["doy"] <= end_doy)]

    # ASHRAE 80% band
    fig.add_trace(go.Scatter(x=active["datetime_display"], y=active["comfort_80_upper"],
                             fill=None, mode="lines", line_color="rgba(128,128,128,0)",
                             showlegend=False, hoverinfo="skip"))
    fig.add_trace(go.Scatter(x=active["datetime_display"], y=active["comfort_80_lower"],
                             fill="tonexty", mode="lines", line_color="rgba(128,128,128,0)",
                             name="ASHRAE adaptive comfort (80%)",
                             fillcolor="rgba(128,128,128,0.2)", hoverinfo="skip"))

    # ASHRAE 90% band
    fig.add_trace(go.Scatter(x=active["datetime_display"], y=active["comfort_90_upper"],
                             fill=None, mode="lines", line_color="rgba(128,128,128,0)",
                             showlegend=False, hoverinfo="skip"))
    fig.add_trace(go.Scatter(x=active["datetime_display"], y=active["comfort_90_lower"],
                             fill="tonexty", mode="lines", line_color="rgba(128,128,128,0)",
                             name="ASHRAE adaptive comfort (90%)",
                             fillcolor="rgba(128,128,128,0.4)", hoverinfo="skip"))

    # Temp min/max band
    fig.add_trace(go.Scatter(x=active["datetime_display"], y=active["temp_max"],
                             fill=None, mode="lines", line_color="rgba(255,0,0,0)",
                             showlegend=False, hoverinfo="skip"))
    fig.add_trace(go.Scatter(x=active["datetime_display"], y=active["temp_min"],
                             fill="tonexty", mode="lines", line_color="rgba(255,0,0,0)",
                             name="Dry bulb temperature Range",
                             fillcolor="rgba(255,173,173,0.4)",
                             customdata=active["temp_max"],
                             hovertemplate="<b>%{x}</b><br>Min: %{y:.2f}°C<br>Max: %{customdata:.2f}°C<extra></extra>"))

    # Average line
    fig.add_trace(go.Scatter(x=active["datetime_display"], y=active["temp_avg"],
                             mode="lines", name="Average Dry bulb temperature",
                             line=dict(color="#d32f2f", width=2),
                             hovertemplate="<b>%{x}</b><br>Avg: %{y:.2f}°C<extra></extra>"))

    # ── Greyed-out: after selected range ──────────────────────────────────────
    if end_doy < 365:
        after = daily_stats[daily_stats["doy"] > end_doy]
        if not after.empty:
            fig.add_trace(go.Scatter(x=after["datetime_display"], y=after["temp_max"],
                                     fill=None, mode="lines", line_color="rgba(100,100,100,0)",
                                     showlegend=False, hoverinfo="skip"))
            fig.add_trace(go.Scatter(x=after["datetime_display"], y=after["temp_min"],
                                     fill="tonexty", mode="lines", line_color="rgba(100,100,100,0)",
                                     fillcolor="rgba(180,180,180,0.15)", showlegend=False, hoverinfo="skip"))
            fig.add_trace(go.Scatter(x=after["datetime_display"], y=after["temp_avg"],
                                     mode="lines", line=dict(color="rgba(150,150,150,0.4)", width=1, dash="dot"),
                                     showlegend=False, hoverinfo="skip"))

    fig.update_layout(
        title="Annual Dry Bulb Temperature Trend",
        xaxis_title=None,
        yaxis_title="Temperature (°C)",
        hovermode="x unified",
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        xaxis_rangeslider_visible=False,
        height=450,
        template="plotly_white",
        margin=dict(b=80),
    )
    st.plotly_chart(fig, use_container_width=True)

    # ── KPI cards ─────────────────────────────────────────────────────────────
    filtered = df[
        (df["datetime"].dt.date >= start_date) &
        (df["datetime"].dt.date <= end_date) &
        (df["hour"].between(start_hour, end_hour))
    ]

    if filtered.empty:
        st.info(f"No data between {start_hour:02d}:00 and {end_hour:02d}:00 in the selected date range.")
        return

    min_row = filtered.loc[filtered["dry_bulb_temperature"].idxmin()]
    max_row = filtered.loc[filtered["dry_bulb_temperature"].idxmax()]
    temp_min  = min_row["dry_bulb_temperature"]
    temp_max  = max_row["dry_bulb_temperature"]
    temp_avg  = filtered["dry_bulb_temperature"].mean()
    diurnal   = temp_max - temp_min
    min_ds    = min_row["datetime"].strftime("%b %d")
    max_ds    = max_row["datetime"].strftime("%b %d")
    min_hr    = int(min_row["hour"])
    max_hr    = int(max_row["hour"])

    hdd18          = (18 - df["dry_bulb_temperature"]).clip(lower=0).sum()
    cdd24          = (df["dry_bulb_temperature"] - 24).clip(lower=0).sum()
    mean_t         = df["dry_bulb_temperature"].mean()
    comfort_hrs    = len(df[(df["dry_bulb_temperature"] >= mean_t - 3.5) &
                            (df["dry_bulb_temperature"] <= mean_t + 3.5)])
    comfort_80_pct = comfort_hrs / len(df) * 100
    cooling_1pct   = df["dry_bulb_temperature"].quantile(0.99)
    overheat_hrs   = len(df[df["dry_bulb_temperature"] > 28])
    cold_hrs       = len(df[df["dry_bulb_temperature"] < 12])

    def _card(label, value, sub, color):
        return f"""
<div style="background:white;padding:16px;border-radius:8px;border-left:4px solid {color};
            box-shadow:0 2px 4px rgba(0,0,0,0.08);text-align:center;">
  <div style="font-size:11px;font-weight:700;color:{color};text-transform:uppercase;letter-spacing:0.5px;">{label}</div>
  <div style="font-size:26px;font-weight:700;color:#2c3e50;margin:8px 0;">{value}</div>
  <div style="font-size:11px;color:#718096;">{sub}</div>
</div>"""

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: st.markdown(_card("Min Temp",       f"{temp_min:.2f} °C", f"{min_ds} · {min_hr:02d}:00", "#f59e0b"), unsafe_allow_html=True)
    with c2: st.markdown(_card("Max Temp",       f"{temp_max:.2f} °C", f"{max_ds} · {max_hr:02d}:00", "#ef4444"), unsafe_allow_html=True)
    with c3: st.markdown(_card("Avg Temp",       f"{temp_avg:.2f} °C", "All year average",             "#8b5cf6"), unsafe_allow_html=True)
    with c4: st.markdown(_card("Diurnal Range",  f"{diurnal:.2f} °C",  "",                             "#3b82f6"), unsafe_allow_html=True)
    with c5: st.markdown(_card("1% Cooling",     f"{cooling_1pct:.2f} °C", "",                         "#06b6d4"), unsafe_allow_html=True)

    c6, c7, c8, c9, c10 = st.columns(5)
    with c6:  st.markdown(_card("HDD18",         f"{hdd18:.0f}",         "",             "#dc2626"), unsafe_allow_html=True)
    with c7:  st.markdown(_card("CDD24",         f"{cdd24:.0f}",         "",             "#0891b2"), unsafe_allow_html=True)
    with c8:  st.markdown(_card("Comfort 80%",   f"{comfort_80_pct:.0f} %", "",          "#06b6d4"), unsafe_allow_html=True)
    with c9:  st.markdown(_card("Overheat Hrs",  f"{overheat_hrs}",      "",             "#8b5cf6"), unsafe_allow_html=True)
    with c10: st.markdown(_card("Cold Hrs",      f"{cold_hrs}",          "",             "#3b82f6"), unsafe_allow_html=True)


def _render_monthly_trend(df, start_date, end_date):
    monthly = df.groupby("month").agg(
        temp_min=("dry_bulb_temperature", "min"),
        temp_max=("dry_bulb_temperature", "max"),
        temp_avg=("dry_bulb_temperature", "mean"),
        rh_min=("relative_humidity", "min"),
        rh_max=("relative_humidity", "max"),
        rh_avg=("relative_humidity", "mean"),
    ).reset_index()

    month_lbl = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    monthly["month_name"] = monthly["month"].apply(lambda x: month_lbl[x - 1])

    start_month = start_date.month
    end_month   = end_date.month

    fig = go.Figure()

    if start_month > 1:
        before = monthly[monthly["month"] < start_month]
        fig.add_trace(go.Scatter(x=before["month_name"], y=before["temp_max"],
                                 fill=None, mode="lines", line_color="rgba(100,100,100,0)",
                                 showlegend=False, hoverinfo="skip"))
        fig.add_trace(go.Scatter(x=before["month_name"], y=before["temp_min"],
                                 fill="tonexty", mode="lines", line_color="rgba(100,100,100,0)",
                                 fillcolor="rgba(180,180,180,0.15)", name="Unselected Period",
                                 showlegend=True, hoverinfo="skip"))
        fig.add_trace(go.Scatter(x=before["month_name"], y=before["temp_avg"],
                                 mode="lines+markers",
                                 line=dict(color="rgba(150,150,150,0.4)", width=1, dash="dot"),
                                 marker=dict(size=4), showlegend=False, hoverinfo="skip"))

    active = monthly[(monthly["month"] >= start_month) & (monthly["month"] <= end_month)]

    fig.add_trace(go.Scatter(x=active["month_name"], y=active["temp_max"],
                             fill=None, mode="lines", line_color="rgba(255,0,0,0)",
                             showlegend=False, hoverinfo="skip"))
    fig.add_trace(go.Scatter(x=active["month_name"], y=active["temp_min"],
                             fill="tonexty", mode="lines", line_color="rgba(255,0,0,0)",
                             name="Monthly Temperature Range",
                             fillcolor="rgba(255,173,173,0.4)",
                             customdata=active["temp_max"],
                             hovertemplate="<b>%{x}</b><br>Min: %{y:.2f}°C<br>Max: %{customdata:.2f}°C<extra></extra>"))
    fig.add_trace(go.Scatter(x=active["month_name"], y=active["temp_avg"],
                             mode="lines+markers", name="Monthly Average Temperature",
                             line=dict(color="#d32f2f", width=2), marker=dict(size=8),
                             hovertemplate="<b>%{x}</b><br>Avg: %{y:.2f}°C<extra></extra>"))

    if end_month < 12:
        after = monthly[monthly["month"] > end_month]
        fig.add_trace(go.Scatter(x=after["month_name"], y=after["temp_max"],
                                 fill=None, mode="lines", line_color="rgba(100,100,100,0)",
                                 showlegend=False, hoverinfo="skip"))
        fig.add_trace(go.Scatter(x=after["month_name"], y=after["temp_min"],
                                 fill="tonexty", mode="lines", line_color="rgba(100,100,100,0)",
                                 fillcolor="rgba(180,180,180,0.15)", showlegend=False, hoverinfo="skip"))
        fig.add_trace(go.Scatter(x=after["month_name"], y=after["temp_avg"],
                                 mode="lines+markers",
                                 line=dict(color="rgba(150,150,150,0.4)", width=1, dash="dot"),
                                 marker=dict(size=4), showlegend=False, hoverinfo="skip"))

    fig.update_layout(
        title="Monthly Temperature Trend",
        xaxis_title="Month", yaxis_title="Temperature (°C)",
        hovermode="x unified", showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        height=450, template="plotly_white", margin=dict(b=80),
    )
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### Monthly Temperature Summary")
    kpi = monthly[["month_name", "temp_min", "temp_max", "temp_avg"]].copy()
    kpi.columns = ["Month", "Min (°C)", "Max (°C)", "Avg (°C)"]
    st.dataframe(kpi, use_container_width=True, hide_index=True,
                 column_config={
                     "Min (°C)": st.column_config.NumberColumn(format="%.2f"),
                     "Max (°C)": st.column_config.NumberColumn(format="%.2f"),
                     "Avg (°C)": st.column_config.NumberColumn(format="%.2f"),
                 })


def _render_diurnal_profile(df, start_hour, end_hour):
    hourly = df.groupby(["month", "hour"]).agg(
        temp_min=("dry_bulb_temperature", "min"),
        temp_max=("dry_bulb_temperature", "max"),
        temp_avg=("dry_bulb_temperature", "mean"),
    ).reset_index()

    avg = hourly.groupby("hour").agg(
        temp_min=("temp_min", "min"),
        temp_max=("temp_max", "max"),
        temp_avg=("temp_avg", "mean"),
    ).reset_index()

    fig = go.Figure()

    if start_hour > 0:
        before = avg[avg["hour"] < start_hour]
        fig.add_trace(go.Scatter(x=before["hour"], y=before["temp_max"],
                                 fill=None, mode="lines", line_color="rgba(100,100,100,0)",
                                 showlegend=False, hoverinfo="skip"))
        fig.add_trace(go.Scatter(x=before["hour"], y=before["temp_min"],
                                 fill="tonexty", mode="lines", line_color="rgba(100,100,100,0)",
                                 fillcolor="rgba(180,180,180,0.15)",
                                 name="Unselected Hours", showlegend=True, hoverinfo="skip"))
        fig.add_trace(go.Scatter(x=before["hour"], y=before["temp_avg"],
                                 mode="lines+markers",
                                 line=dict(color="rgba(150,150,150,0.4)", width=1, dash="dot"),
                                 marker=dict(size=4), showlegend=False, hoverinfo="skip"))

    active = avg[(avg["hour"] >= start_hour) & (avg["hour"] <= end_hour)]
    fig.add_trace(go.Scatter(x=active["hour"], y=active["temp_max"],
                             fill=None, mode="lines", line_color="rgba(255,0,0,0)",
                             showlegend=False, hoverinfo="skip"))
    fig.add_trace(go.Scatter(x=active["hour"], y=active["temp_min"],
                             fill="tonexty", mode="lines", line_color="rgba(255,0,0,0)",
                             name="Daily Range", fillcolor="rgba(255,173,173,0.3)",
                             customdata=active["temp_max"],
                             hovertemplate="<b>Hour %{x}:00</b><br>Min: %{y:.2f}°C<br>Max: %{customdata:.2f}°C<extra></extra>"))
    fig.add_trace(go.Scatter(x=active["hour"], y=active["temp_avg"],
                             mode="lines+markers", name="Average Temperature",
                             line=dict(color="#d32f2f", width=2), marker=dict(size=6),
                             hovertemplate="<b>Hour %{x}:00</b><br>Avg: %{y:.2f}°C<extra></extra>"))

    if end_hour < 23:
        after = avg[avg["hour"] > end_hour]
        fig.add_trace(go.Scatter(x=after["hour"], y=after["temp_max"],
                                 fill=None, mode="lines", line_color="rgba(100,100,100,0)",
                                 showlegend=False, hoverinfo="skip"))
        fig.add_trace(go.Scatter(x=after["hour"], y=after["temp_min"],
                                 fill="tonexty", mode="lines", line_color="rgba(100,100,100,0)",
                                 fillcolor="rgba(180,180,180,0.15)", showlegend=False, hoverinfo="skip"))
        fig.add_trace(go.Scatter(x=after["hour"], y=after["temp_avg"],
                                 mode="lines+markers",
                                 line=dict(color="rgba(150,150,150,0.4)", width=1, dash="dot"),
                                 marker=dict(size=4), showlegend=False, hoverinfo="skip"))

    fig.update_layout(
        title="Diurnal Temperature Profile",
        xaxis_title="Hour of Day", yaxis_title="Temperature (°C)",
        hovermode="x unified", showlegend=True,
        template="plotly_white", height=450,
    )
    st.plotly_chart(fig, use_container_width=True)


def _render_comfort_analysis(df, daily_stats, start_date, end_date):
    start_month_num = start_date.month
    end_month_num   = end_date.month

    start_doy = pd.to_datetime(f"2024-{start_month_num:02d}-01").dayofyear
    end_doy   = (
        366
        if end_month_num == 12
        else pd.to_datetime(f"2024-{end_month_num + 1:02d}-01").dayofyear - 1
    )

    fig = go.Figure()

    if start_doy > 1:
        before = daily_stats[daily_stats["doy"] < start_doy]
        fig.add_trace(go.Scatter(x=before["datetime_display"], y=before["temp_max"],
                                 fill=None, mode="lines", line_color="rgba(100,100,100,0)",
                                 showlegend=False, hoverinfo="skip"))
        fig.add_trace(go.Scatter(x=before["datetime_display"], y=before["temp_min"],
                                 fill="tonexty", mode="lines", line_color="rgba(100,100,100,0)",
                                 fillcolor="rgba(180,180,180,0.15)",
                                 name="Unselected Period", showlegend=True, hoverinfo="skip"))
        fig.add_trace(go.Scatter(x=before["datetime_display"], y=before["temp_avg"],
                                 mode="lines",
                                 line=dict(color="rgba(150,150,150,0.4)", width=1, dash="dot"),
                                 showlegend=False, hoverinfo="skip"))

    active = daily_stats[(daily_stats["doy"] >= start_doy) & (daily_stats["doy"] <= end_doy)]

    # ASHRAE 90% comfort band
    fig.add_trace(go.Scatter(x=active["datetime_display"], y=active["comfort_90_upper"],
                             fill=None, mode="lines", line_color="rgba(128,128,128,0)",
                             showlegend=False, hoverinfo="skip"))
    fig.add_trace(go.Scatter(x=active["datetime_display"], y=active["comfort_90_lower"],
                             fill="tonexty", mode="lines", line_color="rgba(128,128,128,0)",
                             name="ASHRAE 90% acceptability",
                             fillcolor="rgba(76,175,80,0.4)", hoverinfo="skip"))

    # Temperature range
    fig.add_trace(go.Scatter(x=active["datetime_display"], y=active["temp_max"],
                             fill=None, mode="lines", line_color="rgba(255,0,0,0)",
                             showlegend=False, hoverinfo="skip"))
    fig.add_trace(go.Scatter(x=active["datetime_display"], y=active["temp_min"],
                             fill="tonexty", mode="lines", line_color="rgba(255,0,0,0)",
                             name="Daily Temperature Range",
                             fillcolor="rgba(255,173,173,0.3)",
                             customdata=active["temp_max"],
                             hovertemplate="<b>%{x}</b><br>Min: %{y:.2f}°C<br>Max: %{customdata:.2f}°C<extra></extra>"))
    fig.add_trace(go.Scatter(x=active["datetime_display"], y=active["temp_avg"],
                             mode="lines", name="Average Temperature",
                             line=dict(color="#d32f2f", width=2),
                             hovertemplate="<b>%{x}</b><br>Avg: %{y:.2f}°C<extra></extra>"))

    if end_doy < 365:
        after = daily_stats[daily_stats["doy"] > end_doy]
        if not after.empty:
            fig.add_trace(go.Scatter(x=after["datetime_display"], y=after["temp_max"],
                                     fill=None, mode="lines", line_color="rgba(100,100,100,0)",
                                     showlegend=False, hoverinfo="skip"))
            fig.add_trace(go.Scatter(x=after["datetime_display"], y=after["temp_min"],
                                     fill="tonexty", mode="lines", line_color="rgba(100,100,100,0)",
                                     fillcolor="rgba(180,180,180,0.15)", showlegend=False, hoverinfo="skip"))
            fig.add_trace(go.Scatter(x=after["datetime_display"], y=after["temp_avg"],
                                     mode="lines",
                                     line=dict(color="rgba(150,150,150,0.4)", width=1, dash="dot"),
                                     showlegend=False, hoverinfo="skip"))

    fig.update_layout(
        title="Comfort Analysis – ASHRAE Adaptive Comfort",
        xaxis_title="Day", yaxis_title="Temperature (°C)",
        hovermode="x unified", showlegend=True,
        template="plotly_white", height=450,
    )
    st.plotly_chart(fig, use_container_width=True)


def _render_energy_metrics(df, start_date, end_date, start_hour, end_hour):
    filtered = df[
        (df["datetime"].dt.date >= start_date) &
        (df["datetime"].dt.date <= end_date) &
        (df["hour"].between(start_hour, end_hour))
    ]

    if filtered.empty:
        st.info("No data in the selected date/hour range.")
        return

    hdd18          = (18 - df["dry_bulb_temperature"]).clip(lower=0).sum()
    cdd24          = (df["dry_bulb_temperature"] - 24).clip(lower=0).sum()
    hdd18_filtered = (18 - filtered["dry_bulb_temperature"]).clip(lower=0).sum()
    cdd24_filtered = (filtered["dry_bulb_temperature"] - 24).clip(lower=0).sum()

    monthly_hdd = df.groupby("month").apply(
        lambda x: (18 - x["dry_bulb_temperature"]).clip(lower=0).sum()
    )
    monthly_cdd = df.groupby("month").apply(
        lambda x: (x["dry_bulb_temperature"] - 24).clip(lower=0).sum()
    )
    month_names = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

    st.markdown("#### Energy Performance Indicators")
    c1, c2, c3, c4 = st.columns(4)
    with c1: st.metric("HDD18 (Annual)", f"{hdd18:.0f}",          "Heating Degree-Days")
    with c2: st.metric("CDD24 (Annual)", f"{cdd24:.0f}",          "Cooling Degree-Days")
    with c3: st.metric("HDD18 (Period)", f"{hdd18_filtered:.0f}", "Heating Degree-Days")
    with c4: st.metric("CDD24 (Period)", f"{cdd24_filtered:.0f}", "Cooling Degree-Days")

    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(x=month_names, y=monthly_hdd.values, name="HDD18", marker_color="#2196F3"),
                  secondary_y=False)
    fig.add_trace(go.Bar(x=month_names, y=monthly_cdd.values, name="CDD24", marker_color="#FF9800"),
                  secondary_y=False)
    fig.update_layout(
        title="Monthly Degree-Days Distribution",
        xaxis_title="Month", yaxis_title="Degree-Days",
        hovermode="x unified", height=400, barmode="stack",
    )
    st.plotly_chart(fig, use_container_width=True)
