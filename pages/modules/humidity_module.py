"""Relative Humidity tab rendering module."""

import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st


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
        _render_annual_trend(df, daily_stats)
    elif active_tab == "Monthly Trend":
        _render_monthly_trend(df, start_date, end_date)
    elif active_tab == "Diurnal Profile":
        _render_diurnal_profile(df, start_hour, end_hour)
    elif active_tab == "Comfort Analysis":
        _render_comfort_analysis(df, daily_stats, start_date, end_date)
    elif active_tab == "Energy Metrics":
        _render_energy_metrics(df, start_date, end_date, start_hour, end_hour)


# ─────────────────────────────────────────────────────────────────────────────


def _render_annual_trend(df, daily_stats):
    fig = go.Figure()

    # Comfort band (30–65%)
    fig.add_trace(go.Scatter(x=daily_stats["datetime_display"], y=[65] * len(daily_stats),
                             fill=None, mode="lines", line_color="rgba(128,128,128,0)",
                             showlegend=False, hoverinfo="skip"))
    fig.add_trace(go.Scatter(x=daily_stats["datetime_display"], y=[30] * len(daily_stats),
                             fill="tonexty", mode="lines", line_color="rgba(128,128,128,0)",
                             name="Humidity comfort band",
                             fillcolor="rgba(128,128,128,0.2)", hoverinfo="skip"))

    # RH range
    fig.add_trace(go.Scatter(x=daily_stats["datetime_display"], y=daily_stats["rh_max"],
                             fill=None, mode="lines", line_color="rgba(0,0,255,0)",
                             showlegend=False, hoverinfo="skip"))
    fig.add_trace(go.Scatter(x=daily_stats["datetime_display"], y=daily_stats["rh_min"],
                             fill="tonexty", mode="lines", line_color="rgba(0,0,255,0)",
                             name="Relative humidity Range",
                             fillcolor="rgba(0,150,255,0.3)",
                             hovertemplate="<b>%{x}</b><br>Min: %{y:.1f}%<extra></extra>"))

    # Average
    fig.add_trace(go.Scatter(x=daily_stats["datetime_display"], y=daily_stats["rh_avg"],
                             mode="lines", name="Average Relative humidity",
                             line=dict(color="#00a8ff", width=2),
                             hovertemplate="<b>%{x}</b><br>Avg: %{y:.1f}%<extra></extra>"))

    fig.update_layout(
        title="Annual Profile – Relative Humidity",
        xaxis_title="Day", yaxis_title="Relative Humidity (%)",
        hovermode="x unified", showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        xaxis_rangeslider_visible=False, height=450, template="plotly_white", margin=dict(b=80),
    )
    st.plotly_chart(fig, use_container_width=True)

    # ── KPI cards ─────────────────────────────────────────────────────────────
    rh_min = df["relative_humidity"].min()
    rh_max = df["relative_humidity"].max()
    rh_avg = df["relative_humidity"].mean()

    comfort_hrs   = len(df[(df["relative_humidity"] >= 40) & (df["relative_humidity"] <= 60)])
    comfort_pct   = comfort_hrs / len(df) * 100
    high_rh_hrs   = len(df[df["relative_humidity"] > 60])
    cond_risk_hrs = len(df[df["relative_humidity"] > 75])
    low_rh_hrs    = len(df[df["relative_humidity"] < 30])
    mold_risk_hrs = high_rh_hrs
    over_humid_hrs = len(df[df["relative_humidity"] > 70])

    def _card(label, value, sub, color):
        return f"""
<div style="background:white;padding:16px;border-radius:8px;border-left:4px solid {color};
            box-shadow:0 2px 4px rgba(0,0,0,0.08);text-align:center;">
  <div style="font-size:11px;font-weight:700;color:{color};text-transform:uppercase;letter-spacing:0.5px;">{label}</div>
  <div style="font-size:26px;font-weight:700;color:#2c3e50;margin:8px 0;">{value}</div>
  <div style="font-size:11px;color:#718096;">{sub}</div>
</div>"""

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: st.markdown(_card("Comfort 40-60%",   f"{comfort_pct:.0f} %",    "Occupied RH Hrs",          "#f59e0b"), unsafe_allow_html=True)
    with c2: st.markdown(_card("Peak RH (Occupied)", f"{rh_max:.1f} %",       "All year",                 "#ef4444"), unsafe_allow_html=True)
    with c3: st.markdown(_card("High Humidity Hrs", f"{high_rh_hrs}",         "> 60% RH",                 "#8b5cf6"), unsafe_allow_html=True)
    with c4: st.markdown(_card("Condensation Risk", f"{cond_risk_hrs}",       "Surface Temp < Dew Point", "#06b6d4"), unsafe_allow_html=True)
    with c5: st.markdown(_card("Avg RH",            f"{rh_avg:.1f} %",        "",                         "#3b82f6"), unsafe_allow_html=True)

    c6, c7, c8, c9, c10 = st.columns(5)
    with c6:  st.markdown(_card("Low Humidity Hrs",    f"{low_rh_hrs}",        "< 30% RH",                 "#f59e0b"), unsafe_allow_html=True)
    with c7:  st.markdown(_card("Mold Risk Hrs",       f"{mold_risk_hrs}",     "> 60% RH Sustained",       "#ef4444"), unsafe_allow_html=True)
    with c8:  st.markdown(_card("HVAC RH Control",     f"{comfort_pct:.0f} %", "Outside RH vs Inside RH",  "#06b6d4"), unsafe_allow_html=True)
    with c9:  st.markdown(_card("Overhumidification",  f"{over_humid_hrs}",    "System Failure Indicator", "#3b82f6"), unsafe_allow_html=True)
    with c10: st.markdown(_card("Min RH",              f"{rh_min:.1f} %",      "",                         "#0891b2"), unsafe_allow_html=True)


def _render_monthly_trend(df, start_date, end_date):
    monthly = df.groupby("month").agg(
        rh_min=("relative_humidity", "min"),
        rh_max=("relative_humidity", "max"),
        rh_avg=("relative_humidity", "mean"),
    ).reset_index()

    month_lbl = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    monthly["month_name"] = monthly["month"].apply(lambda x: month_lbl[x - 1])

    start_month = start_date.month
    end_month   = end_date.month

    fig = go.Figure()

    def _grey_band(data):
        fig.add_trace(go.Scatter(x=data["month_name"], y=[65] * len(data),
                                 fill=None, mode="lines", line_color="rgba(100,100,100,0)",
                                 showlegend=False, hoverinfo="skip"))
        fig.add_trace(go.Scatter(x=data["month_name"], y=[30] * len(data),
                                 fill="tonexty", mode="lines", line_color="rgba(100,100,100,0)",
                                 fillcolor="rgba(180,180,180,0.15)", showlegend=False, hoverinfo="skip"))
        fig.add_trace(go.Scatter(x=data["month_name"], y=data["rh_max"],
                                 fill=None, mode="lines", line_color="rgba(100,100,100,0)",
                                 showlegend=False, hoverinfo="skip"))
        fig.add_trace(go.Scatter(x=data["month_name"], y=data["rh_min"],
                                 fill="tonexty", mode="lines", line_color="rgba(100,100,100,0)",
                                 fillcolor="rgba(180,180,180,0.15)", showlegend=False, hoverinfo="skip"))
        fig.add_trace(go.Scatter(x=data["month_name"], y=data["rh_avg"],
                                 mode="lines", line=dict(color="rgba(150,150,150,0.4)", width=1, dash="dot"),
                                 showlegend=False, hoverinfo="skip"))

    before = monthly[monthly["month"] < start_month]
    after  = monthly[monthly["month"] > end_month]
    active = monthly[(monthly["month"] >= start_month) & (monthly["month"] <= end_month)]

    if not before.empty:
        fig.add_trace(go.Scatter(x=before["month_name"], y=[65] * len(before),
                                 fill=None, mode="lines", line_color="rgba(100,100,100,0)",
                                 showlegend=False, hoverinfo="skip"))
        fig.add_trace(go.Scatter(x=before["month_name"], y=[30] * len(before),
                                 fill="tonexty", mode="lines", line_color="rgba(100,100,100,0)",
                                 fillcolor="rgba(180,180,180,0.15)", name="Unselected Period",
                                 showlegend=True, hoverinfo="skip"))
        _grey_band(before)

    # Active: comfort band
    fig.add_trace(go.Scatter(x=active["month_name"], y=[65] * len(active),
                             fill=None, mode="lines", line_color="rgba(128,128,128,0)",
                             showlegend=False, hoverinfo="skip"))
    fig.add_trace(go.Scatter(x=active["month_name"], y=[30] * len(active),
                             fill="tonexty", mode="lines", line_color="rgba(128,128,128,0)",
                             name="Humidity comfort band (30-65%)",
                             fillcolor="rgba(128,128,128,0.2)", hoverinfo="skip"))

    fig.add_trace(go.Scatter(x=active["month_name"], y=active["rh_max"],
                             fill=None, mode="lines", line_color="rgba(0,0,255,0)",
                             showlegend=False, hoverinfo="skip"))
    fig.add_trace(go.Scatter(x=active["month_name"], y=active["rh_min"],
                             fill="tonexty", mode="lines", line_color="rgba(0,0,255,0)",
                             name="Monthly Humidity Range", fillcolor="rgba(0,150,255,0.3)",
                             customdata=active["rh_max"],
                             hovertemplate="<b>%{x}</b><br>Min: %{y:.1f}%<br>Max: %{customdata:.1f}%<extra></extra>"))
    fig.add_trace(go.Scatter(x=active["month_name"], y=active["rh_avg"],
                             mode="lines+markers", name="Monthly Average Humidity",
                             line=dict(color="#00a8ff", width=2), marker=dict(size=8),
                             hovertemplate="<b>%{x}</b><br>Avg: %{y:.1f}%<extra></extra>"))

    if not after.empty:
        _grey_band(after)

    fig.update_layout(
        title="Monthly Relative Humidity Trend",
        xaxis_title="Month", yaxis_title="Relative Humidity (%)",
        hovermode="x unified", showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        height=450, template="plotly_white", margin=dict(b=80),
    )
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### Monthly Humidity Summary")
    kpi = monthly[["month_name", "rh_min", "rh_max", "rh_avg"]].copy()
    kpi.columns = ["Month", "Min (%)", "Max (%)", "Avg (%)"]
    st.dataframe(kpi, use_container_width=True, hide_index=True,
                 column_config={
                     "Min (%)": st.column_config.NumberColumn(format="%.1f"),
                     "Max (%)": st.column_config.NumberColumn(format="%.1f"),
                     "Avg (%)": st.column_config.NumberColumn(format="%.1f"),
                 })


def _render_diurnal_profile(df, start_hour, end_hour):
    hourly = df.groupby(["month", "hour"]).agg(
        rh_min=("relative_humidity", "min"),
        rh_max=("relative_humidity", "max"),
        rh_avg=("relative_humidity", "mean"),
    ).reset_index()

    avg = hourly.groupby("hour").agg(
        rh_min=("rh_min", "min"),
        rh_max=("rh_max", "max"),
        rh_avg=("rh_avg", "mean"),
    ).reset_index()

    fig = go.Figure()

    def _grey_rh(data):
        fig.add_trace(go.Scatter(x=data["hour"], y=[65] * len(data),
                                 fill=None, mode="lines", line_color="rgba(100,100,100,0)",
                                 showlegend=False, hoverinfo="skip"))
        fig.add_trace(go.Scatter(x=data["hour"], y=[30] * len(data),
                                 fill="tonexty", mode="lines", line_color="rgba(100,100,100,0)",
                                 fillcolor="rgba(180,180,180,0.15)", showlegend=False, hoverinfo="skip"))
        fig.add_trace(go.Scatter(x=data["hour"], y=data["rh_max"],
                                 fill=None, mode="lines", line_color="rgba(100,100,100,0)",
                                 showlegend=False, hoverinfo="skip"))
        fig.add_trace(go.Scatter(x=data["hour"], y=data["rh_min"],
                                 fill="tonexty", mode="lines", line_color="rgba(100,100,100,0)",
                                 fillcolor="rgba(180,180,180,0.15)", showlegend=False, hoverinfo="skip"))
        fig.add_trace(go.Scatter(x=data["hour"], y=data["rh_avg"],
                                 mode="lines+markers",
                                 line=dict(color="rgba(150,150,150,0.4)", width=1, dash="dot"),
                                 marker=dict(size=4), showlegend=False, hoverinfo="skip"))

    if start_hour > 0:
        before = avg[avg["hour"] < start_hour]
        fig.add_trace(go.Scatter(x=before["hour"], y=[65] * len(before),
                                 fill=None, mode="lines", line_color="rgba(100,100,100,0)",
                                 showlegend=False, hoverinfo="skip"))
        fig.add_trace(go.Scatter(x=before["hour"], y=[30] * len(before),
                                 fill="tonexty", mode="lines", line_color="rgba(100,100,100,0)",
                                 fillcolor="rgba(180,180,180,0.15)",
                                 name="Unselected Hours", showlegend=True, hoverinfo="skip"))
        _grey_rh(before)

    active = avg[(avg["hour"] >= start_hour) & (avg["hour"] <= end_hour)]

    fig.add_trace(go.Scatter(x=active["hour"], y=[65] * len(active),
                             fill=None, mode="lines", line_color="rgba(128,128,128,0)",
                             showlegend=False, hoverinfo="skip"))
    fig.add_trace(go.Scatter(x=active["hour"], y=[30] * len(active),
                             fill="tonexty", mode="lines", line_color="rgba(128,128,128,0)",
                             name="Comfort band (30-65%)",
                             fillcolor="rgba(128,128,128,0.2)", hoverinfo="skip"))
    fig.add_trace(go.Scatter(x=active["hour"], y=active["rh_max"],
                             fill=None, mode="lines", line_color="rgba(0,0,255,0)",
                             showlegend=False, hoverinfo="skip"))
    fig.add_trace(go.Scatter(x=active["hour"], y=active["rh_min"],
                             fill="tonexty", mode="lines", line_color="rgba(0,0,255,0)",
                             name="Humidity Range", fillcolor="rgba(0,150,255,0.3)",
                             customdata=active["rh_max"],
                             hovertemplate="<b>Hour %{x}:00</b><br>Min: %{y:.1f}%<br>Max: %{customdata:.1f}%<extra></extra>"))
    fig.add_trace(go.Scatter(x=active["hour"], y=active["rh_avg"],
                             mode="lines+markers", name="Average Humidity",
                             line=dict(color="#00a8ff", width=2), marker=dict(size=6),
                             hovertemplate="<b>Hour %{x}:00</b><br>Avg: %{y:.1f}%<extra></extra>"))

    if end_hour < 23:
        after = avg[avg["hour"] > end_hour]
        _grey_rh(after)

    fig.update_layout(
        title="Diurnal Humidity Profile",
        xaxis_title="Hour of Day", yaxis_title="Relative Humidity (%)",
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
        fig.add_trace(go.Scatter(x=before["datetime_display"], y=before["rh_max"],
                                 fill=None, mode="lines", line_color="rgba(100,100,100,0)",
                                 showlegend=False, hoverinfo="skip"))
        fig.add_trace(go.Scatter(x=before["datetime_display"], y=before["rh_min"],
                                 fill="tonexty", mode="lines", line_color="rgba(100,100,100,0)",
                                 fillcolor="rgba(180,180,180,0.15)",
                                 name="Unselected Period", showlegend=True, hoverinfo="skip"))
        fig.add_trace(go.Scatter(x=before["datetime_display"], y=before["rh_avg"],
                                 mode="lines",
                                 line=dict(color="rgba(150,150,150,0.4)", width=1, dash="dot"),
                                 showlegend=False, hoverinfo="skip"))

    active = daily_stats[(daily_stats["doy"] >= start_doy) & (daily_stats["doy"] <= end_doy)]

    # Comfort band 40–60%
    fig.add_trace(go.Scatter(x=active["datetime_display"], y=[60] * len(active),
                             fill=None, mode="lines", line_color="rgba(128,128,128,0)",
                             showlegend=False, hoverinfo="skip"))
    fig.add_trace(go.Scatter(x=active["datetime_display"], y=[40] * len(active),
                             fill="tonexty", mode="lines", line_color="rgba(128,128,128,0)",
                             name="Comfort Band (40-60%)",
                             fillcolor="rgba(76,175,80,0.4)", hoverinfo="skip"))

    fig.add_trace(go.Scatter(x=active["datetime_display"], y=active["rh_max"],
                             fill=None, mode="lines", line_color="rgba(0,150,255,0)",
                             showlegend=False, hoverinfo="skip"))
    fig.add_trace(go.Scatter(x=active["datetime_display"], y=active["rh_min"],
                             fill="tonexty", mode="lines", line_color="rgba(0,150,255,0)",
                             name="Daily RH Range", fillcolor="rgba(0,150,255,0.3)",
                             customdata=active["rh_max"],
                             hovertemplate="<b>%{x}</b><br>Min: %{y:.2f}%<br>Max: %{customdata:.2f}%<extra></extra>"))
    fig.add_trace(go.Scatter(x=active["datetime_display"], y=active["rh_avg"],
                             mode="lines", name="Average RH",
                             line=dict(color="#00a8ff", width=2),
                             hovertemplate="<b>%{x}</b><br>Avg: %{y:.2f}%<extra></extra>"))

    if end_doy < 365:
        after = daily_stats[daily_stats["doy"] > end_doy]
        if not after.empty:
            fig.add_trace(go.Scatter(x=after["datetime_display"], y=after["rh_max"],
                                     fill=None, mode="lines", line_color="rgba(100,100,100,0)",
                                     showlegend=False, hoverinfo="skip"))
            fig.add_trace(go.Scatter(x=after["datetime_display"], y=after["rh_min"],
                                     fill="tonexty", mode="lines", line_color="rgba(100,100,100,0)",
                                     fillcolor="rgba(180,180,180,0.15)", showlegend=False, hoverinfo="skip"))
            fig.add_trace(go.Scatter(x=after["datetime_display"], y=after["rh_avg"],
                                     mode="lines",
                                     line=dict(color="rgba(150,150,150,0.4)", width=1, dash="dot"),
                                     showlegend=False, hoverinfo="skip"))

    fig.update_layout(
        title="Humidity Comfort Analysis – Optimal Range (40-60%)",
        xaxis_title="Day", yaxis_title="Relative Humidity (%)",
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

    high_rh_annual    = len(df[df["relative_humidity"] > 60])
    cond_risk_annual  = len(df[df["relative_humidity"] > 75])
    high_rh_filtered  = len(filtered[filtered["relative_humidity"] > 60])
    cond_risk_filtered = len(filtered[filtered["relative_humidity"] > 75])

    monthly_high  = df.groupby("month").apply(lambda x: len(x[x["relative_humidity"] > 60]))
    monthly_cond  = df.groupby("month").apply(lambda x: len(x[x["relative_humidity"] > 75]))
    monthly_low   = df.groupby("month").apply(lambda x: len(x[x["relative_humidity"] < 30]))

    month_names = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

    st.markdown("#### Humidity Performance Indicators")
    c1, c2, c3, c4 = st.columns(4)
    with c1: st.metric("High RH Hrs (Annual)",        f"{high_rh_annual}",    ">60% RH")
    with c2: st.metric("Condensation Risk (Annual)",  f"{cond_risk_annual}",  ">75% RH")
    with c3: st.metric("High RH Hrs (Period)",        f"{high_rh_filtered}",  ">60% RH")
    with c4: st.metric("Condensation (Period)",       f"{cond_risk_filtered}", ">75% RH")

    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(x=month_names, y=monthly_high.values,
                         name="High RH (>60%)", marker_color="#0099ff"), secondary_y=False)
    fig.add_trace(go.Bar(x=month_names, y=monthly_cond.values,
                         name="Condensation Risk (>75%)", marker_color="#FF6B6B"), secondary_y=False)
    fig.add_trace(go.Scatter(x=month_names, y=monthly_low.values,
                             name="Low RH (<30%)", line=dict(color="#FFA500", width=2),
                             mode="lines+markers"), secondary_y=False)
    fig.update_layout(
        title="Monthly Humidity Risk Distribution",
        xaxis_title="Month", yaxis_title="Hours",
        hovermode="x unified", height=400, barmode="group",
    )
    st.plotly_chart(fig, use_container_width=True)
