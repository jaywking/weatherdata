#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Forecast Data Processing
"""

import pandas as pd
from datetime import datetime, timedelta
import pytz

from nws_api import get_points, latest_observation, hourly_forecast, grid_supplement

ET = pytz.timezone("America/New_York")

def feels_like_f(temp_f, wind_mph, rh_pct):
    """
    Apparent temperature:
    - If Wind Chill conditions (<= 50°F and wind >= 3 mph): apply wind chill
    - Else if Heat Index conditions (>= 80°F and RH >= 40%): apply heat index (simplified Rothfusz)
    - Otherwise return temp_f
    """
    t = temp_f
    v = max(0.0, wind_mph)
    rh = max(0.0, min(100.0, rh_pct if rh_pct is not None else 50.0))

    # Wind Chill (NOAA)
    if t <= 50 and v >= 3:
        wc = 35.74 + 0.6215*t - 35.75*(v**0.16) + 0.4275*t*(v**0.16)
        return wc

    # Heat Index (simplified)
    if t >= 80 and rh >= 40:
        hi = (-42.379 + 2.04901523*t + 10.14333127*rh
              - 0.22475541*t*rh - 6.83783e-3*(t**2)
              - 5.481717e-2*(rh**2) + 1.22874e-3*(t**2)*rh
              + 8.5282e-4*t*(rh**2) - 1.99e-6*(t**2)*(rh**2))
        return hi

    return t

def build_site_package(name, lat, lon, now_et):
    print("  Fetching NWS data points...")
    pts = get_points(lat, lon)
    print("  Fetching latest observation...")
    obs = latest_observation(pts)
    print("  Fetching hourly forecast...")
    df_hourly = hourly_forecast(pts)               # temp_f, wind_mph, gust_mph, wind_dir_text, rh, pop, cloud(None)
    print("  Fetching grid supplement data...")
    df_grid   = grid_supplement(pts)               # skyCover, apparentTemperature, windGust (maybe)
    print("  Processing and combining data...")
    # Combine
    df = df_hourly.join(df_grid, how="left")

    # Compute feels-like if not provided
    feels_f_list = []
    for idx, row in df.iterrows():
        temp_f = row.get("temp_f")
        wind_mph = row.get("wind_mph") or 0
        rh = row.get("rh")
        app_f = row.get("apparentTemperature")
        if app_f is not None and not pd.isna(app_f):
            feels_f = app_f
        else:
            feels_f = feels_like_f(temp_f, wind_mph, rh) if temp_f is not None else None
        feels_f_list.append(feels_f)
    df["feels_f"] = feels_f_list

    # Cloud cover from grid
    if "skyCover" in df.columns:
        df["cloud_pct"] = df["skyCover"]
    else:
        df["cloud_pct"] = None

    # If gust missing, use grid gust if present
    if "windGust" in df.columns:
        gust_missing = df["gust_mph"].isna()
        df.loc[gust_missing, "gust_mph"] = df.loc[gust_missing, "windGust"]

    # Conversions + rounding to whole numbers
    df["temp_c"]  = ((df["temp_f"] - 32) * 5/9)
    df["feels_c"] = ((df["feels_f"] - 32) * 5/9)
    for col in ["temp_f","feels_f","temp_c","feels_c","wind_mph","gust_mph","rh","pop","cloud_pct"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').round().astype("Int64")

    # Prepare nice columns for export
    out = pd.DataFrame(index=df.index)
    out["Date/Time (EDT)"] = [ts.strftime("%b %d %H:%M") for ts in df.index]
    out["Temp (°C)"]       = df["temp_c"]
    out["Feels (°C)"]      = df["feels_c"]
    out["Wind (km/h)"]     = (df["wind_mph"] * 1.609).round().astype("Int64")
    out["Gusts (km/h)"]    = (df["gust_mph"] * 1.609).round().astype("Int64")
    out["Dir"]             = df.get("wind_dir_text", pd.Series(index=df.index, dtype="object")).fillna("")
    out["RH (%)"]          = df["rh"]
    out["PoP (%)"]         = df["pop"]
    out["Cloud (%)"]       = df["cloud_pct"]

    # Imperial view
    out_imp = pd.DataFrame(index=df.index)
    out_imp["Date/Time (EDT)"] = out["Date/Time (EDT)"]
    out_imp["Temp (°F)"]       = df["temp_f"]
    out_imp["Feels (°F)"]      = df["feels_f"]
    out_imp["Wind (mph)"]      = df["wind_mph"]
    out_imp["Gusts (mph)"]     = df["gust_mph"]
    out_imp["Dir"]             = out["Dir"]
    out_imp["RH (%)"]          = out["RH (%)"] if "RH (%)" in out.columns else df["rh"]
    out_imp["PoP (%)"]         = out["PoP (%)"] if "PoP (%)" in out.columns else df["pop"]
    out_imp["Cloud (%)"]       = out["Cloud (%)"]

    # Make interval tables
    now_idx = df.index[df.index >= now_et.replace(minute=0, second=0, microsecond=0)]
    if len(now_idx) == 0:
        start = df.index.min()
    else:
        start = now_idx.min()

    # 36h: 0–24h every 2h, 24–36h every 4h
    end_24 = start + timedelta(hours=24)
    end_36 = start + timedelta(hours=36)

    short_0_24 = out.loc[(out.index >= start) & (out.index <= end_24)]
    short_0_24 = short_0_24.iloc[::2]  # every 2 hours

    short_24_36 = out.loc[(out.index > end_24) & (out.index <= end_36)]
    short_24_36 = short_24_36.iloc[::4]  # every 4 hours

    short_metric = pd.concat([short_0_24, short_24_36]).copy()
    short_imper  = out_imp.loc[short_metric.index].copy()

    # 7-day: every 6 hours
    end_7d = start + timedelta(days=7)
    long_metric = out.loc[(out.index >= start) & (out.index <= end_7d)].iloc[::6].copy()
    long_imper  = out_imp.loc[long_metric.index].copy()

    # Current conditions block (from obs)
    curr = None
    if obs and obs.get("observed_at"):
        curr = {
            "observed_at": obs["observed_at"].strftime("%b %d, %Y – %I:%M %p EDT"),
            "temp_c": None if obs["temp_c"] is None else round(obs["temp_c"]),
            "temp_f": None if obs["temp_f"] is None else round(obs["temp_f"]),
            "rh": None if obs["rh"] is None else round(obs["rh"]),
            "wind_mph": None if obs["wind_mph"] is None else round(obs["wind_mph"]),
            "gust_mph": None if obs["gust_mph"] is None else round(obs["gust_mph"]),
            "wind_dir_deg": obs["wind_dir_deg"],
            "conditions": obs["conditions"] or "—",
            "station_id": obs["station_id"],
        }

    # Package everything for downstream consumers
    short_metric_out = short_metric.reset_index(drop=True)
    short_imper_out = short_imper.reset_index(drop=True)
    long_metric_out = long_metric.reset_index(drop=True)
    long_imper_out = long_imper.reset_index(drop=True)

    # Human-friendly source + alert note placeholder
    source_url = f"https://forecast.weather.gov/MapClick.php?lat={lat}&lon={lon}"
    alert_note = "No county-wide alerts shown on NWS point page at generation time."

    return {
        "name": name,
        "current": curr,
        "short_metric": short_metric_out,
        "short_imper": short_imper_out,
        "long_metric": long_metric_out,
        "long_imper": long_imper_out,
        "alert_note": alert_note,
        "source_url": source_url,
    }

def narrative_from_tables(short_imp_df):
    if short_imp_df.empty:
        return [
            "Limited short-term data available from NWS at this time."
        ]
    hi = int(short_imp_df["Temp (°F)"].max())
    lo = int(short_imp_df["Temp (°F)"].min())
    avg_wind = int(short_imp_df["Wind (mph)"].fillna(0).mean())
    gusts = int(short_imp_df["Gusts (mph)"].fillna(0).max())
    popmax = int(short_imp_df["PoP (%)"].fillna(0).max())
    cloud_mid = int(short_imp_df["Cloud (%)"].dropna().median()) if short_imp_df["Cloud (%)"].notna().any() else 0

    sky = "mostly clear" if cloud_mid <= 25 else ("partly cloudy" if cloud_mid <= 60 else "mostly cloudy")
    precip = "dry" if popmax <= 10 else ("a slight chance of showers" if popmax <= 30 else "showers possible")

    avg_kmh = int(round(avg_wind * 1.60934))
    gust_kmh = int(round(gusts * 1.60934)) if gusts > 0 else 0

    wind_narrative = f"Winds generally around {avg_wind} mph ({avg_kmh} km/h)"
    if gusts > 0:
        wind_narrative += f" with gusts up to ~{gusts} mph ({gust_kmh} km/h)."
    else:
        wind_narrative += "."

    return [
        f"Next 36 hours: {sky} and {precip}.",
        f"Temperatures range from about {lo}–{hi} °F ({round((lo-32)*5/9)}–{round((hi-32)*5/9)} °C).",
        wind_narrative,
    ]
