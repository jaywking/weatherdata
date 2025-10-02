#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Lake Placid Forecast Generator
- Sources: NWS / NOAA (Gridpoint + Stations API)
- Locations:
  - DHI – Mt. Whiteface Base (44.35, -73.86)
  - XC  – Mt. Hoevenburg Base (44.2192, -73.9209)

Outputs (saved in the script folder):
  - Lake_Placid_Forecast_Full_<YYYYMMDD-HHMM_ET>.md
  - Lake_Placid_Forecast_Full_<YYYYMMDD-HHMM_ET>.xlsx
  - Lake_Placid_Forecast_With_Tables_<YYYYMMDD-HHMM_ET>.pdf
  - Lake_Placid_Forecast_Summary_<YYYYMMDD-HHMM_ET>.pdf

Requires:
  pip install requests pandas openpyxl reportlab pytz
"""

import re
import os
import pytz
import requests
import pandas as pd
from datetime import datetime, timedelta
from dateutil import parser as dtparser
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors

# ----------------------------
# CONFIG
# ----------------------------
ET = pytz.timezone("America/New_York")
USER_AGENT = "AboveTheLineSafety Forecast Tool (contact: ops@example.com)"  # change if desired

SITES = {
    "DHI – Mt. Whiteface Base": {"lat": 44.35, "lon": -73.86},
    "XC – Mt. Hoevenburg Base": {"lat": 44.2192, "lon": -73.9209},
}

# Output file base name will include ET timestamp
now_et = datetime.now(ET)
STAMP_HUMAN = now_et.strftime("Generated: %b %d, %Y – %I:%M %p EDT")
STAMP_TAG   = now_et.strftime("%Y%m%d-%H%M")  # for filenames

# ----------------------------
# HELPERS
# ----------------------------
def ua_headers():
    return {
        "User-Agent": USER_AGENT,
        "Accept": "application/ld+json, application/json"
    }

def iso_to_et(s):
    if not s:
        return None
    dt = dtparser.isoparse(s)
    if not dt.tzinfo:
        dt = pytz.UTC.localize(dt)
    return dt.astimezone(ET)

def mph_from_wind_string(s):
    """
    NWS hourly windSpeed is often like '5 mph', '5 to 10 mph', or 'Calm'.
    Return integer mph (average if a range).
    """
    if not s or s.strip().lower() == "calm":
        return 0
    nums = list(map(int, re.findall(r"\d+", s)))
    if not nums:
        return 0
    return int(round(sum(nums) / len(nums)))

def feels_like_f(temp_f, wind_mph, rh_pct):
    """
    Apparent temperature:
    - If Wind Chill conditions (<= 50\u00B0F and wind >= 3 mph): apply wind chill
    - Else if Heat Index conditions (>= 80\u00B0F and RH >= 40%): apply heat index (simplified Rothfusz)
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

def round_int(x):
    return None if x is None else int(round(x))

def safe_int(val):
    try:
        return int(round(float(val)))
    except:
        return None

def pick_every(df, minutes):
    """Pick rows at a fixed stride in minutes starting from the first row."""
    if df.empty:
        return df
    base = df.index.min()
    # Select rows whose minute difference is divisible by step
    sel = df.loc[(df.index - base).total_seconds() % (minutes*60) == 0]
    return sel

# ----------------------------
# NWS API
# ----------------------------
def get_points(lat, lon):
    r = requests.get(f"https://api.weather.gov/points/{lat},{lon}", headers=ua_headers(), timeout=30)
    r.raise_for_status()
    data = r.json()
    props = data.get("properties") if isinstance(data, dict) else None
    return props if isinstance(props, dict) else data

def get_json(url):
    r = requests.get(url, headers=ua_headers(), timeout=60)
    r.raise_for_status()
    data = r.json()
    props = data.get("properties") if isinstance(data, dict) else None
    if isinstance(props, dict):
        return props
    return data

def latest_observation(points_json):
    """
    Fetch latest observation using the observationStations link returned by /points.
    Returns dict with keys used in Current Conditions.
    """
    stations_url = points_json["observationStations"]
    st = get_json(stations_url)
    station_ids = [s["stationIdentifier"] for s in st.get("@graph", [])]
    # Try first station; if it fails, iterate
    obs = None
    for sid in station_ids[:5]:
        try:
            o = get_json(f"https://api.weather.gov/stations/{sid}/observations/latest")
            if o.get("timestamp"):
                obs = o
                break
        except Exception:
            continue
    if not obs:
        return None

    p = obs
    when = iso_to_et(p.get("timestamp"))
    temp_c = p.get("temperature", {}).get("value")  # in \u00B0C
    rh = p.get("relativeHumidity", {}).get("value")
    wind_mps = p.get("windSpeed", {}).get("value")   # m/s
    gust_mps = p.get("windGust", {}).get("value")    # m/s
    wdir = p.get("windDirection", {}).get("value")   # degrees
    vis_m = p.get("visibility", {}).get("value")     # meters
    text = p.get("textDescription")
    pressure_pa = p.get("barometricPressure", {}).get("value")

    # Conversions
    temp_f = None if temp_c is None else (temp_c * 9/5 + 32)
    wind_mph = None if wind_mps is None else wind_mps * 2.23694
    gust_mph = None if gust_mps is None else gust_mps * 2.23694
    vis_mi = None if vis_m is None else vis_m / 1609.34

    return {
        "observed_at": when,
        "conditions": text,
        "temp_c": temp_c,
        "temp_f": temp_f,
        "rh": rh,
        "wind_mph": wind_mph,
        "gust_mph": gust_mph,
        "wind_dir_deg": wdir,
        "visibility_mi": vis_mi,
        "pressure_inHg": None if pressure_pa is None else pressure_pa * 0.0002953,
        "station_id": station_ids[0] if station_ids else None
    }

def hourly_forecast(points_json):
    """
    Use the hourly forecast periods from /forecast/hourly (list of 48+ periods)
    Returns pandas DataFrame indexed by ET time.
    """
    hourly_url = points_json["forecastHourly"]
    h = get_json(hourly_url)
    periods = h["periods"]
    rows = []
    for p in periods:
        t = iso_to_et(p["startTime"])
        temp_f = p.get("temperature")
        wind_dir = p.get("windDirection")   # e.g., 'NW'
        wind_spd = mph_from_wind_string(p.get("windSpeed"))
        gust_spd = mph_from_wind_string(p.get("windGust")) if p.get("windGust") else None
        rh = p.get("relativeHumidity", {}).get("value")
        pop = p.get("probabilityOfPrecipitation", {}).get("value")
        # cloud % isn't in hourly; we will try grid data next; placeholder None here
        cloud = None
        rows.append({
            "time": t, "temp_f": temp_f, "wind_mph": wind_spd, "gust_mph": gust_spd,
            "wind_dir_text": wind_dir, "rh": rh, "pop": pop, "cloud": cloud
        })
    df = pd.DataFrame(rows).dropna(subset=["time"]).set_index("time").sort_index()
    return df

def grid_supplement(points_json):
    """
    Pull grid data to supplement cloud cover and (if present) apparentTemperature + windGust.
    """
    grid_url = points_json["forecastGridData"]
    props = get_json(grid_url)

    def series_from_grid(field):
        # Each field is a "values" list with "validTime" ISO duration (e.g., "2025-09-29T18:00:00+00:00/PT1H")
        out = []
        # Handle uom; default for wind is km/h, but could be knots
        uom = props.get(field, {}).get("uom", "")

        for entry in props.get(field, {}).get("values", []):
            valid = entry.get("validTime")  # "start/period"
            value = entry.get("value")
            if not valid or value is None:
                continue

            start_str, dur = valid.split("/")
            start = iso_to_et(start_str)
            # Duration parsing (PT1H, PT3H, etc.)
            m = re.match(r"PT(\d+)H", dur)
            hours = int(m.group(1)) if m else 1

            # Convert value if needed
            if field == "windGust":
                if uom in {"km_h-1", "km/h"}:
                    value = value * 0.621371  # km/h to mph
                elif uom == "kt":
                    value = value * 1.15078  # knots to mph

            for h in range(hours):
                out.append({"time": start + timedelta(hours=h), field: value})
        if not out:
            return pd.DataFrame(columns=["time", field]).set_index("time")
        return pd.DataFrame(out).set_index("time")

    # Pull fields we care about
    sky = series_from_grid("skyCover")           # %
    app = series_from_grid("apparentTemperature")# \u00B0F
    gust = series_from_grid("windGust")

    frames = [df for df in (sky, app, gust) if not df.empty]
    if not frames:
        return pd.DataFrame()

    out = pd.concat(frames, axis=1, join="outer")
    out = out.sort_index()
    return out
# ----------------------------
# BUILD TABLES
# ----------------------------
def build_site_package(name, lat, lon, now_et):
    pts = get_points(lat, lon)
    obs = latest_observation(pts)
    df_hourly = hourly_forecast(pts)               # temp_f, wind_mph, gust_mph, wind_dir_text, rh, pop, cloud(None)
    df_grid   = grid_supplement(pts)               # skyCover, apparentTemperature, windGust (maybe)
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
        df["gust_mph"] = df["gust_mph"].fillna(df["windGust"])

    # Conversions + rounding to whole numbers
    df["temp_c"]  = ((df["temp_f"] - 32) * 5/9)
    df["feels_c"] = ((df["feels_f"] - 32) * 5/9)
    for col in ["temp_f","feels_f","temp_c","feels_c","wind_mph","gust_mph","rh","pop","cloud_pct"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').round().astype("Int64")

    # Prepare nice columns for export
    out = pd.DataFrame(index=df.index)
    out["Date/Time (EDT)"] = [ts.strftime("%b %d %H:%M") for ts in df.index]
    out["Temp (\u00B0C)"]       = df["temp_c"]
    out["Feels (\u00B0C)"]      = df["feels_c"]
    out["Wind (km/h)"]     = (df["wind_mph"] * 1.609).round().astype("Int64")
    out["Gusts (km/h)"]    = (df["gust_mph"] * 1.609).round().astype("Int64")
    out["Dir"]             = df.get("wind_dir_text", pd.Series(index=df.index, dtype="object")).fillna("")
    out["RH (%)"]          = df["rh"]
    out["PoP (%)"]         = df["pop"]
    out["Cloud (%)"]       = df["cloud_pct"]

    # Imperial view
    out_imp = pd.DataFrame(index=df.index)
    out_imp["Date/Time (EDT)"] = out["Date/Time (EDT)"]
    out_imp["Temp (\u00B0F)"]       = df["temp_f"]
    out_imp["Feels (\u00B0F)"]      = df["feels_f"]
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

# ----------------------------
# GENERATE NARRATIVE (brief)
# ----------------------------
def narrative_from_tables(short_imp_df):
    if short_imp_df.empty:
        return [
            "Limited short-term data available from NWS at this time."
        ]
    hi = int(short_imp_df["Temp (\u00B0F)"].max())
    lo = int(short_imp_df["Temp (\u00B0F)"].min())
    avg_wind = int(short_imp_df["Wind (mph)"].fillna(0).mean())
    gusts = int(short_imp_df["Gusts (mph)"].fillna(0).max())
    popmax = int(short_imp_df["PoP (%)"].fillna(0).max())
    cloud_mid = int(short_imp_df["Cloud (%)"].dropna().median()) if short_imp_df["Cloud (%)"].notna().any() else 0

    sky = "mostly clear" if cloud_mid <= 25 else ("partly cloudy" if cloud_mid <= 60 else "mostly cloudy")
    precip = "dry" if popmax <= 10 else ("a slight chance of showers" if popmax <= 30 else "showers possible")

    wind_narrative = f"Winds generally around {avg_wind} mph"
    if gusts > 0:
        wind_narrative += f" with gusts up to ~{gusts} mph."
    else:
        wind_narrative += "."

    return [
        f"Next 36 hours: {sky} and {precip}.",
        f"Temperatures range from about {lo}–{hi} \u00B0F ({round((lo-32)*5/9)}–{round((hi-32)*5/9)} \u00B0C).",
        wind_narrative,
    ]

# ----------------------------
# EXPORTS
# ----------------------------
def save_whatsapp_summary(pkg_dhi, pkg_xc, txt_path, stamp_human=None):
    def section(pkg):
        name = pkg["name"]
        curr = pkg["current"]
        wa = []
        wa.append(f"*{name}*")
        wa.append(f"*Current Conditions*")
        if curr:
            temp_str = f"{curr['temp_c']} \u00B0C / {curr['temp_f']} \u00B0F" if curr['temp_c'] is not None else "Not available"
            rh_str = f"{curr['rh']}%" if curr['rh'] is not None else "Not available"
            wind = f"{curr['wind_mph']} mph" if curr['wind_mph'] is not None else "-"
            gust = f", gusting {curr['gust_mph']} mph" if curr['gust_mph'] else ""
            wdir = f" from {curr['wind_dir_deg']}\u00B0" if curr['wind_dir_deg'] is not None else ""
            wa.extend([
                f"- Conditions: {curr['conditions']}",
                f"- Temp: {temp_str}",
                f"- Wind: {wind}{gust}{wdir}",
                f"- RH: {rh_str}",
            ])
        else:
            wa.append("_No recent observation available._")

        wa.append("\n*Short-term Forecast (36h)*")
        for line in narrative_from_tables(pkg["short_imper"]):
            wa.append(f"- {line}")
        return "\n".join(wa)

    stamp_to_use = stamp_human or STAMP_HUMAN
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("*Weather Forecast - Lake Placid Area*\n")
        f.write(f"_{stamp_to_use}_\n\n")
        f.write(section(pkg_dhi))
        f.write("\n\n---\n\n")
        f.write(section(pkg_xc))
def save_markdown(pkg_dhi, pkg_xc, md_path):
    def section(pkg):
        name = pkg["name"]
        curr = pkg["current"]
        md = []
        md.append(f"## {name}\n")
        md.append("### Current Conditions")
        if curr:
            md.append(f"**Observed at {curr['observed_at']}**")
            wind = f"{curr['wind_mph']} mph" if curr['wind_mph'] is not None else "—"
            gust = f", gusting {curr['gust_mph']} mph" if curr['gust_mph'] else ""
            wdir = f" from {curr['wind_dir_deg']}\u00B0" if curr['wind_dir_deg'] is not None else ""
            md.extend([
                f"- Conditions: {curr['conditions']}",
                f"- Temp: {curr['temp_c']} \u00B0C / {curr['temp_f']} \u00B0F",
                f"- Wind: {wind}{gust}{wdir}",
                f"- RH: {curr['rh']}%",
                f"- Alerts: {pkg['alert_note']}",
                f"- Source: [NWS Gridpoint & Stations]({pkg['source_url']})",
            ])
        else:
            md.append("_No recent observation available from nearest station._")

        md.append("\n### Short-term Forecast (36 Hours, High Detail)\n")
        for line in narrative_from_tables(pkg["short_imper"]):
            md.append(f"- {line}")
        md.append("\n**Metric (\u00B0C, km/h)**")
        md.append(pkg["short_metric"].to_markdown(index=False))
        md.append("\n**Imperial (\u00B0F, mph)**")
        md.append(pkg["short_imper"].to_markdown(index=False))

        md.append(f"\n### Extended Forecast (through {(now_et + timedelta(days=7)).strftime('%b %d')}, 6h blocks)\n")
        md.append("**Metric (\u00B0C, km/h)**")
        md.append(pkg["long_metric"].to_markdown(index=False))
        md.append("\n**Imperial (\u00B0F, mph)**")
        md.append(pkg["long_imper"].to_markdown(index=False))
        return "\n".join(md)

    end_date_str = (now_et + timedelta(days=7)).strftime('%b %d, %Y')
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(f"> {STAMP_HUMAN}\n\n")
        f.write("# Weather / Forecasts – Lake Placid Area\n")
        f.write(f"**Valid through {end_date_str}**\n\n")
        f.write(section(pkg_dhi))
        f.write("\n\n---\n\n")
        f.write(section(pkg_xc))

def save_excel(pkg_dhi, pkg_xc, xlsx_path):
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        pkg_dhi["short_metric"].to_excel(writer, sheet_name="DHI_Short_Metric", index=False)
        pkg_dhi["short_imper"].to_excel(writer, sheet_name="DHI_Short_Imperial", index=False)
        pkg_dhi["long_metric"].to_excel(writer, sheet_name="DHI_Long_Metric", index=False)
        pkg_dhi["long_imper"].to_excel(writer, sheet_name="DHI_Long_Imperial", index=False)

        pkg_xc["short_metric"].to_excel(writer, sheet_name="XC_Short_Metric", index=False)
        pkg_xc["short_imper"].to_excel(writer, sheet_name="XC_Short_Imperial", index=False)
        pkg_xc["long_metric"].to_excel(writer, sheet_name="XC_Long_Metric", index=False)
        pkg_xc["long_imper"].to_excel(writer, sheet_name="XC_Long_Imperial", index=False)

        meta = pd.DataFrame({"Generated (ET)": [STAMP_HUMAN], "Notes": ["Whole numbers; NWS Gridpoint API"]})
        meta.to_excel(writer, sheet_name="Meta", index=False)

def build_pdf_full(pkg_dhi, pkg_xc, pdf_path):
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="Heading", fontSize=14, leading=16, spaceAfter=10))
    styles.add(ParagraphStyle(name="SubHeading", fontSize=12, leading=14, spaceAfter=8))
    styles.add(ParagraphStyle(name="NormalSmall", fontSize=8, leading=10, spaceAfter=5))
    styles.add(ParagraphStyle(name="StampRight", fontSize=8, leading=10, alignment=2, textColor=colors.grey))

    def make_table(df, cols, colWidths=None):
        data = [cols] + df[cols].astype(str).values.tolist()
        table = Table(data, colWidths=colWidths)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
            ('GRID', (0,0), (-1,-1), 0.25, colors.black),
            ('FONTSIZE', (0,0), (-1,-1), 7),
            ('ALIGN', (1,1), (-1,-1), 'CENTER')
        ]))
        return table

    def add_site(elements, pkg):
        elements.append(Paragraph(pkg["name"], styles["SubHeading"]))
        elements.append(Paragraph("Current Conditions", styles["SubHeading"]))

        c = pkg["current"]
        if c:
            lines = [
                f"Observed at {c['observed_at']}",
                f"Conditions: {c['conditions']}",
                f"Temp: {c['temp_c']} \u00B0C / {c['temp_f']} \u00B0F",
                f"Wind: {c['wind_mph']} mph{', gusting ' + str(c['gust_mph']) + ' mph' if c['gust_mph'] else ''}{' from ' + str(c['wind_dir_deg']) + '\u00B0' if c['wind_dir_deg'] is not None else ''}",
                f"RH: {c['rh']}%",
                f"Alerts: {pkg['alert_note']}",
                f'Source: <a href="{pkg["source_url"]}">NWS Gridpoint & Stations</a>'
            ]
            for ln in lines:
                elements.append(Paragraph(ln, styles["NormalSmall"]))
        else:
            elements.append(Paragraph("_No recent observation available from nearest station._", styles["NormalSmall"]))

        elements.append(Spacer(1, 6))
        elements.append(Paragraph("Short-term Forecast (36 Hours, Narrative)", styles["SubHeading"]))
        for ln in narrative_from_tables(pkg["short_imper"]):
            elements.append(Paragraph("• " + ln, styles["NormalSmall"]))
        elements.append(Spacer(1, 4))

        # Short-term tables
        elements.append(Paragraph("Metric (\u00B0C, km/h)", styles["NormalSmall"]))
        elements.append(make_table(pkg["short_metric"],
               ["Date/Time (EDT)","Temp (\u00B0C)","Feels (\u00B0C)","Wind (km/h)","Gusts (km/h)","Dir","RH (%)","PoP (%)","Cloud (%)"],
               [80,45,45,60,60,35,35,35,40]))
        elements.append(Spacer(1, 6))
        elements.append(Paragraph("Imperial (\u00B0F, mph)", styles["NormalSmall"]))
        elements.append(make_table(pkg["short_imper"],
               ["Date/Time (EDT)","Temp (\u00B0F)","Feels (\u00B0F)","Wind (mph)","Gusts (mph)","Dir","RH (%)","PoP (%)","Cloud (%)"],
               [80,45,45,60,60,35,35,35,40]))
        elements.append(Spacer(1, 10))

        # 7-day tables
        elements.append(Paragraph("Extended Forecast (7-Day)", styles["SubHeading"]))
        elements.append(Paragraph("Metric (\u00B0C, km/h)", styles["NormalSmall"]))
        elements.append(make_table(pkg["long_metric"],
               ["Date/Time (EDT)","Temp (\u00B0C)","Feels (\u00B0C)","Wind (km/h)","Gusts (km/h)","Dir","RH (%)","PoP (%)","Cloud (%)"],
               [80,45,45,60,60,35,35,35,40]))
        elements.append(Spacer(1, 6))
        elements.append(Paragraph("Imperial (\u00B0F, mph)", styles["NormalSmall"]))
        elements.append(make_table(pkg["long_imper"],
               ["Date/Time (EDT)","Temp (\u00B0F)","Feels (\u00B0F)","Wind (mph)","Gusts (mph)","Dir","RH (%)","PoP (%)","Cloud (%)"],
               [80,45,45,60,60,35,35,35,40]))
        elements.append(PageBreak())

    elements = []
    doc = SimpleDocTemplate(pdf_path, pagesize=(8.5*inch, 11*inch),
                            rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    elements.append(Paragraph(STAMP_HUMAN, styles["StampRight"]))
    elements.append(Paragraph("Weather / Forecasts – Lake Placid Area", styles["Heading"]))
    end_date_str = (now_et + timedelta(days=7)).strftime('%b %d, %Y')
    elements.append(Paragraph(f"Valid through {end_date_str}", styles["NormalSmall"]))
    elements.append(Spacer(1, 12))

    add_site(elements, pkg_dhi)
    add_site(elements, pkg_xc)
    doc.build(elements)

def build_pdf_summary(pkg_dhi, pkg_xc, pdf_path):
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="Heading", fontSize=14, leading=16, spaceAfter=10))
    styles.add(ParagraphStyle(name="SubHeading", fontSize=12, leading=14, spaceAfter=8))
    styles.add(ParagraphStyle(name="NormalSmall", fontSize=8, leading=10, spaceAfter=5))
    styles.add(ParagraphStyle(name="StampRight", fontSize=8, leading=10, alignment=2, textColor=colors.grey))

    def add_site(elements, pkg):
        elements.append(Paragraph(pkg["name"], styles["SubHeading"]))
        elements.append(Paragraph("Current Conditions", styles["SubHeading"]))
        c = pkg["current"]
        if c:
            wind_str = f"{c['wind_mph']} mph"
            if c['gust_mph']:
                wind_str += f", gusting {c['gust_mph']} mph"
            if c['wind_dir_deg'] is not None:
                wind_str += f" from {c['wind_dir_deg']}\u00B0"

            lines = [
                f"Observed at {c['observed_at']}",
                f"Conditions: {c['conditions']}",
                f"Temp: {c['temp_c']} \u00B0C / {c['temp_f']} \u00B0F",
                f"Wind: {wind_str}",
                f"RH: {c['rh']}%",
                f"Alerts: {pkg['alert_note']}",
                f'Source: <a href="{pkg["source_url"]}">NWS Gridpoint & Stations</a>'
            ]
        else:
            lines = ["No recent observation available from nearest station."]
        for ln in lines:
            elements.append(Paragraph(ln, styles["NormalSmall"]))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph("Short-term Forecast (36 Hours, Narrative)", styles["SubHeading"]))
        for ln in narrative_from_tables(pkg["short_imper"]):
            elements.append(Paragraph("• " + ln, styles["NormalSmall"]))
        elements.append(Spacer(1, 12))

    elements = []
    doc = SimpleDocTemplate(pdf_path, pagesize=(8.5*inch, 11*inch),
                            rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    elements.append(Paragraph(STAMP_HUMAN, styles["StampRight"]))
    elements.append(Paragraph("Weather / Forecasts – Lake Placid Area", styles["Heading"]))
    end_date_str = (now_et + timedelta(days=7)).strftime('%b %d, %Y')
    elements.append(Paragraph(f"Valid through {end_date_str}", styles["NormalSmall"]))
    elements.append(Spacer(1, 12))

    add_site(elements, pkg_dhi)
    add_site(elements, pkg_xc)
    doc.build(elements)

# ----------------------------
# MAIN
# ----------------------------
def main():
    try:
        # Build packages for both sites
        pkg_dhi = build_site_package("DHI – Mt. Whiteface Base", SITES["DHI – Mt. Whiteface Base"]["lat"], SITES["DHI – Mt. Whiteface Base"]["lon"], now_et)
        pkg_xc  = build_site_package("XC – Mt. Hoevenburg Base", SITES["XC – Mt. Hoevenburg Base"]["lat"],  SITES["XC – Mt. Hoevenburg Base"]["lon"], now_et)

        # Create output directory
        output_dir = os.path.join("Forecasts", STAMP_TAG)
        os.makedirs(output_dir, exist_ok=True)

        base = f"Lake_Placid_Forecast_{STAMP_TAG}"

        # WhatsApp Summary
        wa_path = os.path.join(output_dir, f"{base}_Whatsapp.txt")
        save_whatsapp_summary(pkg_dhi, pkg_xc, wa_path)

        # Markdown
        md_path = os.path.join(output_dir, f"{base}_Full.md")
        save_markdown(pkg_dhi, pkg_xc, md_path)

        # Excel
        xlsx_path = os.path.join(output_dir, f"{base}_Full.xlsx")
        save_excel(pkg_dhi, pkg_xc, xlsx_path)

        # PDFs
        pdf_full = os.path.join(output_dir, f"Lake_Placid_Forecast_With_Tables_{STAMP_TAG}.pdf")
        pdf_sum  = os.path.join(output_dir, f"Lake_Placid_Forecast_Summary_{STAMP_TAG}.pdf")
        build_pdf_full(pkg_dhi, pkg_xc, pdf_full)
        build_pdf_summary(pkg_dhi, pkg_xc, pdf_sum)

        print("\nDone. Files created in:", output_dir)
        print("  ", wa_path)
        print("  ", md_path)
        print("  ", xlsx_path)
        print("  ", pdf_full)
        print("  ", pdf_sum)
    except Exception as e:
        print(f"\nAn error occurred: {e}")

if __name__ == "__main__":
    main()
