#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
NWS API Interaction
"""

import re
import pytz
import requests
from datetime import datetime, timedelta
from dateutil import parser as dtparser
import pandas as pd

ET = pytz.timezone("America/New_York")
USER_AGENT = "AboveTheLineSafety Forecast Tool (contact: ops@example.com)"

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
    Return integer mph (average if a range). Handles km/h if present.
    """
    if not s or s.strip().lower() == "calm":
        return 0
    
    s_lower = s.lower()
    is_kmh = "km/h" in s_lower or "kmh" in s_lower
    
    nums = list(map(int, re.findall(r"\d+", s)))
    if not nums:
        return 0
        
    avg_val = sum(nums) / len(nums)
    
    if is_kmh:
        return int(round(avg_val * 0.621371))
    else: # Assume mph
        return int(round(avg_val))

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
    temp_c = p.get("temperature", {}).get("value")  # in °C
    rh = p.get("relativeHumidity", {}).get("value")
    wind_speed_obj = p.get("windSpeed", {})
    wind_val = wind_speed_obj.get("value")
    wind_unit = wind_speed_obj.get("unitCode", "")
    gust_speed_obj = p.get("windGust", {})
    gust_val = gust_speed_obj.get("value")
    gust_unit = gust_speed_obj.get("unitCode", "")
    wdir = p.get("windDirection", {}).get("value")   # degrees
    vis_m = p.get("visibility", {}).get("value")     # meters
    text = p.get("textDescription")
    pressure_pa = p.get("barometricPressure", {}).get("value")

    # Conversions
    temp_f = None if temp_c is None else (temp_c * 9/5 + 32)
    wind_mph = None
    if wind_val is not None:
        if wind_unit and "km_h" in wind_unit:
            wind_mph = wind_val * 0.621371  # km/h to mph
        else:
            wind_mph = wind_val * 2.23694      # m/s to mph (legacy)

    gust_mph = None
    if gust_val is not None:
        if gust_unit and "km_h" in gust_unit:
            gust_mph = gust_val * 0.621371  # km/h to mph
        else:
            gust_mph = gust_val * 2.23694      # m/s to mph (legacy)
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
    app = series_from_grid("apparentTemperature")# °F
    gust = series_from_grid("windGust")

    frames = [df for df in (sky, app, gust) if not df.empty]
    if not frames:
        return pd.DataFrame()

    out = pd.concat(frames, axis=1, join="outer")
    out = out.sort_index()
    return out
