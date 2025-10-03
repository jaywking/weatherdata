#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Report Generation
"""

import pandas as pd
from datetime import datetime, timedelta
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
import pytz

from forecast_processing import narrative_from_tables

ET = pytz.timezone("America/New_York")
now_et = datetime.now(ET)
STAMP_HUMAN = now_et.strftime("Generated: %b %d, %Y – %I:%M %p EDT")
STAMP_TAG   = now_et.strftime("%Y%m%d-%H%M")

def deg_to_cardinal(deg):
    """Convert degrees to a 16-point compass direction (e.g., N, NNE)."""
    if deg is None:
        return None
    try:
        val = float(deg) % 360
    except (TypeError, ValueError):
        return None
    directions = ["N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE", "S", "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW"]
    idx = int((val + 11.25) // 22.5) % 16
    return directions[idx]

def format_speed_imperial_metric(mph):
    """Return speed as 'X mph (Y km/h)' when mph is provided."""
    if mph is None:
        return None
    try:
        if hasattr(mph, 'item'):
            mph = mph.item()
        mph_val = float(mph)
    except (TypeError, ValueError):
        return None
    if pd.isna(mph_val):
        return None
    kmh = mph_val * 1.60934
    return f"{int(round(mph_val))} mph ({int(round(kmh))} km/h)"

def save_whatsapp_summary(pkg_dhi, pkg_xc, txt_path, stamp_human=None):
    def section(pkg):
        name = pkg["name"]
        curr = pkg["current"]
        wa = []
        wa.append(f"*{name}*")
        wa.append(f"*Current Conditions*")
        if curr:
            temp_str = f"{curr['temp_c']} °C / {curr['temp_f']} °F" if curr['temp_c'] is not None else "Not available"
            rh_str = f"{curr['rh']}%" if curr['rh'] is not None else "Not available"
            wind_value = curr['wind_mph']
            wind_base = None if wind_value is None or pd.isna(wind_value) else format_speed_imperial_metric(wind_value)
            direction = deg_to_cardinal(curr['wind_dir_deg'])
            gust_value = curr['gust_mph']
            gust_str = None
            if gust_value is not None and not pd.isna(gust_value) and gust_value != 0:
                gust_str = format_speed_imperial_metric(gust_value)
            wind_parts = []
            if wind_base:
                wind_parts.append(wind_base)
            if direction and wind_base:
                wind_parts.append(f"from {direction}")
            if gust_str:
                wind_parts.append(f"gusting {gust_str}")
            wind_line = ', '.join(wind_parts) if wind_parts else "-"
            wa.extend([
                f"- Conditions: {curr['conditions']}",
                f"- Temp: {temp_str}",
                f"- Wind: {wind_line}",
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

def save_whatsapp_combined(pkg_primary, pkg_secondary, txt_path, stamp_human=None):
    packages = [pkg_primary, pkg_secondary]
    stamp_to_use = stamp_human or STAMP_HUMAN

    lines = []
    lines.append('*Lake Placid Area Forecast*')
    lines.append(f'_{stamp_to_use}_')
    lines.append('')
    lines.append('*Current Conditions*')
    for pkg in packages:
        name = pkg['name']
        curr = pkg['current']
        lines.append(f"- *{name}*:")
        if curr:
            temp_str = 'Not available' if curr['temp_c'] is None or curr['temp_f'] is None else f"{curr['temp_c']} °C / {curr['temp_f']} °F"
            rh_str = 'Not available' if curr['rh'] is None else f"{curr['rh']}%"
            wind_value = curr['wind_mph']
            wind_base = None if wind_value is None or pd.isna(wind_value) else format_speed_imperial_metric(wind_value)
            direction = deg_to_cardinal(curr['wind_dir_deg'])
            gust_value = curr['gust_mph']
            gust_str = None
            if gust_value is not None and not pd.isna(gust_value) and gust_value != 0:
                gust_str = format_speed_imperial_metric(gust_value)
            wind_parts = []
            if wind_base:
                wind_parts.append(wind_base)
            if direction and wind_base:
                wind_parts.append(f"from {direction}")
            if gust_str:
                wind_parts.append(f"gusting {gust_str}")
            wind_str = ', '.join(wind_parts) if wind_parts else "-"
            lines.append(f"  {curr['conditions']};")
            lines.append(f"  Temp: {temp_str};")
            lines.append(f"  Wind: {wind_str};")
            lines.append(f"  RH: {rh_str}.")
        else:
            lines.append('  No recent observation available from nearest station.')
        lines.append(f"  Source: {pkg['source_url']}")
        lines.append('')

    lines.append('*Short-term Forecast (36h)*')
    for pkg in packages:
        lines.append(f"- *{pkg['name']}*")
        for detail in narrative_from_tables(pkg['short_imper']):
            lines.append(f"  - {detail}")

    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("\n".join(line for line in lines if line is not None))

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
            wdir = f" from {curr['wind_dir_deg']}°" if curr['wind_dir_deg'] is not None else ""
            md.extend([
                f"- Conditions: {curr['conditions']}",
                f"- Temp: {curr['temp_c']} °C / {curr['temp_f']} °F",
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
        md.append("\n**Metric (°C, km/h)**")
        md.append(pkg["short_metric"].to_markdown(index=False))
        md.append("\n**Imperial (°F, mph)**")
        md.append(pkg["short_imper"].to_markdown(index=False))

        md.append(f"\n### Extended Forecast (through {(now_et + timedelta(days=7)).strftime('%b %d')}, 6h blocks)\n")
        md.append("**Metric (°C, km/h)**")
        md.append(pkg["long_metric"].to_markdown(index=False))
        md.append("\n**Imperial (°F, mph)**")
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

    def current_conditions_block(pkg):
        curr = pkg["current"]
        lines = [f"<b>{pkg['name']}</b>:"]
        if curr:
            conditions = (curr.get("conditions") or "-") + ';'
            temp_c = curr.get("temp_c")
            temp_f = curr.get("temp_f")
            temp_available = temp_c is not None and not pd.isna(temp_c) and temp_f is not None and not pd.isna(temp_f)
            temp_line = f"Temp: {temp_c} °C / {temp_f} °F;" if temp_available else "Temp: Not available;"

            wind_val = curr.get("wind_mph")
            wind_base = None if wind_val is None or pd.isna(wind_val) else format_speed_imperial_metric(wind_val)
            dir_val = curr.get("wind_dir_deg")
            direction = deg_to_cardinal(dir_val) if dir_val is not None and not pd.isna(dir_val) else None
            gust_val = curr.get("gust_mph")
            gust_str = None
            if gust_val is not None and not pd.isna(gust_val) and gust_val != 0:
                gust_str = format_speed_imperial_metric(gust_val)
            wind_parts = []
            if wind_base:
                wind_parts.append(wind_base)
            if direction and wind_base:
                wind_parts.append(f"from {direction}")
            if gust_str:
                wind_parts.append(f"gusting {gust_str}")
            wind_line = f"Wind: {', '.join(wind_parts)};" if wind_parts else "Wind: -;"

            rh_val = curr.get("rh")
            rh_line = f"RH: {int(rh_val)}%." if rh_val is not None and not pd.isna(rh_val) else "RH: Not available."

            lines.extend([conditions, temp_line, wind_line, rh_line])
        else:
            lines.append("No recent observation available from nearest station.")
        source_url = pkg.get("source_url")
        if source_url:
            lines.append(f'Source: <a href="{source_url}">{source_url}</a>')
        return "<br/>".join(lines)

    packages = [pkg_dhi, pkg_xc]

    elements = []
    doc = SimpleDocTemplate(pdf_path, pagesize=(8.5*inch, 11*inch),
                            rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    elements.append(Paragraph(STAMP_HUMAN, styles["StampRight"]))
    elements.append(Paragraph("Weather / Forecasts - Lake Placid Area", styles["Heading"]))
    end_date_str = (now_et + timedelta(days=7)).strftime('%b %d, %Y')
    elements.append(Paragraph(f"Valid through {end_date_str}", styles["NormalSmall"]))
    elements.append(Spacer(1, 12))

    elements.append(Paragraph("Current Conditions", styles["SubHeading"]))
    for pkg in packages:
        elements.append(Paragraph(current_conditions_block(pkg), styles["NormalSmall"]))
        elements.append(Spacer(1, 6))

    elements.append(Paragraph("Short-term Forecast (36 Hours, Narrative)", styles["SubHeading"]))
    for pkg in packages:
        elements.append(Paragraph(f"<b>{pkg['name']}</b>", styles["NormalSmall"]))
        for detail in narrative_from_tables(pkg["short_imper"]):
            elements.append(Paragraph(f"&nbsp;&nbsp;- {detail}", styles["NormalSmall"]))
        elements.append(Spacer(1, 6))

    elements.append(Paragraph("Short-term Tables (36 Hours)", styles["SubHeading"]))
    for pkg in packages:
        elements.append(Paragraph(f"{pkg['name']} - Metric (°C, km/h)", styles["NormalSmall"]))
        elements.append(make_table(pkg["short_metric"],
               ["Date/Time (EDT)","Temp (°C)","Feels (°C)","Wind (km/h)","Gusts (km/h)","Dir","RH (%)","PoP (%)","Cloud (%)"],
               [80,45,45,60,60,35,35,35,40]))
        elements.append(Spacer(1, 6))
        elements.append(Paragraph(f"{pkg['name']} - Imperial (°F, mph)", styles["NormalSmall"]))
        elements.append(make_table(pkg["short_imper"],
               ["Date/Time (EDT)","Temp (°F)","Feels (°F)","Wind (mph)","Gusts (mph)","Dir","RH (%)","PoP (%)","Cloud (%)"],
               [80,45,45,60,60,35,35,35,40]))
        elements.append(Spacer(1, 10))

    elements.append(Paragraph("Extended Forecast (7-Day)", styles["SubHeading"]))
    for pkg in packages:
        elements.append(Paragraph(f"{pkg['name']} - Metric (°C, km/h)", styles["NormalSmall"]))
        elements.append(make_table(pkg["long_metric"],
               ["Date/Time (EDT)","Temp (°C)","Feels (°C)","Wind (km/h)","Gusts (km/h)","Dir","RH (%)","PoP (%)","Cloud (%)"],
               [80,45,45,60,60,35,35,35,40]))
        elements.append(Spacer(1, 6))
        elements.append(Paragraph(f"{pkg['name']} - Imperial (°F, mph)", styles["NormalSmall"]))
        elements.append(make_table(pkg["long_imper"],
               ["Date/Time (EDT)","Temp (°F)","Feels (°F)","Wind (mph)","Gusts (mph)","Dir","RH (%)","PoP (%)","Cloud (%)"],
               [80,45,45,60,60,35,35,35,40]))
        elements.append(Spacer(1, 10))

    doc.build(elements)

def build_pdf_summary(pkg_dhi, pkg_xc, pdf_path):
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="Heading", fontSize=14, leading=16, spaceAfter=10))
    styles.add(ParagraphStyle(name="SubHeading", fontSize=12, leading=14, spaceAfter=8))
    styles.add(ParagraphStyle(name="NormalSmall", fontSize=8, leading=10, spaceAfter=5))
    styles.add(ParagraphStyle(name="StampRight", fontSize=8, leading=10, alignment=2, textColor=colors.grey))

    def current_conditions_block(pkg):
        curr = pkg["current"]
        lines = [f"<b>{pkg['name']}</b>:"]
        if curr:
            conditions = (curr.get("conditions") or "-") + ';'
            temp_c = curr.get("temp_c")
            temp_f = curr.get("temp_f")
            temp_available = temp_c is not None and not pd.isna(temp_c) and temp_f is not None and not pd.isna(temp_f)
            temp_line = f"Temp: {temp_c} °C / {temp_f} °F;" if temp_available else "Temp: Not available;"

            wind_val = curr.get("wind_mph")
            wind_base = None if wind_val is None or pd.isna(wind_val) else format_speed_imperial_metric(wind_val)
            dir_val = curr.get("wind_dir_deg")
            direction = deg_to_cardinal(dir_val) if dir_val is not None and not pd.isna(dir_val) else None
            gust_val = curr.get("gust_mph")
            gust_str = None
            if gust_val is not None and not pd.isna(gust_val) and gust_val != 0:
                gust_str = format_speed_imperial_metric(gust_val)
            wind_parts = []
            if wind_base:
                wind_parts.append(wind_base)
            if direction and wind_base:
                wind_parts.append(f"from {direction}")
            if gust_str:
                wind_parts.append(f"gusting {gust_str}")
            wind_line = f"Wind: {', '.join(wind_parts)};" if wind_parts else "Wind: -;"

            rh_val = curr.get("rh")
            rh_line = f"RH: {int(rh_val)}%." if rh_val is not None and not pd.isna(rh_val) else "RH: Not available."

            lines.extend([conditions, temp_line, wind_line, rh_line])
        else:
            lines.append("No recent observation available from nearest station.")
        source_url = pkg.get("source_url")
        if source_url:
            lines.append(f'Source: <a href="{source_url}">{source_url}</a>')
        return "<br/>".join(lines)

    packages = [pkg_dhi, pkg_xc]

    elements = []
    doc = SimpleDocTemplate(pdf_path, pagesize=(8.5*inch, 11*inch),
                            rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    elements.append(Paragraph(STAMP_HUMAN, styles["StampRight"]))
    elements.append(Paragraph("Weather / Forecasts - Lake Placid Area", styles["Heading"]))
    end_date_str = (now_et + timedelta(days=7)).strftime('%b %d, %Y')
    elements.append(Paragraph(f"Valid through {end_date_str}", styles["NormalSmall"]))
    elements.append(Spacer(1, 12))

    elements.append(Paragraph("Current Conditions", styles["SubHeading"]))
    for pkg in packages:
        elements.append(Paragraph(current_conditions_block(pkg), styles["NormalSmall"]))
        elements.append(Spacer(1, 6))

    elements.append(Paragraph("Short-term Forecast (36 Hours, Narrative)", styles["SubHeading"]))
    for pkg in packages:
        elements.append(Paragraph(f"<b>{pkg['name']}</b>", styles["NormalSmall"]))
        for detail in narrative_from_tables(pkg["short_imper"]):
            elements.append(Paragraph(f"&nbsp;&nbsp;- {detail}", styles["NormalSmall"]))
        elements.append(Spacer(1, 12))

    doc.build(elements)
