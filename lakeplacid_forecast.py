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

import os
import pytz
from datetime import datetime

from forecast_processing import build_site_package
from reporting import save_whatsapp_summary, save_whatsapp_combined, save_markdown, save_excel, build_pdf_full, build_pdf_summary

# ----------------------------
# CONFIG
# ----------------------------
ET = pytz.timezone("America/New_York")

SITES = {
    "DHI – Mt. Whiteface Base": {"lat": 44.35, "lon": -73.86},
    "XC – Mt. Hoevenburg Base": {"lat": 44.2192, "lon": -73.9209},
}

# Output file base name will include ET timestamp
now_et = datetime.now(ET)
STAMP_TAG   = now_et.strftime("%Y%m%d-%H%M")  # for filenames

# ----------------------------
# MAIN
# ----------------------------
def main():
    total_tasks = 8
    task_num = 1
    print("Starting forecast generation...")
    try:
        # Build packages for both sites
        print(f"[{task_num}/{total_tasks}] Building data package for DHI – Mt. Whiteface Base...")
        pkg_dhi = build_site_package("DHI – Mt. Whiteface Base", SITES["DHI – Mt. Whiteface Base"]["lat"], SITES["DHI – Mt. Whiteface Base"]["lon"], now_et)
        task_num += 1

        print(f"[{task_num}/{total_tasks}] Building data package for XC – Mt. Hoevenburg Base...")
        pkg_xc  = build_site_package("XC – Mt. Hoevenburg Base", SITES["XC – Mt. Hoevenburg Base"]["lat"],  SITES["XC – Mt. Hoevenburg Base"]["lon"], now_et)
        task_num += 1

        # Create output directory
        output_dir = os.path.join("Forecasts", STAMP_TAG)
        os.makedirs(output_dir, exist_ok=True)

        base = f"Lake_Placid_Forecast_{STAMP_TAG}"

        # WhatsApp Summary
        print(f"[{task_num}/{total_tasks}] Generating WhatsApp summary...")
        wa_path = os.path.join(output_dir, f"{base}_Whatsapp.txt")
        save_whatsapp_summary(pkg_dhi, pkg_xc, wa_path)
        task_num += 1

        print(f"[{task_num}/{total_tasks}] Generating combined WhatsApp summary...")
        wa_combined_path = os.path.join(output_dir, f"{base}_Whatsapp_Combined.txt")
        save_whatsapp_combined(pkg_dhi, pkg_xc, wa_combined_path)
        task_num += 1

        # Markdown
        print(f"[{task_num}/{total_tasks}] Generating Markdown report...")
        md_path = os.path.join(output_dir, f"{base}_Full.md")
        save_markdown(pkg_dhi, pkg_xc, md_path)
        task_num += 1

        # Excel
        print(f"[{task_num}/{total_tasks}] Generating Excel report...")
        xlsx_path = os.path.join(output_dir, f"{base}_Full.xlsx")
        save_excel(pkg_dhi, pkg_xc, xlsx_path)
        task_num += 1

        # PDFs
        print(f"[{task_num}/{total_tasks}] Generating full PDF report...")
        pdf_full = os.path.join(output_dir, f"Lake_Placid_Forecast_With_Tables_{STAMP_TAG}.pdf")
        build_pdf_full(pkg_dhi, pkg_xc, pdf_full)
        task_num += 1
        
        print(f"[{task_num}/{total_tasks}] Generating summary PDF report...")
        pdf_sum  = os.path.join(output_dir, f"Lake_Placid_Forecast_Summary_{STAMP_TAG}.pdf")
        build_pdf_summary(pkg_dhi, pkg_xc, pdf_sum)

        print("\nDone. Files created in:", output_dir)
        print("  ", wa_path)
        print("  ", wa_combined_path)
        print("  ", md_path)
        print("  ", xlsx_path)
        print("  ", pdf_full)
        print("  ", pdf_sum)
    except Exception as e:
        print(f"\nAn error occurred: {e}")

if __name__ == "__main__":
    main()