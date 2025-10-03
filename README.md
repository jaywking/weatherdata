# Lake Placid Forecast Generator

A Python utility that pulls National Weather Service (NWS/NOAA) data and produces multi-format forecast packages for the Lake Placid area.

## Features

- Fetches latest observations plus hourly and grid forecasts from weather.gov for multiple sites.
- Builds consolidated metric and imperial tables for the next 36 hours and seven days.
- Exports WhatsApp-friendly text, Markdown, Excel, and PDF summaries in one run.
- Uses a consistent file naming scheme with Eastern Time (ET) timestamps for easy archiving.
- Provides real-time progress updates during execution.

## Requirements

- Python 3.9+
- Packages: `requests`, `pandas`, `openpyxl`, `reportlab`, `pytz`, `python-dateutil`

Install dependencies with:

```bash
pip install requests pandas openpyxl reportlab pytz python-dateutil
```

## Getting Started

1. Clone or download this repository.
2. Install the Python dependencies (see above).
3. From the project root, run:

```bash
python lakeplacid_forecast.py
```

The script pulls data for the configured sites and creates a timestamped folder under `Forecasts/` containing:

- `*_Whatsapp.txt` – short-form narrative and current conditions per site.
- `*_Whatsapp_Combined.txt` – single message covering both sites with source links.
- `*_Full.md` – comprehensive Markdown report with metric and imperial tables.
- `*_Full.xlsx` – Excel workbook with site sheets and a metadata tab.
- `Lake_Placid_Forecast_With_Tables_*.pdf` – printable full report.
- `Lake_Placid_Forecast_Summary_*.pdf` – condensed summary PDF.

## Configuration

Key customisations live at the top of `lakeplacid_forecast.py`:

- `SITES` – add or adjust latitude/longitude pairs for additional locations.
- `USER_AGENT` – override with your contact info to stay within NWS API guidelines.
- Date/time formatting: the script timestamps output in Eastern Time (`America/New_York`).

## Data Sources

- [api.weather.gov](https://api.weather.gov/) gridpoint, hourly forecast, and station observation endpoints.
- Derived products are cached locally only through the generated files; no database is required.

## Troubleshooting

- Ensure outbound network access to `api.weather.gov`.
- If you see pandas `FutureWarning` messages during execution, the run still completes; future library updates may require small adjustments.
- ReportLab requires system fonts capable of rendering the bullet and degree symbols; the defaults on most systems suffice.

## Changelog

### 2025-10-03
- **Fix**: Corrected wind speed calculation for current observations to handle km/h units from the NWS API, resolving data inaccuracies.
- **Enhancement**: Added progress indicators to provide real-time feedback during script execution.
- **Enhancement**: Improved robustness of hourly forecast parsing to handle different wind speed units.

## Contributing

Issues and pull requests are welcome. When adding new outputs or sites, keep filenames consistent with the existing timestamp pattern to simplify automation.

## License

Add your preferred license here (e.g., MIT).