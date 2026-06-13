# AWD Field Data Toolkit

Local web app for AWD rice paddy field data analysis. Two tools:

1. **Pipe Locator** — upload a KMZ/KML map + the water monitoring workbook,
   get a table of which AWD pipes fall inside (or near) the gas chamber polygons.
2. **Water Level Plotter** — upload the workbook, tick the pipes you want,
   get a publication-style time series chart with the AWD threshold line.
   Downloadable as PNG.

Everything runs on your own computer. No data leaves your machine.

## Setup (one time)

You need Python 3.10+ installed (https://python.org — tick "Add to PATH" on Windows).

```
pip install -r requirements.txt
```

## Run

```
python app.py
```
…or double-click `run.bat` (Windows) / `./run.sh` (Mac/Linux).

Then open **http://localhost:5000** in your browser.

## What it understands automatically

- KMZ files (unzipped internally → doc.kml) and plain KML, any namespace
- Workbooks with a "Pipe Information"-style sheet (Pipe No. / Category / Latitude / Longitude),
  including the banner row on top
- Formula cells stored as text, e.g. `=22.5-15` → evaluated to 7.5
- Multi-row pipe blocks (Monitoring / Dry count / Picture / Time / Climate / Remarks) —
  only the Monitoring row is used
- Date-indexed sheets (e.g. "Valencia, Bukidnon") and DAT-indexed sheets
  (e.g. "Water Monitoring_Gas Area")
- Malformed coordinates like `125.150.100` (extra decimal points joined)
- "READ Sample" instruction sheets are skipped

## Options

- **Category filter**: AWD only (default) / CF only / all
- **Near-miss distance**: pipes within N meters of a polygon edge are flagged "near" (default 15 m)
- **AWD threshold line**: default −15 cm, configurable or blank for none
- **Date range**: optional from/to filter (date-axis sheets only)
- **Quick select**: all AWD / all CF / gas area pipes / clear

## Files

```
app.py            Flask server + chart generation
awd_parser.py     Excel / KML / KMZ parsing + cross-check logic
templates/index.html   The UI
```
