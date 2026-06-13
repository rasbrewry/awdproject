"""
AWD Field Data Toolkit — local Flask app
Run:  python app.py   ->  http://localhost:5000
"""

import io
import base64

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import datetime

from flask import Flask, request, jsonify, render_template

from awd_parser import (
    read_pipe_info,
    read_monitoring_sheets,
    read_kml_geometries,
    cross_check,
)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB


@app.route("/")
def index():
    return render_template("index.html")


# ------------------------------------------------------------------ inspect

@app.route("/api/inspect", methods=["POST"])
def inspect():
    """Upload an Excel file -> what's inside it (sheets, pipes, date ranges, metadata)."""
    f = request.files.get("excel")
    if not f:
        return jsonify({"error": "No Excel file received."}), 400
    data = io.BytesIO(f.read())
    try:
        info = read_pipe_info(data)
        data.seek(0)
        sheets = read_monitoring_sheets(data)
    except Exception as e:
        return jsonify({"error": f"Could not read workbook: {e}"}), 400

    out_sheets = []
    for name, d in sheets.items():
        labels = [l for l in d["labels"] if l]
        out_sheets.append({
            "name": name,
            "axis": d["axis"],
            "pipes": sorted(d["series"].keys()),
            "first": labels[0] if labels else None,
            "last": labels[-1] if labels else None,
            "n_columns": len(d["labels"]),
        })
    return jsonify({
        "filename": f.filename,
        "pipe_info": {
            "sheet": info["sheet"],
            "count": len(info["pipes"]),
            "categories": sorted({p["category"] for p in info["pipes"] if p["category"]}),
            "pipes": info["pipes"],
            "warnings": info["warnings"],
        },
        "monitoring_sheets": out_sheets,
    })


# ------------------------------------------------------------------ locate

@app.route("/api/locate", methods=["POST"])
def locate():
    """KMZ/KML + Excel -> which pipes are inside / near gas chamber polygons."""
    kml_f = request.files.get("kml")
    xls_f = request.files.get("excel")
    if not kml_f or not xls_f:
        return jsonify({"error": "Both a KMZ/KML file and an Excel file are required."}), 400

    category = request.form.get("category", "AWD").upper()
    try:
        threshold = float(request.form.get("threshold", 15))
    except ValueError:
        threshold = 15.0

    try:
        geoms = read_kml_geometries(kml_f.read(), kml_f.filename)
    except Exception as e:
        return jsonify({"error": f"Could not read KMZ/KML: {e}"}), 400
    if not geoms["polygons"]:
        return jsonify({"error": "No polygons found in the KMZ/KML. Gas chamber areas must be drawn as polygons."}), 400

    try:
        info = read_pipe_info(io.BytesIO(xls_f.read()))
    except Exception as e:
        return jsonify({"error": f"Could not read workbook: {e}"}), 400
    if not info["pipes"]:
        return jsonify({"error": info["warnings"][0] if info["warnings"] else "No pipes found in workbook."}), 400

    results = cross_check(info["pipes"], geoms, category=category, near_threshold_m=threshold)
    inside = [r for r in results if r["status"] == "inside"]
    near = [r for r in results if r["status"] == "near"]
    return jsonify({
        "polygons": [g["name"] for g in geoms["polygons"]],
        "summary": {
            "checked": len(results),
            "inside": sorted(r["pipe"] for r in inside),
            "near": sorted(r["pipe"] for r in near),
        },
        "results": results,
        "warnings": info["warnings"],
    })


# ------------------------------------------------------------------ plot

PALETTE = ["#2e7fd1", "#1a9e77", "#5b4fc4", "#d95f02", "#c2477f", "#b8860b", "#555555",
           "#17a2b8", "#8c564b", "#e377c2", "#7f7f7f", "#bcbd22"]


@app.route("/api/plot", methods=["POST"])
def plot():
    """Excel + selections -> time series PNG (base64)."""
    f = request.files.get("excel")
    if not f:
        return jsonify({"error": "No Excel file received."}), 400

    sheet = request.form.get("sheet", "")
    pipes_raw = request.form.get("pipes", "")
    try:
        wanted = [int(p) for p in pipes_raw.split(",") if p.strip()]
    except ValueError:
        return jsonify({"error": "Pipe list must be numbers separated by commas."}), 400
    if not wanted:
        return jsonify({"error": "Select at least one pipe."}), 400

    threshold_raw = request.form.get("threshold", "-15").strip()
    threshold = None
    if threshold_raw:
        try:
            threshold = float(threshold_raw)
        except ValueError:
            return jsonify({"error": "AWD threshold must be a number (e.g. -15)."}), 400

    date_from = request.form.get("date_from") or None
    date_to = request.form.get("date_to") or None

    try:
        sheets = read_monitoring_sheets(io.BytesIO(f.read()))
    except Exception as e:
        return jsonify({"error": f"Could not read workbook: {e}"}), 400
    if sheet not in sheets:
        return jsonify({"error": f"Sheet '{sheet}' not found or has no monitoring data."}), 400

    d = sheets[sheet]
    missing = [p for p in wanted if p not in d["series"]]
    if missing:
        return jsonify({"error": f"No data in this sheet for pipe(s): {', '.join(map(str, missing))}."}), 400

    # x axis
    if d["axis"] == "date":
        xs = [datetime.strptime(l, "%Y-%m-%d") if l else None for l in d["labels"]]
    else:
        xs = [float(l) if l else None for l in d["labels"]]

    # optional date filtering
    lo = datetime.strptime(date_from, "%Y-%m-%d") if (date_from and d["axis"] == "date") else None
    hi = datetime.strptime(date_to, "%Y-%m-%d") if (date_to and d["axis"] == "date") else None

    fig, ax = plt.subplots(figsize=(13, 6.5), dpi=140)
    plotted = 0
    for idx, pipe in enumerate(wanted):
        vals = d["series"][pipe]
        pts = [(x, v) for x, v in zip(xs, vals)
               if x is not None and v is not None
               and (lo is None or x >= lo) and (hi is None or x <= hi)]
        if not pts:
            continue
        px, pv = zip(*pts)
        style = "-" if idx < max(1, len(wanted) // 2) or len(wanted) <= 4 else "--"
        # group styling: first half solid, second half dashed (site clusters)
        ax.plot(px, pv, marker="o", markersize=3.5, linewidth=1.6,
                linestyle=style, color=PALETTE[idx % len(PALETTE)], label=f"Pipe {pipe}")
        plotted += 1
    if plotted == 0:
        plt.close(fig)
        return jsonify({"error": "No data points in the selected range."}), 400

    ax.axhline(0, color="#888", linewidth=1)
    if threshold is not None:
        ax.axhline(threshold, color="#a33", linewidth=1, linestyle=":")
        ax.annotate(f"AWD {threshold:g} cm", xy=(1.0, threshold), xycoords=("axes fraction", "data"),
                    xytext=(4, 0), textcoords="offset points", color="#a33", fontsize=9, va="center")

    ax.set_ylabel("Water level (cm, relative to pipe reference)")
    ax.set_xlabel("Date" if d["axis"] == "date" else "Days after transplant (DAT)")
    title_range = ""
    if d["axis"] == "date":
        shown_x = [x for x, v in zip(xs, [1]*len(xs)) if x is not None]
        if lo or hi:
            title_range = f" ({date_from or 'start'} – {date_to or 'end'})"
    ax.set_title(f"Water level — Pipes {', '.join(map(str, wanted))}{title_range}\n[{sheet}]")
    ax.grid(alpha=0.25)
    ax.legend(ncol=min(7, max(1, plotted)), fontsize=9,
              loc="upper center", bbox_to_anchor=(0.5, -0.16), frameon=False)
    if d["axis"] == "date":
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %d"))
        fig.autofmt_xdate(rotation=40)
    fig.tight_layout()

    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    return jsonify({"png": base64.b64encode(buf.getvalue()).decode()})


if __name__ == "__main__":
    print("\n  AWD Field Data Toolkit  ->  http://localhost:5000\n")
    app.run(host="127.0.0.1", port=5000, debug=False)
