# Delivery metrics report by a specific input file

Python script **`delivery_metrics_report.py`** reads a **Delivery Metrics Comparison Analysis** Excel export and builds one self-contained **HTML dashboard**: summary tables, monthly charts (Chart.js), issues to watch, and top dwell time by stage.

The workbook is parsed as OOXML (ZIP + XML) using only the **Python standard library**, so you do not need `pandas`, `openpyxl`, or other pip packages. That also helps with some workbooks that use non-standard OOXML namespaces.

---

## What you need

- **Python 3** (3.9 or newer recommended).
- A **web browser** to open the generated HTML. The report loads **Chart.js** and **html2canvas** from a CDN the first time; use an online connection when opening the file, or host the same JS files locally if you work offline.

---

## Input files

Put comparison workbooks under the **`data/`** folder next to the script, named like:

`Delivery Metrics Comparison Analysis - YYYYMMDD.xlsx`

If you run the script **without** a path argument, it chooses the workbook whose **`YYYYMMDD` in the filename is largest**. If two files share the same date, the newer file on disk wins.

---

## How to run

From this project directory:

```bash
python3 delivery_metrics_report.py
```

**Typical outputs** (under `output_report/`):

- **`delivery_metrics_report.html`** — stable name; good for bookmarks or “latest report” links.  
- **`delivery_metrics_report__YYYYMMDD_HHMM.html`** — same content, timestamped when the script ran.

Use a **specific** workbook:

```bash
python3 delivery_metrics_report.py "data/Delivery Metrics Comparison Analysis - 20260430.xlsx"
```

Write **only** one HTML file (no second copy under `output_report/`):

```bash
python3 delivery_metrics_report.py -o /path/to/custom_report.html
```

---

## What appears in the report

- **Overall Monthly Summary** — per-month issues, story points, average dwell in key stages, 90% line for In Progress, and average cycle time. Months are listed **newest first**; the **latest month row is highlighted**.  
- **Monthly charts** — throughput & story points, stage dwell lines, stacked dwell mix, throughput vs cycle time. Month order on the **chart axis** stays **chronological** (oldest → newest).  
- **Issues Needing More Watch** — sensitive stages with long dwell; grouped by month, **newest month first**.  
- **Top Spent Time by Key Stage** — ranked dwell for each key stage using **only the latest month** in the export (table still includes a **Month** column).  
- **Download screenshot (PNG)** — captures the main report area for sharing or slides.

---

## Durations in the sheet

Durations are interpreted in a **calendar-style** way: **1w = 7×24h**, **1d = 24h**, **1h = 1h**, **1m = 1/60 hour**. Values are shown in **days** (24h = 1 day) where the report says “days.”

---

## Project files

- **`delivery_metrics_report.py`** — generator.  
- **`data/`** — drop your comparison `.xlsx` files here.  
- **`output_report/`** — generated HTML lands here by default.

For field-level definitions and chart logic, see the **module docstring** and comments at the top of `delivery_metrics_report.py`.