#!/usr/bin/env python3
"""
Build an HTML report from Jira-style delivery metrics Excel exports where
openpyxl may fail (non-standard OOXML namespace). Reads all worksheets and
parses duration strings (e.g. 1w 2d 3h 4m) into hours for comparison.

Defaults: reads the latest `data/Delivery Metrics Comparison Analysis - YYYYMMDD.xlsx`
(by date in the filename) and writes `delivery_metrics_report__YYYYMMDD_HHMM.html` (run time)
plus `Delivery_Metrics_Report.html` (same content) under `output_report/` next to this script unless paths are passed.

Duration convention: 1w = 7×24h, 1d = 24h, 1h = 1h, 1m = 1/60h (calendar-style).
"""

from __future__ import annotations

import argparse
import html
import json
import re
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

NS_MAIN = "{http://purl.oclc.org/ooxml/spreadsheetml/main}"

# Stages we care about for dwell / watch (order = typical flow)
FLOW_STAGES = [
    "Backlog",
    "Analysis",
    "Ready for Development",
    "Blocked",
    "In Progress",
    "Ready for Review",
    "Review",
    "Ready for Deployment",
    "QA",
    "Done",
]

WATCH_STAGES = [
    "In Progress",
    "Ready for Review",
    "Review",
    "Blocked",
    "Ready for Deployment",
    "QA",
]

# Focus stages for “Top Spent Time by Key Stage” (Deploy/QA uses max of the two per issue when both exist)
TOP_DWELL_STAGES = [
    "In Progress",
    "Ready for Review",
    "Blocked",
    "__Ready_for_Deployment_or_QA",
]


def month_sort_key(label: str) -> Tuple[int, Any]:
    s = str(label).strip()
    try:
        dt = datetime.strptime(s, "%b %Y")
        return (0, dt)
    except ValueError:
        return (1, s)


def deployment_or_qa_hours(r: Dict[str, Any]) -> Optional[float]:
    dep = r.get("__h_Ready for Deployment")
    qa = r.get("__h_QA")
    vals = [x for x in (dep, qa) if x is not None]
    if not vals:
        return None
    return max(vals)


def cycle_time_hours(r: Dict[str, Any]) -> float:
    """Per issue: In Progress + Blocked + Ready for Review + Review + QA + Ready for Deployment (hours)."""
    total = 0.0
    for key in (
        "__h_In Progress",
        "__h_Blocked",
        "__h_Ready for Review",
        "__h_Review",
        "__h_QA",
        "__h_Ready for Deployment",
    ):
        v = r.get(key)
        if v is not None:
            total += float(v)
    return total


COMPARISON_XLSX_RE = re.compile(
    r"^Delivery Metrics Comparison Analysis - (\d{8})\.xlsx$",
    re.IGNORECASE,
)

# Default directory for generated HTML (created on run if missing).
OUTPUT_REPORT_DIR = "output_report"


def find_latest_comparison_xlsx(data_dir: Path) -> Path:
    """Use the workbook with the largest YYYYMMDD in the filename; tie-break by newer mtime."""
    pat = "Delivery Metrics Comparison Analysis - *.xlsx"
    candidates = list(data_dir.glob(pat))
    if not candidates:
        raise FileNotFoundError(f"No files matching {pat!r} under {data_dir}")

    def sort_key(p: Path) -> Tuple[int, float]:
        m = COMPARISON_XLSX_RE.match(p.name)
        stamp = int(m.group(1)) if m else -1
        return (stamp, p.stat().st_mtime)

    return max(candidates, key=sort_key)


def default_report_html_path(xlsx: Path, report_dir: Path) -> Path:
    """Default timestamped file: delivery_metrics_report__YYYYMMDD_HHMM.html (generation time)."""
    ts = datetime.now().strftime("%Y%m%d_%H%M")
    return report_dir / f"delivery_metrics_report__{ts}.html"


def col_to_idx(col: str) -> int:
    n = 0
    for c in col:
        n = n * 26 + (ord(c) - ord("A") + 1)
    return n - 1


def parse_shared_strings(z: zipfile.ZipFile) -> List[str]:
    try:
        data = z.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ET.fromstring(data)
    strings: List[str] = []
    for si in root.findall(f".//{NS_MAIN}si"):
        parts: List[str] = []
        for t in si.findall(f".//{NS_MAIN}t"):
            parts.append(t.text or "")
        strings.append("".join(parts))
    return strings


def cell_value(c: ET.Element, shared: List[str]) -> str:
    t = c.get("t")
    v = c.find(f"{NS_MAIN}v")
    if v is None or v.text is None:
        is_elem = c.find(f"{NS_MAIN}is")
        if is_elem is not None:
            ts = [x.text or "" for x in is_elem.findall(f".//{NS_MAIN}t")]
            return "".join(ts)
        return ""
    if t == "s":
        return shared[int(v.text)]
    return v.text or ""


def sheet_to_matrix(z: zipfile.ZipFile, sheet_path: str, shared: List[str]) -> List[List[str]]:
    root = ET.fromstring(z.read(sheet_path))
    rows: Dict[int, Dict[int, str]] = defaultdict(dict)
    for c in root.findall(f".//{NS_MAIN}c"):
        ref = c.get("r")
        if not ref:
            continue
        col_letters = "".join(filter(str.isalpha, ref))
        row_num = int("".join(filter(str.isdigit, ref)))
        col_idx = col_to_idx(col_letters)
        rows[row_num][col_idx] = cell_value(c, shared)
    if not rows:
        return []
    max_r = max(rows)
    max_c = max(max(cols) for cols in rows.values())
    matrix: List[List[str]] = []
    for r in range(1, max_r + 1):
        row = rows.get(r, {})
        matrix.append([row.get(c, "") for c in range(max_c + 1)])
    return matrix


def workbook_sheet_paths(z: zipfile.ZipFile) -> List[Tuple[str, str]]:
    """Return list of (sheet_name, path inside zip)."""
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels_root = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rid_to_target: Dict[str, str] = {}
    for rel in rels_root:
        if "Relationship" in rel.tag:
            rid = rel.get("Id")
            tgt = rel.get("Target")
            if rid and tgt:
                rid_to_target[rid] = tgt
    out: List[Tuple[str, str]] = []
    for sh in wb.findall(f".//{NS_MAIN}sheet"):
        name = sh.get("name") or "Sheet"
        rid = sh.get("{http://purl.oclc.org/ooxml/officeDocument/relationships}id")
        if not rid or rid not in rid_to_target:
            continue
        tgt = rid_to_target[rid]
        path = "xl/" + tgt.lstrip("/")
        if not path.startswith("xl/"):
            path = "xl/" + path
        out.append((name, path))
    return out


_DUR_RE = re.compile(r"(\d+)\s*([wdhm])", re.IGNORECASE)


def parse_duration_hours(text: Any) -> Optional[float]:
    if text is None:
        return None
    s = str(text).strip()
    if s in ("-", "", "nan", "NaT"):
        return None
    total = 0.0
    found = False
    for num, unit in _DUR_RE.findall(s):
        found = True
        n = int(num)
        u = unit.lower()
        if u == "m":
            total += n / 60.0
        elif u == "h":
            total += n
        elif u == "d":
            total += n * 24.0
        elif u == "w":
            total += n * 7.0 * 24.0
    if found:
        return total
    # Excel serial or ISO datetime as string
    try:
        if "T" in s and re.match(r"\d{4}-\d{2}-\d{2}", s):
            return None  # not a duration
    except Exception:
        pass
    return None


def fmt_days(hours: Optional[float]) -> str:
    """Format internal hour totals as calendar days for display (24h = 1d)."""
    if hours is None:
        return "—"
    days = hours / 24.0
    if days == 0:
        return "0d"
    if days < 0.01:
        return f"{days:.4f}d"
    if days < 1:
        return f"{days:.2f}d"
    if days < 100:
        return f"{days:.2f}d"
    return f"{days:.1f}d"


def dwell_approx_eq(a: Optional[float], b: float, tol_hours: float = 1.0) -> bool:
    if a is None:
        return False
    return abs(float(a) - float(b)) < tol_hours


def max_watch_heat_class(max_watch_hours: Any) -> str:
    """Background tiers for max sensitive-stage dwell (internal hours)."""
    h = float(max_watch_hours or 0)
    if h < 120:  # 3d–<5d
        return "mw-heat mw-h1"
    if h < 168:  # 5d–<7d
        return "mw-heat mw-h2"
    if h < 240:  # 7d–<10d
        return "mw-heat mw-h3"
    return "mw-heat mw-h4"


def format_resolved_date(val: Any) -> str:
    s = str(val).strip() if val is not None else ""
    if not s or s == "-":
        return "—"
    m = re.match(r"^(\d{4}-\d{2}-\d{2})", s)
    if m:
        return m.group(1)
    return s


def issue_summary_text(r: Dict[str, Any]) -> str:
    """Prefer spreadsheet Summary; fall back to Issue Type for older exports."""
    s = str(r.get("Summary", "")).strip()
    if s and s.lower() not in ("-", "nan", "none"):
        return s
    return str(r.get("Issue Type", "")).strip()


def matrix_to_records(matrix: List[List[str]], month_label: str) -> List[Dict[str, Any]]:
    if not matrix:
        return []
    header = [str(c).strip() for c in matrix[0]]
    # drop trailing empty headers
    while header and header[-1] == "":
        header.pop()
    idx = {h: i for i, h in enumerate(header) if h}
    recs: List[Dict[str, Any]] = []
    for row in matrix[1:]:
        if not any(str(x).strip() for x in row):
            continue
        key = row[idx["Key"]] if "Key" in idx else ""
        if not str(key).strip() or str(key).strip().lower() == "key":
            continue
        r: Dict[str, Any] = {"Month": month_label}
        for h in header:
            if h not in idx:
                continue
            val = row[idx[h]] if idx[h] < len(row) else ""
            r[h] = val
        recs.append(r)
    return recs


def percentile(sorted_vals: List[float], p: float) -> Optional[float]:
    if not sorted_vals:
        return None
    if len(sorted_vals) == 1:
        return sorted_vals[0]
    k = (len(sorted_vals) - 1) * p
    f = int(k)
    c = min(f + 1, len(sorted_vals) - 1)
    if f == c:
        return sorted_vals[f]
    d0 = sorted_vals[f] * (c - k)
    d1 = sorted_vals[c] * (k - f)
    return d0 + d1


@dataclass
class MonthlyAgg:
    month: str
    issue_count: int
    story_points: float
    stage_hours_sum: Dict[str, float]
    stage_hours_list: Dict[str, List[float]]
    avg_cycle_hours: float


def build_report(records: List[Dict[str, Any]], title: str) -> str:
    # enrich with numeric hours
    for r in records:
        r["__issue_summary"] = issue_summary_text(r)
        for st in FLOW_STAGES:
            if st in r:
                r[f"__h_{st}"] = parse_duration_hours(r.get(st))
        r["__h___Ready_for_Deployment_or_QA"] = deployment_or_qa_hours(r)
        r["__cycle_h"] = cycle_time_hours(r)
        r["__points"] = 0.0
        try:
            sp = str(r.get("Story Points", "")).strip()
            if sp and sp != "-":
                r["__points"] = float(sp)
        except ValueError:
            r["__points"] = 0.0

    # watch: max hours in WATCH_STAGES
    watch_rows: List[Tuple[float, str, Dict[str, Any]]] = []
    for r in records:
        key = str(r.get("Key", ""))
        hours_map = {st: r.get(f"__h_{st}") for st in WATCH_STAGES if f"__h_{st}" in r}
        vals = [v for v in hours_map.values() if v is not None]
        max_watch = max(vals) if vals else 0.0
        r["__max_watch_h"] = max_watch
        watch_rows.append((max_watch, key, r))
    watch_rows.sort(key=lambda x: -x[0])

    # monthly aggregates
    by_month: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    for r in records:
        by_month[str(r.get("Month", ""))].append(r)

    month_aggs: List[MonthlyAgg] = []
    for month in sorted(by_month.keys(), key=month_sort_key):
        rs = by_month[month]
        total_cycle_h = sum(float(x.get("__cycle_h", 0.0)) for x in rs)
        n_issues = len(rs)
        avg_cycle_h = total_cycle_h / n_issues if n_issues else 0.0
        agg = MonthlyAgg(
            month=month,
            issue_count=n_issues,
            story_points=sum(x["__points"] for x in rs),
            stage_hours_sum={s: 0.0 for s in FLOW_STAGES},
            stage_hours_list=defaultdict(list),
            avg_cycle_hours=avg_cycle_h,
        )
        for r in rs:
            for s in FLOW_STAGES:
                h = r.get(f"__h_{s}")
                if h is None:
                    continue
                agg.stage_hours_sum[s] += h
                agg.stage_hours_list[s].append(h)
            deff = r.get("__h___Ready_for_Deployment_or_QA")
            if deff is not None:
                agg.stage_hours_sum.setdefault("__DeployQA", 0.0)
                agg.stage_hours_list.setdefault("__DeployQA", []).append(deff)
        month_aggs.append(agg)

    # per-stage top dwellers (TOP_DWELL_STAGES): latest month in export only; pool top 15, HTML shows 10
    latest_month = month_aggs[-1].month if month_aggs else ""
    records_top_dwell = list(by_month.get(latest_month, [])) if latest_month else []
    top_by_stage: Dict[str, List[Tuple[float, str, str, str]]] = {}
    for st in TOP_DWELL_STAGES:
        ranked: List[Tuple[float, str, str, str]] = []
        for r in records_top_dwell:
            h = r.get(f"__h_{st}")
            if h is None or h <= 0:
                continue
            ranked.append(
                (
                    h,
                    str(r.get("Key", "")),
                    str(r.get("__issue_summary", "")),
                    str(r.get("Month", "")),
                )
            )
        ranked.sort(key=lambda x: -x[0])
        top_by_stage[st] = ranked[:15]

    def esc(s: Any) -> str:
        return html.escape(str(s), quote=True)

    parts: List[str] = []
    parts.append("<!DOCTYPE html><html><head><meta charset='utf-8'>")
    parts.append(f"<title>{esc(title)}</title>")
    parts.append(
        """<style>
*, *::before, *::after { box-sizing: border-box; }
html { overflow-x: hidden; }
body {
  font-family: system-ui, -apple-system, Segoe UI, sans-serif;
  margin: 0;
  padding: 24px clamp(16px, 3vw, 40px) 48px;
  color: #1a1a1a;
  width: 100%;
}
h1 { font-size: 1.35rem; }
h2 { font-size: 1.1rem; margin-top: 2rem; border-bottom: 1px solid #ddd; padding-bottom: 6px; }
h3 { font-size: 1rem; margin-top: 1.25rem; }
p.note { color: #555; font-size: 0.9rem; width: 100%; max-width: 100%; line-height: 1.45; }
table { border-collapse: collapse; width: 100%; max-width: 100%; margin: 12px 0; font-size: 0.88rem; table-layout: auto; }
th, td { border: 1px solid #ddd; padding: 8px 10px; text-align: left; }
.summary-col { max-width: 36rem; }
.summary-cell { max-width: 36rem; white-space: normal; word-break: break-word; line-height: 1.35; vertical-align: top; }
th { background: #f4f4f4; }
tr:nth-child(even) { background: #fafafa; }
.num { text-align: right; font-variant-numeric: tabular-nums; }
.badge { display: inline-block; padding: 2px 8px; border-radius: 999px; font-size: 0.75rem; background: #eef; }
.warn { background: #fff3cd; }
.th-sub { display: block; font-weight: 500; font-size: 0.72rem; color: #444; margin-top: 4px; line-height: 1.25; max-width: 100%; }
.chart-row {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: clamp(16px, 2vw, 32px);
  margin: 24px 0 16px;
  width: 100%;
  align-items: stretch;
}
.chart-box {
  width: 100%;
  min-width: 0;
  height: 460px;
  min-height: 460px;
  position: relative;
}
@media (max-width: 960px) {
  .chart-row { grid-template-columns: 1fr; }
  .chart-box { height: 460px; min-height: 460px; }
}
/* Watch table: max dwell heat + bottleneck stage */
.mw-heat { font-weight: 600; }
.mw-h1 { background: #fef9c3; color: #713f12; }
.mw-h2 { background: #fde047; color: #422006; }
.mw-h3 { background: #fdba74; color: #431407; }
.mw-h4 { background: #ef4444; color: #fff; }
.th-mw {
  background: linear-gradient(180deg, #fef9c3 0%, #fde68a 100%);
  color: #422006;
  font-weight: 600;
}
.th-sens { background: #ecfdf5; color: #14532d; font-weight: 600; }
.note-em { padding: 0.12em 0.4em; border-radius: 4px; font-weight: 600; }
.note-em-mw { background: #fef9c3; color: #713f12; }
/* Top Spent Time by Key Stage: top 5 = same hue (amber), stronger → lighter */
tr.dwell-rank-1 > td { background: #fbbf24 !important; }
tr.dwell-rank-2 > td { background: #fcd34d !important; }
tr.dwell-rank-3 > td { background: #fde68a !important; }
tr.dwell-rank-4 > td { background: #fef3c7 !important; }
tr.dwell-rank-5 > td { background: #fffbeb !important; }
/* Overall Monthly Summary: newest row (latest month) at top */
.monthly-summary-table tbody tr.summary-latest-month td {
  background: #c5e0cd;
  font-weight: 600;
  border-color: #6B9E78;
  box-shadow: inset 0px 0.5px #6B9E78;
}
.report-toolbar {
  position: sticky;
  top: 0;
  z-index: 50;
  display: flex;
  justify-content: flex-end;
  padding: 0 0 12px;
  margin: 0 0 8px;
  background: #fff;
  border-bottom: 1px solid #e8e8e8;
}
#btn-download-screenshot {
  font-family: inherit;
  font-size: 0.875rem;
  font-weight: 600;
  padding: 8px 16px;
  border-radius: 8px;
  border: 1px solid #003D4F;
  background: #003D4F;
  color: #fff;
  cursor: pointer;
}
#btn-download-screenshot:hover { filter: brightness(1.08); }
#btn-download-screenshot:disabled { opacity: 0.65; cursor: wait; }
</style></head><body>
<div class="report-toolbar"><button type="button" id="btn-download-screenshot" title="Save full report as PNG (charts included)">Download screenshot (PNG)</button></div>
<div id="report-capture-root">"""
    )
    parts.append(f"<h1>{esc(title)}</h1>")
    parts.append(
        "<p class='note'>Durations are parsed from the spreadsheet using "
        "<strong>1w = 7×24h</strong>, <strong>1d = 24h</strong>, then <strong>shown in days</strong> (24h = 1d). "
        "Issues are ranked by the longest single dwell among: In Progress, Ready for Review, Review, Blocked, "
        "Ready for Deployment, and QA (when present). Use this to decide where to add process attention.</p>"
    )

    # Chart series (chronological months)
    ch_labels: List[str] = []
    ch_issues: List[int] = []
    ch_points: List[float] = []
    ch_sp_per_issue: List[Optional[float]] = []
    ch_avg_ip: List[Optional[float]] = []
    ch_p90_ip: List[Optional[float]] = []
    ch_avg_rfr: List[Optional[float]] = []
    ch_avg_rev: List[Optional[float]] = []
    ch_avg_dep: List[Optional[float]] = []
    ch_avg_blocked: List[Optional[float]] = []
    ch_avg_cycle: List[Optional[float]] = []

    def h_to_d(x: Optional[float]) -> Optional[float]:
        if x is None:
            return None
        return round(x / 24.0, 4)

    for agg in month_aggs:
        ch_labels.append(agg.month)
        ch_issues.append(agg.issue_count)
        ch_points.append(round(agg.story_points, 1))
        if agg.issue_count:
            ch_sp_per_issue.append(round(agg.story_points / float(agg.issue_count), 2))
        else:
            ch_sp_per_issue.append(None)
        ip = agg.stage_hours_list.get("In Progress", [])
        rfr = agg.stage_hours_list.get("Ready for Review", [])
        rev = agg.stage_hours_list.get("Review", [])
        dep = agg.stage_hours_list.get("__DeployQA", [])
        blk = agg.stage_hours_list.get("Blocked", [])
        ip_sorted = sorted(ip)
        p90h = percentile(ip_sorted, 0.9) if ip_sorted else None
        avgh_ip = sum(ip) / len(ip) if ip else None
        avgh_rfr = sum(rfr) / len(rfr) if rfr else None
        avgh_rev = sum(rev) / len(rev) if rev else None
        avgh_dep = sum(dep) / len(dep) if dep else None
        avgh_blk = sum(blk) / len(blk) if blk else None
        ch_p90_ip.append(h_to_d(p90h))
        ch_avg_ip.append(h_to_d(avgh_ip))
        ch_avg_rfr.append(h_to_d(avgh_rfr))
        ch_avg_rev.append(h_to_d(avgh_rev))
        ch_avg_dep.append(h_to_d(avgh_dep))
        ch_avg_blocked.append(h_to_d(avgh_blk))
        ch_avg_cycle.append(h_to_d(agg.avg_cycle_hours))

    max_monthly_story_pts = max(ch_points) if ch_points else 0.0
    max_monthly_issues = max(ch_issues) if ch_issues else 0
    # Shared left axis on “Throughput & Story Points” chart: both Issues and story-point sum bars.
    throughput_y1_max = max(
        1.0, 2.0 * max(float(max_monthly_issues), float(max_monthly_story_pts))
    )
    throughput_y_issues_max = max(1, int(max_monthly_issues) * 2)

    chart_json = json.dumps(
        {
            "labels": ch_labels,
            "issues": ch_issues,
            "points": ch_points,
            "storyPointsPerIssue": ch_sp_per_issue,
            "throughputY1Max": throughput_y1_max,
            "throughputYMax": throughput_y_issues_max,
            "avgIpDays": ch_avg_ip,
            "p90IpDays": ch_p90_ip,
            "avgRfrDays": ch_avg_rfr,
            "avgReviewDays": ch_avg_rev,
            "avgDepDays": ch_avg_dep,
            "avgBlockedDays": ch_avg_blocked,
            "avgCycleDays": ch_avg_cycle,
        }
    )

    # Executive monthly table: newest month first; charts use chronological month order (unchanged).
    parts.append("<h2>Overall Monthly Summary</h2>")
    parts.append(
        "<p class='note'>In this table, months are listed <strong>newest first</strong> (latest month at the top; that row is lightly highlighted). "
        "The <strong>90% line</strong> column means: if you sorted every issue’s In Progress time from shortest to longest, "
        "about <strong>90% of issues</strong> would be at or below this value (only the slowest ~10% took longer). "
        "All dwell columns use <strong>days</strong> (24h = 1d).<br><br>"
        "<strong>Average Cycle Time</strong> is the mean, over all issues in that month, of each issue’s cycle time, "
        "where cycle time = In Progress + Blocked + Ready for Review + Review + QA + Ready for Deployment.</p>"
    )
    parts.append("<table class='monthly-summary-table'><thead><tr>")
    parts.append("<th>Month</th>")
    parts.append("<th>Issues</th>")
    parts.append("<th>Story points<br><span class='th-sub'>(sum)</span></th>")
    parts.append("<th>Avg In Progress<br><span class='th-sub'>(days)</span></th>")
    parts.append(
        "<th>In Progress — “90% line”<span class='th-sub' title='Same as 90th percentile: 9 in 10 issues finished In Progress in this time or less.'>"
        "9 in 10 issues spent ≤ this time (days)</span></th>"
    )
    parts.append("<th>Avg Ready for Review<br><span class='th-sub'>(days)</span></th>")
    parts.append("<th>Avg Review<br><span class='th-sub'>(days)</span></th>")
    parts.append("<th>Avg Blocked<br><span class='th-sub'>(days)</span></th>")
    parts.append(
        "<th>Avg Ready for Deploy / QA<span class='th-sub'>(days). "
        "Sheets use either Ready for Deployment or QA; we take the longer of the two per issue when both exist.</span></th>"
    )
    parts.append(
        "<th>Average Cycle Time<br><span class='th-sub'>(mean per issue, days)</span></th>"
    )
    parts.append("</tr></thead><tbody>")
    for row_i, agg in enumerate(reversed(month_aggs)):
        ip = agg.stage_hours_list.get("In Progress", [])
        rfr = agg.stage_hours_list.get("Ready for Review", [])
        dep = agg.stage_hours_list.get("__DeployQA", [])
        blk = agg.stage_hours_list.get("Blocked", [])
        rev = agg.stage_hours_list.get("Review", [])
        ip_sorted = sorted(ip)
        p90_ip = percentile(ip_sorted, 0.9) if ip_sorted else None
        avg_ip = sum(ip) / len(ip) if ip else None
        avg_rfr = sum(rfr) / len(rfr) if rfr else None
        avg_rev = sum(rev) / len(rev) if rev else None
        avg_dep = sum(dep) / len(dep) if dep else None
        avg_blk = sum(blk) / len(blk) if blk else None
        tr_cls = " class='summary-latest-month'" if row_i == 0 else ""
        parts.append(f"<tr{tr_cls}>")
        parts.append(f"<td>{esc(agg.month)}</td>")
        parts.append(f"<td class='num'>{agg.issue_count}</td>")
        parts.append(f"<td class='num'>{agg.story_points:.1f}</td>")
        parts.append(f"<td class='num'>{fmt_days(avg_ip)}</td>")
        parts.append(f"<td class='num'>{fmt_days(p90_ip)}</td>")
        parts.append(f"<td class='num'>{fmt_days(avg_rfr)}</td>")
        parts.append(f"<td class='num'>{fmt_days(avg_rev)}</td>")
        parts.append(f"<td class='num'>{fmt_days(avg_blk)}</td>")
        parts.append(f"<td class='num'>{fmt_days(avg_dep)}</td>")
        parts.append(f"<td class='num'>{fmt_days(agg.avg_cycle_hours)}</td>")
        parts.append("</tr>")
    parts.append("</tbody></table>")

    parts.append("<h3>Monthly Charts</h3>")
    parts.append("<div class='chart-row'>")
    parts.append("<div class='chart-box'><canvas id='chartVolume' aria-label='Issues and story points by month'></canvas></div>")
    parts.append(
        "<div class='chart-box'><canvas id='chartStages' aria-label='Average days in key stages including blocked by month'></canvas></div>"
    )
    parts.append("</div>")
    parts.append("<div class='chart-row'>")
    parts.append(
        "<div class='chart-box'><canvas id='chartStageMix' aria-label='Average dwell mix by month stacked'></canvas></div>"
    )
    parts.append(
        "<div class='chart-box'><canvas id='chartIssuesCycle' aria-label='Issues completed and average cycle time by month'></canvas></div>"
    )
    parts.append("</div>")
    parts.append(f"<script type='application/json' id='chart-data'>{chart_json}</script>")
    parts.append(
        """<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<script>
(function() {
  var el = document.getElementById('chart-data');
  if (!el || typeof Chart === 'undefined') return;
  var d = JSON.parse(el.textContent);
  var labels = d.labels;
  function d2(v) { return v == null ? null : Math.round(v * 1000) / 1000; }
  function d0(v) { return v == null || !isFinite(v) ? 0 : Math.round(v * 1000) / 1000; }
  new Chart(document.getElementById('chartVolume'), {
    type: 'bar',
    plugins: [{
      id: 'throughputVolumeBarLabels',
      afterDatasetsDraw: function(chart) {
        var ctx = chart.ctx;
        var datasets = chart.data.datasets;
        var styles = [
          { font: '600 10px system-ui, -apple-system, Segoe UI, sans-serif', color: '#6B9E78',
            fmt: function(v) { return String(Math.round(v)); } },
          { font: '600 10px system-ui, -apple-system, Segoe UI, sans-serif', color: '#47A1AD',
            fmt: function(v) {
              var r = Math.round(v * 10) / 10;
              return (Math.abs(r - Math.round(r)) < 1e-9) ? String(Math.round(r)) : String(r);
            } },
          { font: '600 10px system-ui, -apple-system, Segoe UI, sans-serif', color: '#634F7D',
            fmt: function(v) { return String(Math.round(v * 100) / 100); } }
        ];
        for (var di = 0; di < datasets.length; di++) {
          var meta = chart.getDatasetMeta(di);
          if (!meta || meta.hidden) continue;
          var ds = datasets[di];
          var st = styles[di] || styles[0];
          for (var i = 0; i < meta.data.length; i++) {
            var el = meta.data[i];
            if (!el || el.skip) continue;
            var raw = ds.data[i];
            if (raw == null) continue;
            var v = Number(raw);
            if (!isFinite(v)) continue;
            var x = el.x;
            var yTop = Math.min(el.y, el.base) - 3;
            ctx.save();
            ctx.textAlign = 'center';
            ctx.textBaseline = 'bottom';
            ctx.font = st.font;
            ctx.fillStyle = st.color;
            ctx.fillText(st.fmt(v), x, yTop);
            ctx.restore();
          }
        }
      }
    }],
    data: {
      labels: labels,
      datasets: [
        { label: 'Issues', data: d.issues, yAxisID: 'y', backgroundColor: '#6B9E78' },
        { label: 'Story points (sum)', data: d.points, yAxisID: 'y', backgroundColor: '#47A1AD' },
        { label: 'Story points / issue', data: d.storyPointsPerIssue, yAxisID: 'y1', backgroundColor: '#634F7D' }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      layout: { padding: { top: 24, right: 10, bottom: 8, left: 8 } },
      plugins: {
        title: { display: true, text: 'Delivery Throughput & Story Points by Month', font: { size: 15 } },
        legend: { position: 'bottom', labels: { boxWidth: 12, padding: 10, font: { size: 11 }, usePointStyle: true } }
      },
      scales: {
        x: { 
            ticks: { maxRotation: 0, font: { size: 12 } }, 
            grid: { display: false } 
        },
        y: { 
            type: 'linear', 
            position: 'left', 
            title: { display: true, text: 'Issues & Story Points', 
            font: { size: 12 } }, 
            beginAtZero: true, 
            max: d.throughputY1Max,
            grid: { drawOnChartArea: false }, 
            ticks: { font: { size: 11 }, display: false } 
        },
        y1: { 
            type: 'linear', 
            position: 'right', 
            title: { display: true, text: 'Story points per issue', 
            font: { size: 12 } }, 
            grid: { drawOnChartArea: true }, 
            beginAtZero: true, 
            max: 5,
            min: -6,
            ticks: { font: { size: 11 }, display: false, stepSize: 3} 
        }
      }
    }
  });
  new Chart(document.getElementById('chartStages'), {
    type: 'line',
    data: {
      labels: labels,
      datasets: [
        { label: 'Avg In Progress (d)', data: d.avgIpDays.map(d2), borderColor: '#003D4F', backgroundColor: '#003D4F', pointBackgroundColor: '#003D4F', pointBorderColor: '#003D4F', pointRadius: 4, pointHoverRadius: 6, borderWidth: 2, fill: false, tension: 0.2, spanGaps: true },
        { label: 'In Progress 90% line (d)', data: d.p90IpDays.map(d2), borderColor: '#47A1AD', backgroundColor: '#47A1AD', pointBackgroundColor: '#47A1AD', pointBorderColor: '#47A1AD', pointRadius: 4, pointHoverRadius: 6, borderWidth: 2, fill: false, borderDash: [4,4], tension: 0.2, spanGaps: true },
        { label: 'Avg Ready for Review (d)', data: d.avgRfrDays.map(d2), borderColor: '#CC850A', backgroundColor: '#CC850A', pointBackgroundColor: '#CC850A', pointBorderColor: '#CC850A', pointRadius: 4, pointHoverRadius: 6, borderWidth: 2, fill: false, tension: 0.2, spanGaps: true },
        { label: 'Avg Review (d)', data: d.avgReviewDays.map(d2), borderColor: '#6B9E78', backgroundColor: '#6B9E78', pointBackgroundColor: '#6B9E78', pointBorderColor: '#6B9E78', pointRadius: 4, pointHoverRadius: 6, borderWidth: 2, fill: false, tension: 0.2, spanGaps: true },
        { label: 'Avg Blocked (d)', data: d.avgBlockedDays.map(d2), borderColor: '#F2617A', backgroundColor: '#F2617A', pointBackgroundColor: '#F2617A', pointBorderColor: '#F2617A', pointRadius: 4, pointHoverRadius: 6, borderWidth: 2, fill: false, tension: 0.2, spanGaps: true },
        { label: 'Avg Ready for Deploy / QA (d)', data: d.avgDepDays.map(d2), borderColor: '#003D4F', backgroundColor: '#003D4F', pointBackgroundColor: '#003D4F', pointBorderColor: '#003D4F', pointRadius: 4, pointHoverRadius: 6, borderWidth: 2, fill: false, tension: 0.2, spanGaps: true },
        { label: 'Avg Cycle Time / issue (d)', data: d.avgCycleDays.map(d2), borderColor: '#634F7D', backgroundColor: '#634F7D', pointBackgroundColor: '#634F7D', pointBorderColor: '#634F7D', pointRadius: 4, pointHoverRadius: 6, fill: false, borderWidth: 2.5, tension: 0.2, spanGaps: true }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      layout: { padding: { top: 24, right: 10, bottom: 8, left: 8 } },
      plugins: {
        title: { display: true, text: 'Stage Dwell and Average Cycle Time (d)', font: { size: 15 } },
        legend: { position: 'bottom', labels: { boxWidth: 12, padding: 10, font: { size: 11 }, usePointStyle: true } }
      },
      scales: {
        x: { ticks: { maxRotation: 0, font: { size: 12 } }, grid: { display: false } },
        y: { beginAtZero: true, title: { display: true, text: 'Days', font: { size: 12 } }, ticks: { font: { size: 11 } } }
      }
    }
  });
  new Chart(document.getElementById('chartStageMix'), {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [
        { label: 'In Progress', data: d.avgIpDays.map(d0), backgroundColor: '#003D4F' },
        { label: 'Ready for Review', data: d.avgRfrDays.map(d0), backgroundColor: '#CC850A' },
        { label: 'Review', data: d.avgReviewDays.map(d0), backgroundColor: '#6B9E78' },
        { label: 'Blocked', data: d.avgBlockedDays.map(d0), backgroundColor: '#F2617A' },
        { label: 'Deploy / QA', data: d.avgDepDays.map(d0), backgroundColor: '#47A1AD' }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      layout: { padding: { top: 24, right: 10, bottom: 8, left: 8 } },
      plugins: {
        title: { display: true, text: 'Avg Dwell Mix by Month (Stacked days)', font: { size: 15 } },
        legend: { position: 'bottom', labels: { boxWidth: 12, padding: 10, font: { size: 11 }, usePointStyle: true } },
        tooltip: { mode: 'index', intersect: false }
      },
      scales: {
        x: { stacked: true, ticks: { maxRotation: 0, font: { size: 12 } }, grid: { display: false } },
        y: { stacked: true, beginAtZero: true, title: { display: true, text: 'Mean days per issue (sum of segments)', font: { size: 12 } }, ticks: { font: { size: 11 }, stepSize: 5 } }
      }
    }
  });
  new Chart(document.getElementById('chartIssuesCycle'), {
    type: 'bar',
    plugins: [{
      id: 'issuesCycleValueLabels',
      afterDatasetsDraw: function(chart) {
        var ctx = chart.ctx;
        var datasets = chart.data.datasets;
        for (var di = 0; di < datasets.length; di++) {
          var meta = chart.getDatasetMeta(di);
          if (!meta || meta.hidden) continue;
          var ds = datasets[di];
          for (var i = 0; i < meta.data.length; i++) {
            var el = meta.data[i];
            if (!el || el.skip) continue;
            var raw = ds.data[i];
            if (raw == null) continue;
            var v = Number(raw);
            if (!isFinite(v)) continue;
            var x = el.x;
            var yTop;
            if (di === 0) {
              yTop = Math.min(el.y, el.base) - 4;
              ctx.save();
              ctx.textAlign = 'center';
              ctx.textBaseline = 'bottom';
              ctx.font = '600 11px system-ui, -apple-system, Segoe UI, sans-serif';
              ctx.fillStyle = '#6B9E78';
              ctx.fillText(String(Math.round(v)), x, yTop);
              ctx.restore();
            } else {
              yTop = el.y - 10;
              ctx.save();
              ctx.textAlign = 'center';
              ctx.textBaseline = 'bottom';
              ctx.font = '600 10px system-ui, -apple-system, Segoe UI, sans-serif';
              ctx.fillStyle = '#634F7D';
              ctx.fillText((Math.round(v * 100) / 100) + ' d', x, yTop);
              ctx.restore();
            }
          }
        }
      }
    }],
    data: {
      labels: labels,
      datasets: [
        { type: 'bar', label: 'Issues', data: d.issues, yAxisID: 'y', backgroundColor: '#6B9E78' },
        { type: 'line', label: 'Avg cycle time / issue (d)', data: d.avgCycleDays.map(d2), yAxisID: 'y1', borderColor: '#634F7D', backgroundColor: '#634F7D', pointBackgroundColor: '#634F7D', pointBorderColor: '#634F7D', pointRadius: 4, pointHoverRadius: 6, borderWidth: 2.5, fill: false, tension: 0.2, spanGaps: true }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      layout: { padding: { top: 24, right: 10, bottom: 8, left: 8 } },
      plugins: {
        title: { display: true, text: 'Throughput vs Average Cycle Time', font: { size: 15 } },
        legend: { position: 'bottom', labels: { boxWidth: 12, padding: 10, font: { size: 11 }, usePointStyle: true } }
      },
      scales: {
        x: { ticks: { maxRotation: 0, font: { size: 12 } }, grid: { display: false } },
        y: { 
            type: 'linear', 
            position: 'left', 
            beginAtZero: true, 
            max: d.throughputYMax,
            title: { display: true, text: 'Issues', font: { size: 12 } }, 
            grid: { drawOnChartArea: false }, 
            ticks: { 
                font: { size: 11 },
                display: false 
            } 
        },
        y1: { 
            type: 'linear', 
            position: 'right', 
            beginAtZero: true, 
            max: 6,
            min: -5,
            title: { display: true, text: 'Avg cycle time / issue (d)', font: { size: 12 } }, 
            grid: { drawOnChartArea: false }, 
            ticks: { font: { size: 11 }, display: false } 
        }
      }
    }
  });
})();
</script>"""
    )

    # Watch list — one table per month (latest month first), no combined global table
    WATCH_THRESHOLD_H = 3.0 * 24.0  # 3 days
    parts.append("<h2>Issues Needing More Watch (Longest Dwell in Key Stage)</h2>")
    parts.append(
        "<p class='note'>Only issues where <span class='note-em note-em-mw'>max watch-stage dwell</span> "
        "reached at least <strong>3 days</strong> in <span class='note-em note-em-mw'>a single sensitive stage</span> "
        "are listed (others are omitted).</p>"
    )
    watch_by_month: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    for _mw, _key, r in watch_rows:
        watch_by_month[str(r.get("Month", ""))].append(r)
    month_order_watch = [agg.month for agg in reversed(month_aggs)]
    watch_cols = [
        "Key",
        "Summary",
        "Status",
        "Max watch stage days",
        "In Progress (d)",
        "Ready for Review (d)",
        "Review (d)",
        "Blocked (d)",
        "Ready for Deployment / QA (d)",
        "Resolved",
    ]
    for month in month_order_watch:
        rows_m = watch_by_month.get(month, [])
        rows_m.sort(key=lambda rr: -(rr.get("__max_watch_h") or 0.0))
        rows_m = [rr for rr in rows_m if (rr.get("__max_watch_h") or 0.0) >= WATCH_THRESHOLD_H]
        parts.append(f"<h3>{esc(month)}</h3>")
        parts.append("<table><thead><tr>")
        parts.append(f"<th>{esc('Key')}</th>")
        parts.append("<th class='summary-col'>Summary</th>")
        parts.append(f"<th>{esc('Status')}</th>")
        parts.append("<th class='th-mw'>Max watch stage days</th>")
        parts.append("<th class='th-sens'>In Progress (d)</th>")
        parts.append("<th class='th-sens'>Ready for Review (d)</th>")
        parts.append("<th class='th-sens'>Review (d)</th>")
        parts.append("<th class='th-sens'>Blocked (d)</th>")
        parts.append("<th class='th-sens'>Ready for Deployment / QA (d)</th>")
        parts.append(f"<th>{esc('Resolved')}</th>")
        parts.append("</tr></thead><tbody>")
        if not rows_m:
            parts.append(
                f"<tr><td colspan='{len(watch_cols)}'>"
                "No issues reached 3 days in a single sensitive stage this month.</td></tr>"
            )
        for r in rows_m:
            qa_h = r.get("__h_QA")
            dep_h = r.get("__h_Ready for Deployment")
            dep_qa = dep_h if dep_h is not None else qa_h
            mw = float(r.get("__max_watch_h") or 0.0)
            heat = max_watch_heat_class(mw)

            # Same tier background on Max watch and on whichever sensitive-stage value equals it
            def heat_if_bottleneck(h: Optional[float]) -> str:
                return f" {heat}" if dwell_approx_eq(h, mw) else ""

            dep_qa_heat = ""
            if dwell_approx_eq(dep_h, mw) or dwell_approx_eq(qa_h, mw):
                dep_qa_heat = f" {heat}"

            parts.append("<tr>")
            parts.append(f"<td><strong>{esc(r.get('Key'))}</strong></td>")
            parts.append(f"<td class='summary-cell'>{esc(r.get('__issue_summary', ''))}</td>")
            parts.append(f"<td>{esc(r.get('Status'))}</td>")
            parts.append(f"<td class='num {heat}'>{fmt_days(r.get('__max_watch_h'))}</td>")
            parts.append(f"<td class='num{heat_if_bottleneck(r.get('__h_In Progress'))}'>{fmt_days(r.get('__h_In Progress'))}</td>")
            parts.append(f"<td class='num{heat_if_bottleneck(r.get('__h_Ready for Review'))}'>{fmt_days(r.get('__h_Ready for Review'))}</td>")
            parts.append(f"<td class='num{heat_if_bottleneck(r.get('__h_Review'))}'>{fmt_days(r.get('__h_Review'))}</td>")
            parts.append(f"<td class='num{heat_if_bottleneck(r.get('__h_Blocked'))}'>{fmt_days(r.get('__h_Blocked'))}</td>")
            parts.append(f"<td class='num{dep_qa_heat}'>{fmt_days(dep_qa)}</td>")
            parts.append(f"<td>{esc(format_resolved_date(r.get('Resolved')))}</td>")
            parts.append("</tr>")
        parts.append("</tbody></table>")

    dwell_title = {
        "In Progress": "In Progress",
        "Ready for Review": "Ready for Review",
        "Blocked": "Blocked",
        "__Ready_for_Deployment_or_QA": "Ready for Deployment or QA",
    }
    dwell_last_col_header = {
        "In Progress": "In Progress (d)",
        "Ready for Review": "Ready for Review (d)",
        "Blocked": "Blocked (d)",
        "__Ready_for_Deployment_or_QA": "Ready for Deployment / QA (d)",
    }

    parts.append("<h2>Top Spent Time by Key Stage</h2>")
    parts.append(
        "<p class='note'><strong>How these tables are built:</strong> Only issues from the "
        f"<strong>latest month in this export ({esc(latest_month) if latest_month else '—'})</strong> are included. "
        "For each stage, we take that stage’s dwell time only (days shown; parsed from the sheet).<br><br>"
        "Issues with no time or zero in that stage are skipped. Rows are sorted by that stage’s dwell "
        "<strong>longest first</strong>; up to the <strong>top 10</strong> are shown per stage (the ranking pool keeps 15).<br>"
        "<strong>Ready for Deployment / QA</strong> uses, per issue, the longer of Ready for Deployment and QA when both exist.</p>"
    )
    for st in TOP_DWELL_STAGES:
        top = top_by_stage.get(st, [])
        if not top:
            continue
        parts.append(f"<h3>{esc(dwell_title.get(st, st))}</h3>")
        last_col = esc(dwell_last_col_header.get(st, f"{dwell_title.get(st, st)} (d)"))
        parts.append(
            f"<table><thead><tr><th>Key</th><th class='summary-col'>Summary</th><th>Month</th><th>{last_col}</th></tr></thead><tbody>"
        )
        for i, (h, key, summary, month) in enumerate(top[:10]):
            rk = i + 1
            tr_cls = f" class='dwell-rank-{rk}'" if rk <= 5 else ""
            parts.append(f"<tr{tr_cls}>")
            parts.append(f"<td><strong>{esc(key)}</strong></td>")
            parts.append(f"<td class='summary-cell'>{esc(summary)}</td>")
            parts.append(f"<td>{esc(month)}</td>")
            parts.append(f"<td class='num'>{fmt_days(h)}</td>")
            parts.append("</tr>")
        parts.append("</tbody></table>")

    parts.append("<p class='note'>Generated by delivery_metrics_report.py</p>")
    parts.append("</div>")
    parts.append(
        """<script src="https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js"></script>
<script>
(function() {
  var btn = document.getElementById('btn-download-screenshot');
  if (!btn) return;
  btn.addEventListener('click', function() {
    var root = document.getElementById('report-capture-root');
    if (!root || typeof html2canvas === 'undefined') {
      alert('Screenshot helper failed to load. Check your network and reload the page.');
      return;
    }
    btn.disabled = true;
    var label = btn.textContent;
    btn.textContent = 'Working…';
    var scale = Math.min(2, Math.max(1, window.devicePixelRatio || 1));
    html2canvas(root, { scale: scale, useCORS: true, logging: false, backgroundColor: '#ffffff' })
      .then(function(canvas) {
        canvas.toBlob(function(blob) {
          btn.disabled = false;
          btn.textContent = label;
          if (!blob) {
            alert('Could not create image.');
            return;
          }
          var url = URL.createObjectURL(blob);
          var a = document.createElement('a');
          a.href = url;
          a.download = 'delivery-metrics-report-' + new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-') + '.png';
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
          URL.revokeObjectURL(url);
        }, 'image/png', 0.92);
      })
      .catch(function(err) {
        btn.disabled = false;
        btn.textContent = label;
        alert('Screenshot failed: ' + (err && err.message ? err.message : String(err)));
      });
  });
})();
</script>
</body></html>"""
    )
    return "".join(parts)


def load_all_records(xlsx: Path) -> List[Dict[str, Any]]:
    all_recs: List[Dict[str, Any]] = []
    with zipfile.ZipFile(xlsx, "r") as z:
        shared = parse_shared_strings(z)
        sheets = workbook_sheet_paths(z)
        for name, path in sheets:
            matrix = sheet_to_matrix(z, path, shared)
            all_recs.extend(matrix_to_records(matrix, name))
    return all_recs


def main() -> None:
    script_dir = Path(__file__).resolve().parent
    data_dir = script_dir / "data"
    ap = argparse.ArgumentParser(
        description="Build HTML delivery dwell report from Delivery Metrics Comparison Analysis workbooks."
    )
    ap.add_argument(
        "xlsx",
        nargs="?",
        default=None,
        help="Path to metrics xlsx (default: latest Delivery Metrics Comparison Analysis - YYYYMMDD.xlsx in data/)",
    )
    ap.add_argument(
        "-o",
        "--output",
        default=None,
        help="Output HTML (default: under output_report/: timestamped file and Delivery_Metrics_Report.html)",
    )
    args = ap.parse_args()
    xlsx = Path(args.xlsx).expanduser().resolve() if args.xlsx else find_latest_comparison_xlsx(data_dir)
    if not xlsx.is_file():
        raise FileNotFoundError(f"Workbook not found: {xlsx}")
    report_dir = script_dir / OUTPUT_REPORT_DIR
    report_dir.mkdir(parents=True, exist_ok=True)
    out = (
        Path(args.output).expanduser().resolve()
        if args.output
        else default_report_html_path(xlsx, report_dir)
    )
    records = load_all_records(xlsx)
    title = f"Delivery Metrics Report — {xlsx.name}"
    html_out = build_report(records, title)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(html_out, encoding="utf-8")
    print(f"Wrote {out} ({len(records)} rows from {xlsx.name})")
    if args.output is None:
        static_default = report_dir / "Delivery_Metrics_Report.html"
        static_default.write_text(html_out, encoding="utf-8")
        print(f"Wrote {static_default} (same as timestamped report)")


if __name__ == "__main__":
    main()
