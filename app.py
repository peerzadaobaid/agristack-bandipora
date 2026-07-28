"""
AGRISTACK Dashboard — District Bandipora
Reads the newest snapshot from snapshots/YYYY-MM-DD.xlsx and shows:
  - Tehsil-wise progress with % completion
  - Patwari-wise progress (two sorts: by %, by count)
  - Villages not started
Also serves a password-gated Excel download.
"""
import os
import re
import io
import sys
from datetime import datetime, date

import pandas as pd
from flask import Flask, render_template, request, Response, send_file
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)

SNAPSHOTS_DIR = "snapshots"
LEGACY_EXCEL = "AGRISTACK.xlsx"

# Tolerant column matchers
COL_MATCHERS = {
    "tehsil":    re.compile(r"tehsil", re.I),
    "village":   re.compile(r"^village$", re.I),
    "patwari":   re.compile(r"patwari", re.I),
    "khasras":   re.compile(r"khasra|survey", re.I),
    "submitted": re.compile(r"^submitted$", re.I),
}


# ---------- Snapshot discovery ----------

def newest_snapshot():
    """Return (date, path) of the newest snapshot. Falls back to legacy file."""
    if os.path.isdir(SNAPSHOTS_DIR):
        candidates = []
        for name in os.listdir(SNAPSHOTS_DIR):
            if not name.endswith(".xlsx") or name.startswith("~"):
                continue
            stem = name[:-5]
            try:
                d = datetime.strptime(stem, "%Y-%m-%d").date()
            except ValueError:
                continue
            candidates.append((d, os.path.join(SNAPSHOTS_DIR, name)))
        if candidates:
            candidates.sort(key=lambda x: x[0], reverse=True)
            return candidates[0]
    if os.path.exists(LEGACY_EXCEL):
        mtime = os.path.getmtime(LEGACY_EXCEL)
        return (datetime.fromtimestamp(mtime).date(), LEGACY_EXCEL)
    return None


# ---------- Excel loading ----------

_df_cache = {}  # path -> DataFrame


def _pick_sheet(xl):
    """Prefer Sheet2, then MAIN, then first sheet with all 5 required columns."""
    scored = []
    for name in xl.sheet_names:
        try:
            head = pd.read_excel(xl, sheet_name=name, nrows=0)
        except Exception:
            continue
        keys = [str(c) for c in head.columns]
        score = 0
        for logical, rx in COL_MATCHERS.items():
            for k in keys:
                if rx.search(k.strip()):
                    if logical == "khasras" and re.search(r"submitted", k, re.I):
                        continue
                    score += 1
                    break
        rank = 0 if name == "Sheet2" else 1 if name == "MAIN" else 2
        scored.append((score, -rank, name))
    scored.sort(reverse=True)
    if not scored:
        raise RuntimeError("No usable sheets found.")
    picked = scored[0][2]
    print(f"[sheet-picker] Picked '{picked}' from {[n for _, _, n in scored]}", file=sys.stderr)
    return picked


def _detect_columns(df):
    cols = {}
    for logical, rx in COL_MATCHERS.items():
        for c in df.columns:
            key = str(c).strip()
            if rx.search(key):
                if logical == "khasras" and re.search(r"submitted", key, re.I):
                    continue
                cols[logical] = c
                break
    missing = [k for k in COL_MATCHERS if k not in cols]
    if missing:
        raise RuntimeError(f"Missing required columns: {', '.join(missing)}")
    return cols


def load_snapshot(path):
    """Load and normalize a snapshot file. Cached by path."""
    if path in _df_cache:
        return _df_cache[path]
    xl = pd.ExcelFile(path)
    sheet = _pick_sheet(xl)
    df = pd.read_excel(xl, sheet_name=sheet)
    cols = _detect_columns(df)
    df = df.rename(columns={
        cols["tehsil"]: "tehsil",
        cols["village"]: "village",
        cols["patwari"]: "patwari",
        cols["khasras"]: "khasras",
        cols["submitted"]: "submitted",
    })[["tehsil", "village", "patwari", "khasras", "submitted"]].copy()
    df["khasras"] = pd.to_numeric(df["khasras"], errors="coerce").fillna(0).astype(int)
    df["submitted"] = pd.to_numeric(df["submitted"], errors="coerce").fillna(0).astype(int)
    for col in ("tehsil", "village", "patwari"):
        df[col] = df[col].astype(str).str.strip()
    df = df[(df["village"] != "") & (df["village"].str.lower() != "nan")]
    _df_cache[path] = df
    print(f"[load] {path}: {len(df)} rows, sheet '{sheet}'", file=sys.stderr)
    return df


# ---------- View building ----------

def _pct(sub, tot):
    return (sub / tot * 100) if tot > 0 else 0.0


def build_views(df):
    """Aggregate into all dashboard views."""
    # --- Tehsil-wise (sorted by submitted desc) ---
    t = (df.groupby("tehsil", as_index=False)
           .agg(villages=("village", "count"),
                total=("khasras", "sum"),
                submitted=("submitted", "sum"))
           .sort_values("submitted", ascending=False))
    tehsil_rows = []
    for _, row in t.iterrows():
        tehsil_rows.append({
            "tehsil": row["tehsil"],
            "villages": int(row["villages"]),
            "total": int(row["total"]),
            "submitted": int(row["submitted"]),
            "pct": _pct(int(row["submitted"]), int(row["total"])),
        })

    # --- Patwari-wise (two sorts) ---
    p = (df.groupby("patwari", as_index=False)
           .agg(villages_list=("village", lambda s: ", ".join(sorted(s.tolist()))),
                total=("khasras", "sum"),
                submitted=("submitted", "sum")))
    p["pct"] = p.apply(lambda r: _pct(int(r["submitted"]), int(r["total"])), axis=1)

    def _records(frame):
        return [{
            "patwari": r["patwari"],
            "villages_list": r["villages_list"],
            "total": int(r["total"]),
            "submitted": int(r["submitted"]),
            "pct": float(r["pct"]),
        } for _, r in frame.iterrows()]

    patwari_by_pct = _records(p.sort_values(["pct", "submitted"], ascending=[False, False]))
    patwari_by_count = _records(p.sort_values(["submitted", "pct"], ascending=[False, False]))

    # --- Not started ---
    ns = df[df["submitted"] == 0].sort_values(["tehsil", "village"])
    not_started = ns.to_dict("records")

    top_patwari = patwari_by_count[0] if patwari_by_count else None
    top_tehsil = tehsil_rows[0] if tehsil_rows else None

    grand = {
        "tehsils": int(len(t)),
        "villages": int(len(df)),
        "patwaris": int(len(p)),
        "total_khasras": int(df["khasras"].sum()),
        "submitted": int(df["submitted"].sum()),
        "not_started": int(len(ns)),
        "overall_pct": _pct(int(df["submitted"].sum()), int(df["khasras"].sum())),
    }
    return {
        "tehsil_rows": tehsil_rows,
        "patwari_by_pct": patwari_by_pct,
        "patwari_by_count": patwari_by_count,
        "not_started": not_started,
        "top_patwari": top_patwari,
        "top_tehsil": top_tehsil,
        "grand": grand,
    }


# ---------- Header meta ----------

def read_as_of_override():
    if os.path.exists("AS_OF.txt"):
        with open("AS_OF.txt", "r", encoding="utf-8") as f:
            text = f.read().strip()
        if text:
            return text
    return None


def format_date(d):
    return d.strftime("%d %b %Y")


# ---------- Routes ----------

@app.route("/")
def index():
    snap = newest_snapshot()
    if not snap:
        return "No snapshots found. Add a file to snapshots/ folder.", 500
    snap_date, snap_path = snap
    df = load_snapshot(snap_path)
    views = build_views(df)

    override = read_as_of_override()
    as_of_display = override if override else format_date(snap_date)

    return render_template("index.html", as_of=as_of_display, **views)


def _check_download_password():
    expected = os.environ.get("DOWNLOAD_PASSWORD", "")
    if not expected:
        return False
    auth = request.authorization
    if not auth or auth.password != expected:
        return False
    return True


@app.route("/download.xlsx")
def download():
    if not _check_download_password():
        return Response(
            "Authentication required",
            401,
            {"WWW-Authenticate": 'Basic realm="AgriStack Download"'},
        )
    snap = newest_snapshot()
    if not snap:
        return "No snapshots found.", 500
    snap_date, snap_path = snap
    df = load_snapshot(snap_path)
    views = build_views(df)
    buf = _generate_workbook(views, snap_date)
    filename = f"AGRISTACK_Dashboard_{snap_date.strftime('%Y-%m-%d')}.xlsx"
    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename,
    )


def _generate_workbook(views, snap_date):
    """Generate an xlsx mirroring the dashboard."""
    hdr_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    hdr_fill = PatternFill(start_color="1F3F2E", end_color="1F3F2E", fill_type="solid")
    hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    body_font = Font(name="Arial", size=11)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    tot_font = Font(name="Arial", size=11, bold=True)
    tot_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    ns_fill = PatternFill(start_color="A54B2A", end_color="A54B2A", fill_type="solid")
    thin = Side(border_style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def write_hdr(ws, headers, fill=hdr_fill):
        for i, h in enumerate(headers, start=1):
            c = ws.cell(row=1, column=i, value=h)
            c.font = hdr_font; c.fill = fill; c.alignment = hdr_align; c.border = border
        ws.row_dimensions[1].height = 38

    wb = Workbook()
    grand = views["grand"]

    # === TEHSIL WISE ===
    ws1 = wb.active
    ws1.title = "TEHSIL WISE"
    write_hdr(ws1, ["S.NO", "TEHSIL", "NUMBER OF VILLAGES",
                    "TOTAL SURVEY NOS", "SUBMITTED", "% COMPLETION"])
    for i, row in enumerate(views["tehsil_rows"], start=1):
        r = i + 1
        vals = [i, row["tehsil"], row["villages"], row["total"], row["submitted"], round(row["pct"], 2)]
        aligns = [center, left, center, center, center, center]
        for ci, (v, a) in enumerate(zip(vals, aligns), start=1):
            c = ws1.cell(row=r, column=ci, value=v)
            c.alignment = a; c.font = body_font; c.border = border
            if ci in (3, 4, 5): c.number_format = "#,##0"
            if ci == 6: c.number_format = '0.00"%"'
    tt = len(views["tehsil_rows"]) + 2
    ws1.cell(row=tt, column=2, value="TOTAL").alignment = left
    ws1.cell(row=tt, column=3, value=grand["villages"]).alignment = center
    ws1.cell(row=tt, column=4, value=grand["total_khasras"]).alignment = center
    ws1.cell(row=tt, column=5, value=grand["submitted"]).alignment = center
    ws1.cell(row=tt, column=6, value=round(grand["overall_pct"], 2)).alignment = center
    for c in range(1, 7):
        cc = ws1.cell(row=tt, column=c)
        cc.font = tot_font; cc.fill = tot_fill; cc.border = border
        if c in (3, 4, 5): cc.number_format = "#,##0"
        if c == 6: cc.number_format = '0.00"%"'
    for i, w in enumerate([7, 18, 20, 22, 16, 16], start=1):
        ws1.column_dimensions[get_column_letter(i)].width = w
    ws1.freeze_panes = "A2"

    # === Patwari sheets ===
    for title, rows in [("PATWARI BY %", views["patwari_by_pct"]),
                        ("PATWARI BY COUNT", views["patwari_by_count"])]:
        ws = wb.create_sheet(title)
        write_hdr(ws, ["S.NO", "NAME OF PATWARI", "VILLAGES",
                       "TOTAL SURVEY NOS", "SUBMITTED", "% COMPLETION"])
        for i, row in enumerate(rows, start=1):
            r = i + 1
            vals = [i, row["patwari"], row["villages_list"],
                    row["total"], row["submitted"], round(row["pct"], 2)]
            aligns = [center, left, left, center, center, center]
            for ci, (v, a) in enumerate(zip(vals, aligns), start=1):
                c = ws.cell(row=r, column=ci, value=v)
                c.alignment = a; c.font = body_font; c.border = border
                if ci in (4, 5): c.number_format = "#,##0"
                if ci == 6: c.number_format = '0.00"%"'
        pt = len(rows) + 2
        ws.cell(row=pt, column=2, value="TOTAL").alignment = left
        ws.cell(row=pt, column=4, value=grand["total_khasras"]).alignment = center
        ws.cell(row=pt, column=5, value=grand["submitted"]).alignment = center
        ws.cell(row=pt, column=6, value=round(grand["overall_pct"], 2)).alignment = center
        for c in range(1, 7):
            cc = ws.cell(row=pt, column=c)
            cc.font = tot_font; cc.fill = tot_fill; cc.border = border
            if c in (4, 5): cc.number_format = "#,##0"
            if c == 6: cc.number_format = '0.00"%"'
        for i, w in enumerate([7, 28, 55, 22, 16, 16], start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"

    # === NOT STARTED ===
    ws3 = wb.create_sheet("NOT STARTED")
    write_hdr(ws3, ["S.NO", "VILLAGE", "TEHSIL", "NAME OF PATWARI"], fill=ns_fill)
    for i, row in enumerate(views["not_started"], start=1):
        r = i + 1
        vals = [i, row["village"], row["tehsil"], row["patwari"]]
        aligns = [center, left, left, left]
        for ci, (v, a) in enumerate(zip(vals, aligns), start=1):
            c = ws3.cell(row=r, column=ci, value=v)
            c.alignment = a; c.font = body_font; c.border = border
    for i, w in enumerate([7, 28, 18, 28], start=1):
        ws3.column_dimensions[get_column_letter(i)].width = w
    ws3.freeze_panes = "A2"

    # === META sheet ===
    ws4 = wb.create_sheet("META")
    ws4.cell(row=1, column=1, value="Data as of").font = tot_font
    ws4.cell(row=1, column=2, value=snap_date.strftime("%d %b %Y"))
    ws4.column_dimensions["A"].width = 18
    ws4.column_dimensions["B"].width = 20

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


@app.route("/healthz")
def healthz():
    return "ok", 200


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
