"""
AGRISTACK Dashboard — District Bandipora
Reads dated snapshots from snapshots/YYYY-MM-DD.xlsx and shows:
  - Tehsil-wise progress with % completion and rolling change
  - Patwari-wise progress (three sorts: by %, by count, by recent activity)
  - Villages not started
Also serves a password-gated Excel download of the current dashboard state.
"""
import os
import re
import io
from datetime import datetime, date, timedelta

import pandas as pd
from flask import Flask, render_template, request, Response, send_file
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)

SNAPSHOTS_DIR = "snapshots"
LEGACY_EXCEL = "AGRISTACK.xlsx"
HISTORY_WINDOW_DAYS = 7

# Tolerant column matchers
COL_MATCHERS = {
    "tehsil":    re.compile(r"tehsil", re.I),
    "village":   re.compile(r"^village$", re.I),
    "patwari":   re.compile(r"patwari", re.I),
    "khasras":   re.compile(r"khasra|survey", re.I),
    "submitted": re.compile(r"^submitted$", re.I),
}


# ---------- Snapshot discovery ----------

def list_snapshots():
    """Return [(date, path), ...] sorted newest first.

    Prefers files in snapshots/ named YYYY-MM-DD.xlsx.
    Falls back to legacy AGRISTACK.xlsx if snapshots dir is empty.
    """
    out = []
    if os.path.isdir(SNAPSHOTS_DIR):
        for name in os.listdir(SNAPSHOTS_DIR):
            if not name.endswith(".xlsx") or name.startswith("~"):
                continue
            stem = name[:-5]
            try:
                d = datetime.strptime(stem, "%Y-%m-%d").date()
            except ValueError:
                continue
            out.append((d, os.path.join(SNAPSHOTS_DIR, name)))
    if not out and os.path.exists(LEGACY_EXCEL):
        # Legacy fallback — treat AGRISTACK.xlsx as today's snapshot
        mtime = os.path.getmtime(LEGACY_EXCEL)
        out.append((datetime.fromtimestamp(mtime).date(), LEGACY_EXCEL))
    out.sort(key=lambda x: x[0], reverse=True)
    return out


def choose_history_pair(snapshots):
    """Given snapshots newest-first, return (newest, oldest_in_window, gap_days).

    If only one snapshot exists, oldest == newest and gap == 0.
    """
    if not snapshots:
        return None
    newest_date, newest_path = snapshots[0]
    cutoff = newest_date - timedelta(days=HISTORY_WINDOW_DAYS)
    within = [(d, p) for d, p in snapshots if d >= cutoff]
    within.sort(key=lambda x: x[0])
    oldest_date, oldest_path = within[0]
    gap = (newest_date - oldest_date).days
    return newest_date, newest_path, oldest_date, oldest_path, gap


# ---------- Excel loading ----------

_df_cache = {}  # path -> DataFrame (paths are date-stamped so cache is safe)


def _pick_sheet(xl):
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
    return scored[0][2]


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
    return df


# ---------- View building ----------

def _pct(sub, tot):
    return (sub / tot * 100) if tot > 0 else 0.0


def build_views(current_df, old_df=None, gap_days=0):
    """Aggregate current + (optional) old snapshot into all dashboard views.

    old_df: previous snapshot to compute deltas against. None or same as current
            (day-one case) means all Change values are None (rendered as "—").
    gap_days: calendar days between current and old snapshot.
    """
    has_history = old_df is not None and gap_days > 0

    # --- Old-snapshot lookups (for deltas) ---
    old_by_tehsil = {}
    old_by_patwari = {}
    if has_history:
        ot = old_df.groupby("tehsil")["submitted"].sum()
        old_by_tehsil = ot.to_dict()
        op = old_df.groupby("patwari")["submitted"].sum()
        old_by_patwari = op.to_dict()

    # --- Tehsil-wise ---
    t = (current_df.groupby("tehsil", as_index=False)
                   .agg(villages=("village", "count"),
                        total=("khasras", "sum"),
                        submitted=("submitted", "sum"))
                   .sort_values("submitted", ascending=False))
    tehsil_rows = []
    for _, row in t.iterrows():
        change = None
        if has_history:
            prev = int(old_by_tehsil.get(row["tehsil"], 0))
            change = int(row["submitted"]) - prev
        tehsil_rows.append({
            "tehsil": row["tehsil"],
            "villages": int(row["villages"]),
            "total": int(row["total"]),
            "submitted": int(row["submitted"]),
            "pct": _pct(int(row["submitted"]), int(row["total"])),
            "change": change,
        })

    # --- Patwari-wise (three sorts) ---
    p = (current_df.groupby("patwari", as_index=False)
                   .agg(villages_list=("village", lambda s: ", ".join(sorted(s.tolist()))),
                        total=("khasras", "sum"),
                        submitted=("submitted", "sum")))
    p["pct"] = p.apply(lambda r: _pct(int(r["submitted"]), int(r["total"])), axis=1)
    if has_history:
        p["change"] = p["patwari"].map(lambda n: int(p.loc[p["patwari"] == n, "submitted"].iloc[0]) - int(old_by_patwari.get(n, 0)))
    else:
        p["change"] = None

    def _records(frame):
        out = []
        for _, row in frame.iterrows():
            out.append({
                "patwari": row["patwari"],
                "villages_list": row["villages_list"],
                "total": int(row["total"]),
                "submitted": int(row["submitted"]),
                "pct": float(row["pct"]),
                "change": None if row["change"] is None else int(row["change"]),
            })
        return out

    patwari_by_pct = _records(p.sort_values(["pct", "submitted"], ascending=[False, False]))
    patwari_by_count = _records(p.sort_values(["submitted", "pct"], ascending=[False, False]))
    patwari_by_recent = None
    if has_history:
        # Sort by change desc, then submitted desc as tie-breaker
        patwari_by_recent = _records(p.sort_values(["change", "submitted"], ascending=[False, False]))

    # --- Not started ---
    ns = current_df[current_df["submitted"] == 0].sort_values(["tehsil", "village"])
    not_started = ns.to_dict("records")

    # --- Top performer cards (by absolute submitted count) ---
    top_patwari = patwari_by_count[0] if patwari_by_count else None
    top_tehsil = None
    if tehsil_rows:
        # Tehsil rows are already sorted by submitted desc
        top_tehsil = tehsil_rows[0]

    grand = {
        "tehsils": int(len(t)),
        "villages": int(len(current_df)),
        "patwaris": int(len(p)),
        "total_khasras": int(current_df["khasras"].sum()),
        "submitted": int(current_df["submitted"].sum()),
        "not_started": int(len(ns)),
        "overall_pct": _pct(int(current_df["submitted"].sum()), int(current_df["khasras"].sum())),
    }
    if has_history:
        grand["change"] = int(current_df["submitted"].sum()) - int(old_df["submitted"].sum())
    else:
        grand["change"] = None

    return {
        "tehsil_rows": tehsil_rows,
        "patwari_by_pct": patwari_by_pct,
        "patwari_by_count": patwari_by_count,
        "patwari_by_recent": patwari_by_recent,
        "not_started": not_started,
        "top_patwari": top_patwari,
        "top_tehsil": top_tehsil,
        "grand": grand,
        "has_history": has_history,
        "gap_days": gap_days,
    }


# ---------- Header meta ----------

def read_as_of_override():
    """If AS_OF.txt exists and has content, return it. Otherwise None."""
    if os.path.exists("AS_OF.txt"):
        with open("AS_OF.txt", "r", encoding="utf-8") as f:
            text = f.read().strip()
        if text:
            return text
    return None


def format_date(d):
    return d.strftime("%d %b %Y")


# ---------- Main route ----------

@app.route("/")
def index():
    snapshots = list_snapshots()
    if not snapshots:
        return "No snapshots found. Add a file to snapshots/ folder.", 500

    pair = choose_history_pair(snapshots)
    newest_date, newest_path, oldest_date, oldest_path, gap = pair

    current_df = load_snapshot(newest_path)
    old_df = load_snapshot(oldest_path) if gap > 0 else None
    views = build_views(current_df, old_df, gap)

    override = read_as_of_override()
    as_of_display = override if override else format_date(newest_date)
    compared_with = format_date(oldest_date) if gap > 0 else None

    return render_template(
        "index.html",
        as_of=as_of_display,
        compared_with=compared_with,
        **views,
    )


# ---------- Password-gated Excel download ----------

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

    snapshots = list_snapshots()
    if not snapshots:
        return "No snapshots found.", 500
    pair = choose_history_pair(snapshots)
    newest_date, newest_path, oldest_date, oldest_path, gap = pair
    current_df = load_snapshot(newest_path)
    old_df = load_snapshot(oldest_path) if gap > 0 else None
    views = build_views(current_df, old_df, gap)

    buf = _generate_workbook(views, newest_date, oldest_date, gap)
    filename = f"AGRISTACK_Dashboard_{newest_date.strftime('%Y-%m-%d')}.xlsx"
    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename,
    )


def _generate_workbook(views, newest_date, oldest_date, gap):
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

    change_label = f"Change (last {gap} days)" if gap > 0 else "Change"

    # === TEHSIL WISE ===
    ws1 = wb.active
    ws1.title = "TEHSIL WISE"
    write_hdr(ws1, ["S.NO", "TEHSIL", "NUMBER OF VILLAGES",
                    "TOTAL SURVEY NOS", "SUBMITTED", "% COMPLETION", change_label])
    for i, row in enumerate(views["tehsil_rows"], start=1):
        r = i + 1
        change_val = row["change"] if row["change"] is not None else "—"
        vals = [i, row["tehsil"], row["villages"], row["total"], row["submitted"],
                round(row["pct"], 2), change_val]
        aligns = [center, left, center, center, center, center, center]
        for ci, (v, a) in enumerate(zip(vals, aligns), start=1):
            c = ws1.cell(row=r, column=ci, value=v)
            c.alignment = a; c.font = body_font; c.border = border
            if ci in (3, 4, 5): c.number_format = "#,##0"
            if ci == 6: c.number_format = '0.00"%"'
            if ci == 7 and isinstance(v, int): c.number_format = "+#,##0;-#,##0;0"
    tt = len(views["tehsil_rows"]) + 2
    grand = views["grand"]
    ws1.cell(row=tt, column=2, value="TOTAL").alignment = left
    ws1.cell(row=tt, column=3, value=grand["villages"]).alignment = center
    ws1.cell(row=tt, column=4, value=grand["total_khasras"]).alignment = center
    ws1.cell(row=tt, column=5, value=grand["submitted"]).alignment = center
    ws1.cell(row=tt, column=6, value=round(grand["overall_pct"], 2)).alignment = center
    ws1.cell(row=tt, column=7, value=grand["change"] if grand["change"] is not None else "—").alignment = center
    for c in range(1, 8):
        cc = ws1.cell(row=tt, column=c)
        cc.font = tot_font; cc.fill = tot_fill; cc.border = border
        if c in (3, 4, 5): cc.number_format = "#,##0"
        if c == 6: cc.number_format = '0.00"%"'
        if c == 7 and isinstance(cc.value, int): cc.number_format = "+#,##0;-#,##0;0"
    for i, w in enumerate([7, 16, 20, 20, 14, 14, 20], start=1):
        ws1.column_dimensions[get_column_letter(i)].width = w
    ws1.freeze_panes = "A2"

    # === PATWARI sheets (by % and by count, plus by recent if history exists) ===
    patwari_variants = [
        ("PATWARI BY %", views["patwari_by_pct"]),
        ("PATWARI BY COUNT", views["patwari_by_count"]),
    ]
    if views["patwari_by_recent"]:
        patwari_variants.append(("PATWARI BY RECENT", views["patwari_by_recent"]))

    for title, rows in patwari_variants:
        ws = wb.create_sheet(title)
        write_hdr(ws, ["S.NO", "NAME OF PATWARI", "VILLAGES",
                       "TOTAL SURVEY NOS", "SUBMITTED", "% COMPLETION", change_label])
        for i, row in enumerate(rows, start=1):
            r = i + 1
            change_val = row["change"] if row["change"] is not None else "—"
            vals = [i, row["patwari"], row["villages_list"],
                    row["total"], row["submitted"],
                    round(row["pct"], 2), change_val]
            aligns = [center, left, left, center, center, center, center]
            for ci, (v, a) in enumerate(zip(vals, aligns), start=1):
                c = ws.cell(row=r, column=ci, value=v)
                c.alignment = a; c.font = body_font; c.border = border
                if ci in (4, 5): c.number_format = "#,##0"
                if ci == 6: c.number_format = '0.00"%"'
                if ci == 7 and isinstance(v, int): c.number_format = "+#,##0;-#,##0;0"
        pt = len(rows) + 2
        ws.cell(row=pt, column=2, value="TOTAL").alignment = left
        ws.cell(row=pt, column=4, value=grand["total_khasras"]).alignment = center
        ws.cell(row=pt, column=5, value=grand["submitted"]).alignment = center
        ws.cell(row=pt, column=6, value=round(grand["overall_pct"], 2)).alignment = center
        ws.cell(row=pt, column=7, value=grand["change"] if grand["change"] is not None else "—").alignment = center
        for c in range(1, 8):
            cc = ws.cell(row=pt, column=c)
            cc.font = tot_font; cc.fill = tot_fill; cc.border = border
            if c in (4, 5): cc.number_format = "#,##0"
            if c == 6: cc.number_format = '0.00"%"'
            if c == 7 and isinstance(cc.value, int): cc.number_format = "+#,##0;-#,##0;0"
        for i, w in enumerate([7, 26, 50, 20, 14, 14, 20], start=1):
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

    # === META sheet (dates so file is self-describing) ===
    ws4 = wb.create_sheet("META")
    ws4.cell(row=1, column=1, value="Data as of").font = tot_font
    ws4.cell(row=1, column=2, value=newest_date.strftime("%d %b %Y"))
    ws4.cell(row=2, column=1, value="Compared with").font = tot_font
    ws4.cell(row=2, column=2, value=oldest_date.strftime("%d %b %Y") if gap > 0 else "—")
    ws4.cell(row=3, column=1, value="Window (days)").font = tot_font
    ws4.cell(row=3, column=2, value=gap)
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
