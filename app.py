"""
AGRISTACK Dashboard — District Bandipora

Reads snapshots from snapshots/YYYY-MM-DD.xlsx and shows:
  - Tehsil-wise: total, daily target, submitted, % completion, additions since a chosen date
  - Patwari-wise (two sorts): additions since a chosen date
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
DEFAULT_DAILY_TARGET = 19302

COL_MATCHERS = {
    "tehsil":    re.compile(r"tehsil(?!dar)", re.I),   # must not match TEHSILDAR
    "village":   re.compile(r"^village$", re.I),
    "patwari":   re.compile(r"patwari", re.I),
    "khasras":   re.compile(r"khasra|survey", re.I),
    "submitted": re.compile(r"^submitted$", re.I),
}

# Optional — dashboard degrades gracefully if these columns aren't present yet
OPTIONAL_MATCHERS = {
    "checker":     re.compile(r"maker|checker", re.I),   # CONCERNED MAKER acts as the checker
    "subdivision": re.compile(r"sub.?div", re.I),
}


# ---------- Snapshot discovery ----------

def all_snapshots():
    """Return [(date, path), ...] sorted newest first."""
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
        mtime = os.path.getmtime(LEGACY_EXCEL)
        out.append((datetime.fromtimestamp(mtime).date(), LEGACY_EXCEL))
    out.sort(key=lambda x: x[0], reverse=True)
    return out


def snapshot_for_date(snapshots, target_date):
    """Return path of the snapshot whose date matches target_date, or None."""
    for d, p in snapshots:
        if d == target_date:
            return p
    return None


# ---------- Config ----------

def read_daily_target():
    if os.path.exists("DAILY_TARGET.txt"):
        try:
            with open("DAILY_TARGET.txt", "r", encoding="utf-8") as f:
                v = int(f.read().strip())
                return v if v > 0 else DEFAULT_DAILY_TARGET
        except Exception:
            pass
    return DEFAULT_DAILY_TARGET


# ---------- Excel loading ----------

_df_cache = {}


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
    # Optional columns — dashboard shows extra tabs only when present
    for logical, rx in OPTIONAL_MATCHERS.items():
        for c in df.columns:
            key = str(c).strip()
            if rx.search(key):
                cols[logical] = c
                break
    return cols


def load_snapshot(path):
    if path in _df_cache:
        return _df_cache[path]
    xl = pd.ExcelFile(path)
    sheet = _pick_sheet(xl)
    df = pd.read_excel(xl, sheet_name=sheet)
    cols = _detect_columns(df)

    rename_map = {
        cols["tehsil"]: "tehsil",
        cols["village"]: "village",
        cols["patwari"]: "patwari",
        cols["khasras"]: "khasras",
        cols["submitted"]: "submitted",
    }
    select_cols = ["tehsil", "village", "patwari", "khasras", "submitted"]
    if "checker" in cols:
        rename_map[cols["checker"]] = "checker"
        select_cols.append("checker")
    if "subdivision" in cols:
        rename_map[cols["subdivision"]] = "subdivision"
        select_cols.append("subdivision")

    df = df.rename(columns=rename_map)[select_cols].copy()
    df["khasras"] = pd.to_numeric(df["khasras"], errors="coerce").fillna(0).astype(int)
    df["submitted"] = pd.to_numeric(df["submitted"], errors="coerce").fillna(0).astype(int)
    for col in ["tehsil", "village", "patwari"] + \
               (["checker"] if "checker" in select_cols else []) + \
               (["subdivision"] if "subdivision" in select_cols else []):
        df[col] = df[col].astype(str).str.strip()
        df[col] = df[col].replace({"nan": "", "NaN": "", "None": ""})
    df = df[(df["village"] != "") & (df["village"].str.lower() != "nan")]
    _df_cache[path] = df
    return df


# ---------- View building ----------

def _pct(sub, tot):
    return (sub / tot * 100) if tot > 0 else 0.0


def build_views(current_df, from_df=None, daily_target=DEFAULT_DAILY_TARGET):
    """Build all dashboard views.
    from_df: optional older snapshot for computing 'Additions'.
    daily_target: district-level daily target for computing per-tehsil targets.
    """
    has_additions = from_df is not None
    district_total = int(current_df["khasras"].sum())

    # Per-tehsil daily target = proportional share of district daily target
    tehsil_totals = current_df.groupby("tehsil")["khasras"].sum().to_dict()
    tehsil_daily_target = {
        t: round((int(total) / district_total) * daily_target) if district_total > 0 else 0
        for t, total in tehsil_totals.items()
    }

    # Old-snapshot lookups for additions
    old_by_tehsil = from_df.groupby("tehsil")["submitted"].sum().to_dict() if has_additions else {}
    old_by_patwari = from_df.groupby("patwari")["submitted"].sum().to_dict() if has_additions else {}

    # --- Tehsil-wise ---
    t = (current_df.groupby("tehsil", as_index=False)
                   .agg(villages=("village", "count"),
                        total=("khasras", "sum"),
                        submitted=("submitted", "sum"))
                   .sort_values("submitted", ascending=False))
    tehsil_rows = []
    for _, row in t.iterrows():
        additions = None
        if has_additions:
            additions = int(row["submitted"]) - int(old_by_tehsil.get(row["tehsil"], 0))
        tehsil_rows.append({
            "tehsil": row["tehsil"],
            "villages": int(row["villages"]),
            "total": int(row["total"]),
            "daily_target": tehsil_daily_target.get(row["tehsil"], 0),
            "submitted": int(row["submitted"]),
            "pct": _pct(int(row["submitted"]), int(row["total"])),
            "additions": additions,
        })

    # --- Patwari-wise ---
    p = (current_df.groupby("patwari", as_index=False)
                   .agg(villages_list=("village", lambda s: ", ".join(sorted(s.tolist()))),
                        total=("khasras", "sum"),
                        submitted=("submitted", "sum")))
    p["pct"] = p.apply(lambda r: _pct(int(r["submitted"]), int(r["total"])), axis=1)
    if has_additions:
        p["additions"] = p["patwari"].map(lambda n: int(old_by_patwari.get(n, 0)))
        p["additions"] = p["submitted"] - p["additions"]
    else:
        p["additions"] = None

    def _records(frame):
        return [{
            "patwari": r["patwari"],
            "villages_list": r["villages_list"],
            "total": int(r["total"]),
            "submitted": int(r["submitted"]),
            "pct": float(r["pct"]),
            "additions": None if r["additions"] is None else int(r["additions"]),
        } for _, r in frame.iterrows()]

    patwari_by_pct = _records(p.sort_values(["pct", "submitted"], ascending=[False, False]))
    patwari_by_count = _records(p.sort_values(["submitted", "pct"], ascending=[False, False]))

    # --- Not started ---
    ns = current_df[current_df["submitted"] == 0].sort_values(["tehsil", "village"])
    not_started = ns.to_dict("records")

    # --- Checker wise (optional) ---
    checker_rows = None
    if "checker" in current_df.columns and current_df["checker"].str.strip().replace("", pd.NA).notna().any():
        c_df = current_df[current_df["checker"].str.strip() != ""].copy()
        agg_dict = {
            "villages": ("village", "count"),
            "villages_list": ("village", lambda s: ", ".join(sorted(s.tolist()))),
            "total": ("khasras", "sum"),
            "submitted": ("submitted", "sum"),
        }
        if "subdivision" in c_df.columns:
            agg_dict["subdivisions"] = ("subdivision", lambda s: " / ".join(sorted(set(x for x in s if x))))
        c = c_df.groupby("checker", as_index=False).agg(**agg_dict).sort_values("submitted", ascending=False)
        checker_rows = []
        for _, row in c.iterrows():
            additions = None
            if has_additions:
                old_val = int(from_df[from_df["checker"] == row["checker"]]["submitted"].sum()) if "checker" in from_df.columns else 0
                additions = int(row["submitted"]) - old_val
            daily_tgt = round((int(row["total"]) / district_total) * daily_target) if district_total > 0 else 0
            checker_rows.append({
                "name": row["checker"],
                "subdivision": row.get("subdivisions", "") if "subdivision" in c_df.columns else "",
                "villages_list": row["villages_list"],
                "villages": int(row["villages"]),
                "total": int(row["total"]),
                "daily_target": daily_tgt,
                "submitted": int(row["submitted"]),
                "pct": _pct(int(row["submitted"]), int(row["total"])),
                "additions": additions,
            })

    # --- Sub-Division wise (optional) ---
    subdiv_rows = None
    if "subdivision" in current_df.columns and current_df["subdivision"].str.strip().replace("", pd.NA).notna().any():
        s_df = current_df[current_df["subdivision"].str.strip() != ""].copy()
        s = (s_df.groupby("subdivision", as_index=False)
                 .agg(villages=("village", "count"),
                      total=("khasras", "sum"),
                      submitted=("submitted", "sum"))
                 .sort_values("submitted", ascending=False))
        subdiv_rows = []
        for _, row in s.iterrows():
            additions = None
            if has_additions:
                old_val = int(from_df[from_df["subdivision"] == row["subdivision"]]["submitted"].sum()) if "subdivision" in from_df.columns else 0
                additions = int(row["submitted"]) - old_val
            daily_tgt = round((int(row["total"]) / district_total) * daily_target) if district_total > 0 else 0
            subdiv_rows.append({
                "name": row["subdivision"],
                "villages": int(row["villages"]),
                "total": int(row["total"]),
                "daily_target": daily_tgt,
                "submitted": int(row["submitted"]),
                "pct": _pct(int(row["submitted"]), int(row["total"])),
                "additions": additions,
            })

    top_patwari = patwari_by_count[0] if patwari_by_count else None
    top_tehsil = tehsil_rows[0] if tehsil_rows else None

    # Totals for the Checker Wise and Sub-Division Wise footer rows.
    # Computed from the rows in the view (not district-wide) so they honestly
    # reflect coverage — if some villages have blank checker/subdivision they're
    # excluded from these sums.
    def _totals(rows):
        if not rows:
            return None
        subs = sum(r["submitted"] for r in rows)
        tot = sum(r["total"] for r in rows)
        return {
            "villages": sum(r["villages"] for r in rows),
            "total": tot,
            "daily_target": sum(r["daily_target"] for r in rows),
            "submitted": subs,
            "pct": _pct(subs, tot),
            "additions": sum(r["additions"] for r in rows) if has_additions else None,
        }
    checker_totals = _totals(checker_rows)
    subdiv_totals = _totals(subdiv_rows)

    grand = {
        "tehsils": int(len(t)),
        "villages": int(len(current_df)),
        "patwaris": int(len(p)),
        "total_khasras": int(current_df["khasras"].sum()),
        "daily_target": sum(tehsil_daily_target.values()),
        "submitted": int(current_df["submitted"].sum()),
        "not_started": int(len(ns)),
        "overall_pct": _pct(int(current_df["submitted"].sum()), int(current_df["khasras"].sum())),
    }
    if has_additions:
        grand["additions"] = int(current_df["submitted"].sum()) - int(from_df["submitted"].sum())
    else:
        grand["additions"] = None

    return {
        "tehsil_rows": tehsil_rows,
        "patwari_by_pct": patwari_by_pct,
        "patwari_by_count": patwari_by_count,
        "checker_rows": checker_rows,
        "checker_totals": checker_totals,
        "subdiv_rows": subdiv_rows,
        "subdiv_totals": subdiv_totals,
        "not_started": not_started,
        "top_patwari": top_patwari,
        "top_tehsil": top_tehsil,
        "grand": grand,
        "has_additions": has_additions,
    }


# ---------- Helpers ----------

def read_as_of_override():
    if os.path.exists("AS_OF.txt"):
        with open("AS_OF.txt", "r", encoding="utf-8") as f:
            text = f.read().strip()
        if text:
            return text
    return None


def format_date(d):
    return d.strftime("%d %b %Y")


def format_date_short(d):
    return d.strftime("%d %b")


def parse_date_param(s):
    try:
        return datetime.strptime(s, "%Y-%m-%d").date()
    except Exception:
        return None


def resolve_dates(snapshots, from_param, to_param):
    """Given query-string dates and available snapshots, return (from_date, to_date) or (None, to_date).
    to_date defaults to newest snapshot.
    from_date defaults to second-newest (if any).
    """
    if not snapshots:
        return None, None
    dates_available = [d for d, _ in snapshots]
    to_date = parse_date_param(to_param) if to_param else None
    if not to_date or to_date not in dates_available:
        to_date = dates_available[0]  # newest
    from_date = parse_date_param(from_param) if from_param else None
    if not from_date or from_date not in dates_available or from_date >= to_date:
        # Default to next-oldest snapshot before to_date
        older = [d for d in dates_available if d < to_date]
        from_date = older[0] if older else None
    return from_date, to_date


# ---------- Routes ----------

@app.route("/")
def index():
    snapshots = all_snapshots()
    if not snapshots:
        return "No snapshots found. Add a file to snapshots/ folder.", 500

    from_date, to_date = resolve_dates(snapshots,
                                        request.args.get("from"),
                                        request.args.get("to"))

    to_path = snapshot_for_date(snapshots, to_date)
    current_df = load_snapshot(to_path)

    from_df = None
    if from_date:
        from_path = snapshot_for_date(snapshots, from_date)
        if from_path:
            from_df = load_snapshot(from_path)

    daily_target = read_daily_target()
    views = build_views(current_df, from_df, daily_target)

    override = read_as_of_override()
    as_of_display = override if override else format_date(to_date)

    # For the dropdowns
    date_options = [(d.isoformat(), format_date(d)) for d, _ in snapshots]

    additions_label = None
    if views["has_additions"]:
        additions_label = f"Additions ({format_date_short(from_date)} → {format_date_short(to_date)})"

    return render_template(
        "index.html",
        as_of=as_of_display,
        date_options=date_options,
        from_date_iso=from_date.isoformat() if from_date else None,
        to_date_iso=to_date.isoformat(),
        additions_label=additions_label,
        snapshot_count=len(snapshots),
        **views,
    )


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
    snapshots = all_snapshots()
    if not snapshots:
        return "No snapshots found.", 500

    from_date, to_date = resolve_dates(snapshots,
                                        request.args.get("from"),
                                        request.args.get("to"))
    to_path = snapshot_for_date(snapshots, to_date)
    current_df = load_snapshot(to_path)
    from_df = load_snapshot(snapshot_for_date(snapshots, from_date)) if from_date else None
    daily_target = read_daily_target()
    views = build_views(current_df, from_df, daily_target)

    buf = _generate_workbook(views, to_date, from_date)
    filename = f"AGRISTACK_Dashboard_{to_date.strftime('%Y-%m-%d')}.xlsx"
    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename,
    )


def _generate_workbook(views, to_date, from_date):
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

    additions_hdr = None
    if views["has_additions"]:
        additions_hdr = f"Additions ({from_date.strftime('%d %b')} to {to_date.strftime('%d %b')})"

    # === TEHSIL WISE ===
    ws1 = wb.active
    ws1.title = "TEHSIL WISE"
    tehsil_headers = ["S.NO", "TEHSIL", "NUMBER OF VILLAGES", "TOTAL SURVEY NOS",
                      "DAILY TARGET", "SUBMITTED", "% COMPLETION"]
    if additions_hdr:
        tehsil_headers.append(additions_hdr)
    write_hdr(ws1, tehsil_headers)

    for i, row in enumerate(views["tehsil_rows"], start=1):
        r = i + 1
        vals = [i, row["tehsil"], row["villages"], row["total"],
                row["daily_target"], row["submitted"], round(row["pct"], 2)]
        aligns = [center, left, center, center, center, center, center]
        if additions_hdr:
            vals.append(row["additions"] if row["additions"] is not None else "—")
            aligns.append(center)
        for ci, (v, a) in enumerate(zip(vals, aligns), start=1):
            c = ws1.cell(row=r, column=ci, value=v)
            c.alignment = a; c.font = body_font; c.border = border
            if ci in (3, 4, 5, 6): c.number_format = "#,##0"
            if ci == 7: c.number_format = '0.00"%"'
            if additions_hdr and ci == 8 and isinstance(v, int):
                c.number_format = "+#,##0;-#,##0;0"

    tt = len(views["tehsil_rows"]) + 2
    ws1.cell(row=tt, column=2, value="TOTAL").alignment = left
    ws1.cell(row=tt, column=3, value=grand["villages"]).alignment = center
    ws1.cell(row=tt, column=4, value=grand["total_khasras"]).alignment = center
    ws1.cell(row=tt, column=5, value=grand["daily_target"]).alignment = center
    ws1.cell(row=tt, column=6, value=grand["submitted"]).alignment = center
    ws1.cell(row=tt, column=7, value=round(grand["overall_pct"], 2)).alignment = center
    if additions_hdr:
        ws1.cell(row=tt, column=8, value=grand["additions"] if grand["additions"] is not None else "—").alignment = center

    tot_cols = 8 if additions_hdr else 7
    for c in range(1, tot_cols + 1):
        cc = ws1.cell(row=tt, column=c)
        cc.font = tot_font; cc.fill = tot_fill; cc.border = border
        if c in (3, 4, 5, 6): cc.number_format = "#,##0"
        if c == 7: cc.number_format = '0.00"%"'
        if additions_hdr and c == 8 and isinstance(cc.value, int):
            cc.number_format = "+#,##0;-#,##0;0"
    widths = [7, 16, 16, 18, 14, 14, 14]
    if additions_hdr: widths.append(22)
    for i, w in enumerate(widths, start=1):
        ws1.column_dimensions[get_column_letter(i)].width = w
    ws1.freeze_panes = "A2"

    # === Patwari sheets ===
    for title, rows in [("PATWARI BY %", views["patwari_by_pct"]),
                        ("PATWARI BY COUNT", views["patwari_by_count"])]:
        ws = wb.create_sheet(title)
        p_headers = ["S.NO", "NAME OF PATWARI", "VILLAGES", "TOTAL SURVEY NOS",
                     "SUBMITTED", "% COMPLETION"]
        if additions_hdr:
            p_headers.append(additions_hdr)
        write_hdr(ws, p_headers)

        for i, row in enumerate(rows, start=1):
            r = i + 1
            vals = [i, row["patwari"], row["villages_list"], row["total"],
                    row["submitted"], round(row["pct"], 2)]
            aligns = [center, left, left, center, center, center]
            if additions_hdr:
                vals.append(row["additions"] if row["additions"] is not None else "—")
                aligns.append(center)
            for ci, (v, a) in enumerate(zip(vals, aligns), start=1):
                c = ws.cell(row=r, column=ci, value=v)
                c.alignment = a; c.font = body_font; c.border = border
                if ci in (4, 5): c.number_format = "#,##0"
                if ci == 6: c.number_format = '0.00"%"'
                if additions_hdr and ci == 7 and isinstance(v, int):
                    c.number_format = "+#,##0;-#,##0;0"
        pt = len(rows) + 2
        ws.cell(row=pt, column=2, value="TOTAL").alignment = left
        ws.cell(row=pt, column=4, value=grand["total_khasras"]).alignment = center
        ws.cell(row=pt, column=5, value=grand["submitted"]).alignment = center
        ws.cell(row=pt, column=6, value=round(grand["overall_pct"], 2)).alignment = center
        if additions_hdr:
            ws.cell(row=pt, column=7, value=grand["additions"] if grand["additions"] is not None else "—").alignment = center
        tot_cols = 7 if additions_hdr else 6
        for c in range(1, tot_cols + 1):
            cc = ws.cell(row=pt, column=c)
            cc.font = tot_font; cc.fill = tot_fill; cc.border = border
            if c in (4, 5): cc.number_format = "#,##0"
            if c == 6: cc.number_format = '0.00"%"'
            if additions_hdr and c == 7 and isinstance(cc.value, int):
                cc.number_format = "+#,##0;-#,##0;0"
        widths = [7, 26, 50, 18, 14, 14]
        if additions_hdr: widths.append(22)
        for i, w in enumerate(widths, start=1):
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

    # === NAIB TEHSILDAR WISE (only if data exists) ===
    def _write_group_sheet(title, rows, name_header, totals):
        ws = wb.create_sheet(title)
        headers = ["S.NO", name_header, "NUMBER OF VILLAGES", "TOTAL SURVEY NOS",
                   "DAILY TARGET", "SUBMITTED", "% COMPLETION"]
        if additions_hdr:
            headers.append(additions_hdr)
        write_hdr(ws, headers)
        for i, row in enumerate(rows, start=1):
            r = i + 1
            vals = [i, row["name"], row["villages"], row["total"],
                    row["daily_target"], row["submitted"], round(row["pct"], 2)]
            aligns = [center, left, center, center, center, center, center]
            if additions_hdr:
                vals.append(row["additions"] if row["additions"] is not None else "—")
                aligns.append(center)
            for ci, (v, a) in enumerate(zip(vals, aligns), start=1):
                c = ws.cell(row=r, column=ci, value=v)
                c.alignment = a; c.font = body_font; c.border = border
                if ci in (3, 4, 5, 6): c.number_format = "#,##0"
                if ci == 7: c.number_format = '0.00"%"'
                if additions_hdr and ci == 8 and isinstance(v, int):
                    c.number_format = "+#,##0;-#,##0;0"
        # TOTAL row
        if totals:
            pt = len(rows) + 2
            ws.cell(row=pt, column=2, value="TOTAL").alignment = left
            ws.cell(row=pt, column=3, value=totals["villages"]).alignment = center
            ws.cell(row=pt, column=4, value=totals["total"]).alignment = center
            ws.cell(row=pt, column=5, value=totals["daily_target"]).alignment = center
            ws.cell(row=pt, column=6, value=totals["submitted"]).alignment = center
            ws.cell(row=pt, column=7, value=round(totals["pct"], 2)).alignment = center
            if additions_hdr:
                ws.cell(row=pt, column=8, value=totals["additions"] if totals["additions"] is not None else "—").alignment = center
            tot_col_max = 8 if additions_hdr else 7
            for c in range(1, tot_col_max + 1):
                cc = ws.cell(row=pt, column=c)
                cc.font = tot_font; cc.fill = tot_fill; cc.border = border
                if c in (3, 4, 5, 6): cc.number_format = "#,##0"
                if c == 7: cc.number_format = '0.00"%"'
                if additions_hdr and c == 8 and isinstance(cc.value, int):
                    cc.number_format = "+#,##0;-#,##0;0"
        widths = [7, 24, 18, 18, 14, 14, 14]
        if additions_hdr: widths.append(22)
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"

    if views.get("checker_rows"):
        ws = wb.create_sheet("CHECKER WISE")
        headers = ["S.NO", "CHECKER", "SUB-DIVISION", "VILLAGES",
                   "TOTAL SURVEY NOS", "DAILY TARGET", "SUBMITTED", "% COMPLETION"]
        if additions_hdr:
            headers.append(additions_hdr)
        write_hdr(ws, headers)
        for i, row in enumerate(views["checker_rows"], start=1):
            r = i + 1
            vals = [i, row["name"], row["subdivision"], row["villages_list"],
                    row["total"], row["daily_target"], row["submitted"],
                    round(row["pct"], 2)]
            aligns = [center, left, left, left, center, center, center, center]
            if additions_hdr:
                vals.append(row["additions"] if row["additions"] is not None else "—")
                aligns.append(center)
            for ci, (v, a) in enumerate(zip(vals, aligns), start=1):
                c = ws.cell(row=r, column=ci, value=v)
                c.alignment = a; c.font = body_font; c.border = border
                if ci in (5, 6, 7): c.number_format = "#,##0"
                if ci == 8: c.number_format = '0.00"%"'
                if additions_hdr and ci == 9 and isinstance(v, int):
                    c.number_format = "+#,##0;-#,##0;0"
        widths = [7, 32, 14, 46, 16, 14, 14, 14]
        if additions_hdr: widths.append(22)
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        # TOTAL row
        ct = views.get("checker_totals")
        if ct:
            pt = len(views["checker_rows"]) + 2
            ws.cell(row=pt, column=2, value="TOTAL").alignment = left
            ws.cell(row=pt, column=5, value=ct["total"]).alignment = center
            ws.cell(row=pt, column=6, value=ct["daily_target"]).alignment = center
            ws.cell(row=pt, column=7, value=ct["submitted"]).alignment = center
            ws.cell(row=pt, column=8, value=round(ct["pct"], 2)).alignment = center
            if additions_hdr:
                ws.cell(row=pt, column=9, value=ct["additions"] if ct["additions"] is not None else "—").alignment = center
            tot_col_max = 9 if additions_hdr else 8
            for c in range(1, tot_col_max + 1):
                cc = ws.cell(row=pt, column=c)
                cc.font = tot_font; cc.fill = tot_fill; cc.border = border
                if c in (5, 6, 7): cc.number_format = "#,##0"
                if c == 8: cc.number_format = '0.00"%"'
                if additions_hdr and c == 9 and isinstance(cc.value, int):
                    cc.number_format = "+#,##0;-#,##0;0"
        ws.freeze_panes = "A2"
    if views.get("subdiv_rows"):
        _write_group_sheet("SUB-DIVISION WISE", views["subdiv_rows"], "SUB-DIVISION", views.get("subdiv_totals"))

    # === META ===
    ws4 = wb.create_sheet("META")
    ws4.cell(row=1, column=1, value="Data as of").font = tot_font
    ws4.cell(row=1, column=2, value=to_date.strftime("%d %b %Y"))
    ws4.cell(row=2, column=1, value="Additions from").font = tot_font
    ws4.cell(row=2, column=2, value=from_date.strftime("%d %b %Y") if from_date else "—")
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
