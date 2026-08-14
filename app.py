"""
AGRISTACK Dashboard — District Bandipora

Tabs (left to right):
  1. Tehsil Wise      — S.No | Tehsil | Villages | Total Survey Nos | Submitted | % Completion | Additions | % Approved
  2. Overall All      — S.No | Tehsil | Total Khasras | Khasras Submitted | Khasras Verified | Khasras Approved
  3. Patwari Wise     — By % Completion (with Relative Effort) and By Total Submissions
  4. Checker Wise     — S.No | Checker | Sub-Division | Villages | Total Survey Nos | Submitted | Verified+Approved | % Completion | Additions
                       (% Completion here = (V + A + Seek Clarification) / Submitted)
  5. Village Wise     — S.No | Tehsil | Village | Total Survey Nos | Name of Patwari | Submitted | Verified | Approved | Additions

Sub-Division is derived from Tehsil via a hardcoded map (robust to broken SUB DIVISION source columns).
Tehsil filter checkboxes apply to Patwari and Village views: filter + high→low sort + color banding.
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
from xhtml2pdf import pisa

app = Flask(__name__)

SNAPSHOTS_DIR = "snapshots"
LEGACY_EXCEL = "AGRISTACK.xlsx"

# Tehsil -> Sub-Division (stable mapping; used because the SUB DIVISION column in source
# files has been unreliable — sometimes text, sometimes numbers). If any tehsil hierarchy
# ever changes, edit here.
TEHSIL_TO_SUBDIV = {
    "BANDIPORA": "BANDIPORA",
    "ALOOSA":    "BANDIPORA",
    "AJAS":      "BANDIPORA",
    "SUMBAL":    "SUMBAL",
    "HAJIN":     "SUMBAL",
    "GUREZ":     "GUREZ",
    "TULAIL":    "GUREZ",
}

COL_MATCHERS = {
    "tehsil":    re.compile(r"tehsil(?!dar)", re.I),   # not TEHSILDAR
    "village":   re.compile(r"^village$", re.I),
    "patwari":   re.compile(r"patwari", re.I),
    "khasras":   re.compile(r"khasra|survey", re.I),
    # 'submitted' is the patwari's total output. The new file introduces a "Total Submitted"
    # column which represents this (the file's "submitted" column now means "in-queue awaiting
    # first check" — a workflow stage, not the total). Prefer "Total Submitted" when present;
    # fall back to plain "submitted" for older snapshot files.
    "submitted": re.compile(r"^(total\s+)?submitted$", re.I),
}

# Optional workflow columns — added recently. Loaded when present, treated as 0 otherwise.
OPTIONAL_MATCHERS = {
    "checker":            re.compile(r"maker|checker", re.I),        # CONCERNED MAKER = checker
    "approved":           re.compile(r"^approved$", re.I),           # Approved
    "verified":           re.compile(r"^verified$", re.I),           # verified (excludes approved rows)
    "seek_clarification": re.compile(r"seek.*clar|clarif", re.I),    # Seek Clarification
}


# ---------- Snapshot discovery ----------

def all_snapshots():
    """Return [(date, path), ...] newest first."""
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
    for d, p in snapshots:
        if d == target_date:
            return p
    return None


# ---------- Excel loading ----------

_df_cache = {}


def _pick_sheet(xl):
    """Prefer Sheet2 (corrected), then MAIN, then first sheet scoring highest on required columns."""
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

    # Build the working DataFrame by extracting each column using its ORIGINAL name,
    # then assigning under its logical name. This avoids collisions when the file has
    # both "Total Submitted" (which we want as 'submitted') and the file's own
    # workflow-stage column also named "submitted" — pd.DataFrame.rename would leave
    # two columns with the same target name.
    new_df = pd.DataFrame()
    for logical in ("tehsil", "village", "patwari", "khasras", "submitted"):
        new_df[logical] = df[cols[logical]].values
    for opt in ("checker", "approved", "verified", "seek_clarification"):
        if opt in cols:
            new_df[opt] = df[cols[opt]].values

    # Numeric coercion
    for numcol in ("khasras", "submitted", "approved", "verified", "seek_clarification"):
        if numcol in new_df.columns:
            new_df[numcol] = pd.to_numeric(new_df[numcol], errors="coerce").fillna(0).astype(int)
        else:
            new_df[numcol] = 0
    # String cleanup
    for scol in ("tehsil", "village", "patwari", "checker"):
        if scol in new_df.columns:
            new_df[scol] = new_df[scol].astype(str).str.strip().replace({"nan": "", "NaN": "", "None": ""})
    new_df = new_df[(new_df["village"] != "") & (new_df["village"].str.lower() != "nan")]
    # Derive sub-division from tehsil
    new_df["subdivision"] = new_df["tehsil"].str.upper().map(TEHSIL_TO_SUBDIV).fillna("UNKNOWN")
    _df_cache[path] = new_df
    return new_df


# ---------- Helpers ----------

def _pct(num, denom):
    return (num / denom * 100) if denom > 0 else 0.0


def apply_bands(rows, sort_key):
    """Attach 'band' key to each row. Rows must already be sorted high-to-low."""
    n = len(rows)
    if n == 0:
        return
    top_n = round(n * 0.3)
    bot_n = round(n * 0.3)
    for i, r in enumerate(rows):
        if i < top_n:
            r["band"] = "green"
        elif i >= n - bot_n:
            r["band"] = "red"
        else:
            r["band"] = "yellow"


# ---------- View building ----------

def build_views(current_df, from_df=None, tehsils_filter=None):
    """Build every view. tehsils_filter: None | 'all' | set of tehsil names."""
    apply_coloring = tehsils_filter is not None
    active_tehsils = tehsils_filter if isinstance(tehsils_filter, set) else None
    has_additions = from_df is not None

    # === Tehsil Wise ===
    t = (current_df.groupby("tehsil", as_index=False)
                   .agg(villages=("village", "count"),
                        total=("khasras", "sum"),
                        submitted=("submitted", "sum"),
                        approved=("approved", "sum"))
                   .sort_values("submitted", ascending=False))
    old_by_tehsil = from_df.groupby("tehsil")["submitted"].sum().to_dict() if has_additions else {}
    tehsil_rows = []
    for _, row in t.iterrows():
        additions = int(row["submitted"]) - int(old_by_tehsil.get(row["tehsil"], 0)) if has_additions else None
        tehsil_rows.append({
            "tehsil": row["tehsil"],
            "villages": int(row["villages"]),
            "total": int(row["total"]),
            "submitted": int(row["submitted"]),
            "approved": int(row["approved"]),
            "pct": _pct(int(row["submitted"]), int(row["total"])),
            "pct_approved": _pct(int(row["approved"]), int(row["total"])),
            "additions": additions,
        })

    # === Overall All (new) ===
    # Verified column = verified + approved. Approved column = approved alone.
    o = (current_df.groupby("tehsil", as_index=False)
                   .agg(total=("khasras", "sum"),
                        submitted=("submitted", "sum"),
                        verified=("verified", "sum"),
                        approved=("approved", "sum"))
                   .sort_values("submitted", ascending=False))
    overall_rows = []
    for _, row in o.iterrows():
        overall_rows.append({
            "tehsil": row["tehsil"],
            "total": int(row["total"]),
            "submitted": int(row["submitted"]),
            "verified": int(row["verified"]) + int(row["approved"]),  # roll-in per spec
            "approved": int(row["approved"]),
        })
    overall_totals = {
        "total": sum(r["total"] for r in overall_rows),
        "submitted": sum(r["submitted"] for r in overall_rows),
        "verified": sum(r["verified"] for r in overall_rows),
        "approved": sum(r["approved"] for r in overall_rows),
    }

    # === Patwari Wise (two sorts) — filtered by tehsils picker ===
    if active_tehsils is not None:
        p_source = current_df[current_df["tehsil"].str.upper().isin(active_tehsils)]
        p_old_source = from_df[from_df["tehsil"].str.upper().isin(active_tehsils)] if has_additions else None
    else:
        p_source = current_df
        p_old_source = from_df if has_additions else None

    p = (p_source.groupby("patwari", as_index=False)
                 .agg(villages_list=("village", lambda s: ", ".join(sorted(s.tolist()))),
                      total=("khasras", "sum"),
                      submitted=("submitted", "sum")))
    p["pct"] = p.apply(lambda r: _pct(int(r["submitted"]), int(r["total"])), axis=1)
    if has_additions and p_old_source is not None:
        old_by_p = p_old_source.groupby("patwari")["submitted"].sum().to_dict()
        p["additions"] = p["patwari"].map(lambda n: int(old_by_p.get(n, 0)))
        p["additions"] = p["submitted"] - p["additions"]
    else:
        p["additions"] = None

    def _p_records(frame):
        return [{
            "patwari": r["patwari"],
            "villages_list": r["villages_list"],
            "total": int(r["total"]),
            "submitted": int(r["submitted"]),
            "pct": float(r["pct"]),
            "additions": None if r["additions"] is None else int(r["additions"]),
        } for _, r in frame.iterrows()]

    patwari_by_pct = _p_records(p.sort_values(["pct", "submitted"], ascending=[False, False]))
    patwari_by_count = _p_records(p.sort_values(["submitted", "pct"], ascending=[False, False]))

    # Relative Effort — district-wide top submission count, only for By %
    district_top = int(current_df.groupby("patwari")["submitted"].sum().max() or 0)
    for r in patwari_by_pct:
        r["relative_effort"] = (r["submitted"] / district_top * 100) if district_top > 0 else None

    if apply_coloring:
        apply_bands(patwari_by_pct, "pct")
        apply_bands(patwari_by_count, "submitted")

    # Patwari totals for TOTAL row — always from displayed rows so filtered/unfiltered both work
    if patwari_by_pct:
        p_tot = sum(r["total"] for r in patwari_by_pct)
        p_sub = sum(r["submitted"] for r in patwari_by_pct)
        patwari_totals = {
            "total": p_tot,
            "submitted": p_sub,
            "pct": _pct(p_sub, p_tot),
            "additions": sum(r["additions"] for r in patwari_by_pct) if has_additions else None,
        }
    else:
        patwari_totals = None

    # === Checker Wise ===
    # Columns: Checker | Sub-Division | Villages | Total Survey Nos | Submitted | Verified+Approved | % Completion | Additions
    # % Completion = (verified + approved + seek_clarification) / submitted * 100
    checker_rows = None
    checker_totals = None
    if "checker" in current_df.columns and current_df["checker"].str.strip().replace("", pd.NA).notna().any():
        c_df = current_df[current_df["checker"].str.strip() != ""].copy()
        c = (c_df.groupby("checker", as_index=False)
                 .agg(villages=("village", "count"),
                      villages_list=("village", lambda s: ", ".join(sorted(s.tolist()))),
                      subdivisions=("subdivision", lambda s: " / ".join(sorted(set(s)))),
                      total=("khasras", "sum"),
                      submitted=("submitted", "sum"),
                      verified=("verified", "sum"),
                      approved=("approved", "sum"),
                      seek_clarification=("seek_clarification", "sum"))
                 .sort_values("submitted", ascending=False))
        old_by_ck = from_df.groupby("checker")["submitted"].sum().to_dict() if has_additions and "checker" in from_df.columns else {}
        checker_rows = []
        for _, row in c.iterrows():
            v_plus_a = int(row["verified"]) + int(row["approved"])
            processed = v_plus_a + int(row["seek_clarification"])
            checker_rows.append({
                "name": row["checker"],
                "subdivision": row["subdivisions"],
                "villages_list": row["villages_list"],
                "villages": int(row["villages"]),
                "total": int(row["total"]),
                "submitted": int(row["submitted"]),
                "verified_plus_approved": v_plus_a,
                "processed_incl_seekclar": processed,
                "pct": _pct(processed, int(row["submitted"])),
                "additions": (int(row["submitted"]) - int(old_by_ck.get(row["checker"], 0))) if has_additions else None,
            })
        # Sort by V+A (the value column) descending — matches "checker's actual output"
        checker_rows.sort(key=lambda r: r["verified_plus_approved"], reverse=True)
        c_sub = sum(r["submitted"] for r in checker_rows)
        c_va = sum(r["verified_plus_approved"] for r in checker_rows)
        c_proc = sum(r["processed_incl_seekclar"] for r in checker_rows)
        checker_totals = {
            "villages": sum(r["villages"] for r in checker_rows),
            "total": sum(r["total"] for r in checker_rows),
            "submitted": c_sub,
            "verified_plus_approved": c_va,
            "pct": _pct(c_proc, c_sub),
            "additions": sum(r["additions"] for r in checker_rows) if has_additions else None,
        }

    # === Village Wise ===
    # Default: grouped tehsil-then-village. Filter/coloring: sorted by submitted desc + banding.
    if apply_coloring:
        if active_tehsils is not None:
            v_src = current_df[current_df["tehsil"].str.upper().isin(active_tehsils)]
        else:
            v_src = current_df
        v_df = v_src[["village", "tehsil", "patwari", "khasras", "submitted", "verified", "approved"]].copy()
        v_df = v_df.sort_values("submitted", ascending=False, kind="mergesort").reset_index(drop=True)
    else:
        v_df = current_df[["village", "tehsil", "patwari", "khasras", "submitted", "verified", "approved"]].copy()
        v_df = v_df.sort_values(["tehsil", "village"], kind="mergesort").reset_index(drop=True)
    old_by_village = from_df.groupby(["tehsil", "village"])["submitted"].sum().to_dict() if has_additions else {}
    village_rows = []
    for _, r in v_df.iterrows():
        additions = int(r["submitted"]) - int(old_by_village.get((r["tehsil"], r["village"]), 0)) if has_additions else None
        village_rows.append({
            "village": r["village"],
            "tehsil": r["tehsil"],
            "patwari": r["patwari"],
            "total": int(r["khasras"]),
            "submitted": int(r["submitted"]),
            "verified": int(r["verified"]) + int(r["approved"]),  # roll-in: verified+ = at-least-verified count
            "approved": int(r["approved"]),
            "additions": additions,
        })
    if apply_coloring:
        apply_bands(village_rows, "submitted")
    if village_rows:
        v_tot = sum(r["total"] for r in village_rows)
        v_sub = sum(r["submitted"] for r in village_rows)
        village_totals = {
            "total": v_tot,
            "submitted": v_sub,
            "verified": sum(r["verified"] for r in village_rows),
            "approved": sum(r["approved"] for r in village_rows),
            "pct": _pct(v_sub, v_tot),
            "additions": sum(r["additions"] for r in village_rows) if has_additions else None,
        }
    else:
        village_totals = None

    # Top cards / stats
    top_patwari = patwari_by_count[0] if patwari_by_count else None
    top_tehsil = tehsil_rows[0] if tehsil_rows else None
    grand = {
        "tehsils": int(len(t)),
        "villages": int(len(current_df)),
        "patwaris": int(current_df["patwari"].nunique()),
        "total_khasras": int(current_df["khasras"].sum()),
        "submitted": int(current_df["submitted"].sum()),
        "approved": int(current_df["approved"].sum()),
        "overall_pct": _pct(int(current_df["submitted"].sum()), int(current_df["khasras"].sum())),
        "overall_pct_approved": _pct(int(current_df["approved"].sum()), int(current_df["khasras"].sum())),
        "additions": (int(current_df["submitted"].sum()) - int(from_df["submitted"].sum())) if has_additions else None,
    }

    return {
        "tehsil_rows": tehsil_rows,
        "overall_rows": overall_rows,
        "overall_totals": overall_totals,
        "patwari_by_pct": patwari_by_pct,
        "patwari_by_count": patwari_by_count,
        "patwari_totals": patwari_totals,
        "checker_rows": checker_rows,
        "checker_totals": checker_totals,
        "village_rows": village_rows,
        "village_totals": village_totals,
        "top_patwari": top_patwari,
        "top_tehsil": top_tehsil,
        "grand": grand,
        "has_additions": has_additions,
    }


# ---------- Config ----------

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


def parse_tehsils_param(request_args):
    values = request_args.getlist("tehsils")
    if not values:
        return None
    cleaned = [v.strip() for v in values if v and v.strip()]
    if not cleaned:
        return None
    if any(v.lower() == "all" for v in cleaned):
        return "all"
    return set(v.upper() for v in cleaned)


def resolve_dates(snapshots, from_param, to_param):
    if not snapshots:
        return None, None
    dates_available = [d for d, _ in snapshots]
    to_date = parse_date_param(to_param) if to_param else None
    if not to_date or to_date not in dates_available:
        to_date = dates_available[0]
    from_date = parse_date_param(from_param) if from_param else None
    if not from_date or from_date not in dates_available or from_date >= to_date:
        older = [d for d in dates_available if d < to_date]
        from_date = older[0] if older else None
    return from_date, to_date


# ---------- Routes ----------

@app.route("/")
def index():
    snapshots = all_snapshots()
    if not snapshots:
        return "No snapshots found. Add a file to snapshots/ folder.", 500

    from_date, to_date = resolve_dates(snapshots, request.args.get("from"), request.args.get("to"))
    current_df = load_snapshot(snapshot_for_date(snapshots, to_date))
    from_df = load_snapshot(snapshot_for_date(snapshots, from_date)) if from_date else None
    tehsils_filter = parse_tehsils_param(request.args)
    views = build_views(current_df, from_df, tehsils_filter)

    override = read_as_of_override()
    as_of_display = override if override else format_date(to_date)
    date_options = [(d.isoformat(), format_date(d)) for d, _ in snapshots]
    additions_label = f"Additions ({format_date_short(from_date)} → {format_date_short(to_date)})" if views["has_additions"] else None

    all_tehsils = sorted(current_df["tehsil"].dropna().unique().tolist())
    is_all_selected = tehsils_filter == "all"
    selected_tehsils = tehsils_filter if isinstance(tehsils_filter, set) else set()

    return render_template(
        "index.html",
        as_of=as_of_display,
        date_options=date_options,
        from_date_iso=from_date.isoformat() if from_date else None,
        to_date_iso=to_date.isoformat(),
        additions_label=additions_label,
        snapshot_count=len(snapshots),
        all_tehsils=all_tehsils,
        is_all_selected=is_all_selected,
        selected_tehsils=selected_tehsils,
        coloring_active=(tehsils_filter is not None),
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


def _resolve_download_context():
    snapshots = all_snapshots()
    if not snapshots:
        return None
    from_date, to_date = resolve_dates(snapshots, request.args.get("from"), request.args.get("to"))
    current_df = load_snapshot(snapshot_for_date(snapshots, to_date))
    from_df = load_snapshot(snapshot_for_date(snapshots, from_date)) if from_date else None
    tehsils_filter = parse_tehsils_param(request.args)
    views = build_views(current_df, from_df, tehsils_filter)
    return {"views": views, "to_date": to_date, "from_date": from_date,
            "gap_days": (to_date - from_date).days if from_date else 0}


def _panels_for_view(views, view_key):
    """Return list of {title, type, key, rows, totals} for the view."""
    all_panels = [
        ("Tehsil Wise", "tehsil", views.get("tehsil_rows"), None, "tehsil"),
        ("Overall All", "overall", views.get("overall_rows"), views.get("overall_totals"), "overall-all"),
        ("Patwari Wise — By % Completion", "patwari", views.get("patwari_by_pct"), views.get("patwari_totals"), "patwari-pct"),
        ("Patwari Wise — By Total Submissions", "patwari", views.get("patwari_by_count"), views.get("patwari_totals"), "patwari-count"),
        ("Checker Wise", "checker", views.get("checker_rows"), views.get("checker_totals"), "checker"),
        ("Village Wise", "village", views.get("village_rows"), views.get("village_totals"), "village"),
    ]
    out = []
    for title, ptype, rows, totals, key in all_panels:
        if rows is None:
            continue
        if view_key != "all" and key != view_key:
            continue
        out.append({"title": title, "type": ptype, "rows": rows, "totals": totals, "key": key})
    return out


@app.route("/download.xlsx")
def download():
    if not _check_download_password():
        return Response("Authentication required", 401,
                        {"WWW-Authenticate": 'Basic realm="AgriStack Download"'})
    ctx = _resolve_download_context()
    if not ctx:
        return "No snapshots found.", 500
    view_key = request.args.get("view", "all")
    buf = _generate_workbook(ctx["views"], ctx["to_date"], ctx["from_date"], view_key)
    tail = "" if view_key == "all" else f"_{view_key}"
    filename = f"AGRISTACK_Dashboard_{ctx['to_date'].strftime('%Y-%m-%d')}{tail}.xlsx"
    return send_file(buf,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True, download_name=filename)


@app.route("/download.pdf")
def download_pdf():
    if not _check_download_password():
        return Response("Authentication required", 401,
                        {"WWW-Authenticate": 'Basic realm="AgriStack Download"'})
    ctx = _resolve_download_context()
    if not ctx:
        return "No snapshots found.", 500
    view_key = request.args.get("view", "tehsil")
    panels = _panels_for_view(ctx["views"], view_key)
    if not panels:
        return "No data for this view.", 400

    additions_label = None
    if ctx["gap_days"] > 0:
        additions_label = f"Additions ({ctx['from_date'].strftime('%d %b')} → {ctx['to_date'].strftime('%d %b')})"

    # Landscape works for every view here (Overall, Tehsil, Patwari, Checker, Village are all wide)
    orientation = "landscape"

    html = render_template("print.html",
                           panels=panels,
                           as_of=format_date(ctx["to_date"]),
                           from_date_display=format_date(ctx["from_date"]) if ctx["from_date"] else None,
                           gap_days=ctx["gap_days"],
                           has_additions=ctx["views"]["has_additions"],
                           additions_label=additions_label,
                           grand=ctx["views"]["grand"],
                           orientation=orientation)
    buf = io.BytesIO()
    result = pisa.CreatePDF(html, dest=buf)
    if result.err:
        return "PDF generation failed", 500
    buf.seek(0)
    tail = "all" if view_key == "all" else view_key
    filename = f"AGRISTACK_Dashboard_{ctx['to_date'].strftime('%Y-%m-%d')}_{tail}.pdf"
    return send_file(buf, mimetype="application/pdf",
                     as_attachment=True, download_name=filename)


# ---------- Excel workbook ----------

def _generate_workbook(views, to_date, from_date, view_key="all"):
    def _include(key):
        return view_key == "all" or view_key == key

    hdr_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    hdr_fill = PatternFill(start_color="1F3F2E", end_color="1F3F2E", fill_type="solid")
    hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    body_font = Font(name="Arial", size=11)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    tot_font = Font(name="Arial", size=11, bold=True)
    tot_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    thin = Side(border_style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    band_fills = {
        "green":  PatternFill(start_color="DDF0D8", end_color="DDF0D8", fill_type="solid"),
        "yellow": PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid"),
        "red":    PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid"),
    }

    def write_hdr(ws, headers):
        for i, h in enumerate(headers, start=1):
            c = ws.cell(row=1, column=i, value=h)
            c.font = hdr_font; c.fill = hdr_fill; c.alignment = hdr_align; c.border = border
        ws.row_dimensions[1].height = 38

    def apply_band_fill(ws, row_num, num_cols, band):
        if not band or band not in band_fills:
            return
        for ci in range(1, num_cols + 1):
            ws.cell(row=row_num, column=ci).fill = band_fills[band]

    wb = Workbook()
    wb.remove(wb.active)
    grand = views["grand"]

    additions_hdr = None
    if views["has_additions"]:
        additions_hdr = f"Additions ({from_date.strftime('%d %b')} to {to_date.strftime('%d %b')})"

    # === TEHSIL WISE ===
    if _include("tehsil"):
        ws = wb.create_sheet("TEHSIL WISE")
        headers = ["S.NO", "TEHSIL", "NUMBER OF VILLAGES", "TOTAL SURVEY NOS",
                   "SUBMITTED", "% COMPLETION"]
        if additions_hdr:
            headers.append(additions_hdr)
        headers.append("% APPROVED")
        write_hdr(ws, headers)
        approved_col = len(headers)
        additions_col = len(headers) - 1 if additions_hdr else None
        for i, row in enumerate(views["tehsil_rows"], start=1):
            r = i + 1
            vals = [i, row["tehsil"], row["villages"], row["total"], row["submitted"], round(row["pct"], 2)]
            aligns = [center, left, center, center, center, center]
            if additions_hdr:
                vals.append(row["additions"] if row["additions"] is not None else "—")
                aligns.append(center)
            vals.append(round(row["pct_approved"], 2))
            aligns.append(center)
            for ci, (v, a) in enumerate(zip(vals, aligns), start=1):
                c = ws.cell(row=r, column=ci, value=v)
                c.alignment = a; c.font = body_font; c.border = border
                if ci in (3, 4, 5): c.number_format = "#,##0"
                if ci == 6: c.number_format = '0.00"%"'
                if additions_col and ci == additions_col and isinstance(v, int):
                    c.number_format = "+#,##0;-#,##0;0"
                if ci == approved_col: c.number_format = '0.00"%"'
        # TOTAL
        tt = len(views["tehsil_rows"]) + 2
        ws.cell(row=tt, column=2, value="TOTAL").alignment = left
        ws.cell(row=tt, column=3, value=grand["villages"]).alignment = center
        ws.cell(row=tt, column=4, value=grand["total_khasras"]).alignment = center
        ws.cell(row=tt, column=5, value=grand["submitted"]).alignment = center
        ws.cell(row=tt, column=6, value=round(grand["overall_pct"], 2)).alignment = center
        if additions_col:
            ws.cell(row=tt, column=additions_col, value=grand["additions"] if grand["additions"] is not None else "—").alignment = center
        ws.cell(row=tt, column=approved_col, value=round(grand["overall_pct_approved"], 2)).alignment = center
        for c in range(1, len(headers) + 1):
            cc = ws.cell(row=tt, column=c)
            cc.font = tot_font; cc.fill = tot_fill; cc.border = border
            if c in (3, 4, 5): cc.number_format = "#,##0"
            if c == 6: cc.number_format = '0.00"%"'
            if additions_col and c == additions_col and isinstance(cc.value, int):
                cc.number_format = "+#,##0;-#,##0;0"
            if c == approved_col: cc.number_format = '0.00"%"'
        widths = [7, 16, 16, 18, 14, 14]
        if additions_hdr: widths.append(22)
        widths.append(14)
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"

    # === OVERALL ALL ===
    if _include("overall-all"):
        ws = wb.create_sheet("OVERALL ALL")
        headers = ["S.NO", "TEHSIL", "TOTAL KHASRAS", "KHASRAS SUBMITTED",
                   "KHASRAS VERIFIED", "KHASRAS APPROVED"]
        write_hdr(ws, headers)
        for i, row in enumerate(views["overall_rows"], start=1):
            r = i + 1
            vals = [i, row["tehsil"], row["total"], row["submitted"], row["verified"], row["approved"]]
            aligns = [center, left, center, center, center, center]
            for ci, (v, a) in enumerate(zip(vals, aligns), start=1):
                c = ws.cell(row=r, column=ci, value=v)
                c.alignment = a; c.font = body_font; c.border = border
                if ci in (3, 4, 5, 6): c.number_format = "#,##0"
        ot = views["overall_totals"]
        tt = len(views["overall_rows"]) + 2
        ws.cell(row=tt, column=2, value="TOTAL").alignment = left
        ws.cell(row=tt, column=3, value=ot["total"]).alignment = center
        ws.cell(row=tt, column=4, value=ot["submitted"]).alignment = center
        ws.cell(row=tt, column=5, value=ot["verified"]).alignment = center
        ws.cell(row=tt, column=6, value=ot["approved"]).alignment = center
        for c in range(1, 7):
            cc = ws.cell(row=tt, column=c)
            cc.font = tot_font; cc.fill = tot_fill; cc.border = border
            if c in (3, 4, 5, 6): cc.number_format = "#,##0"
        for i, w in enumerate([7, 16, 18, 20, 20, 20], start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"

    # === Patwari sheets ===
    for title, rows, key in [("PATWARI BY %", views["patwari_by_pct"], "patwari-pct"),
                              ("PATWARI BY COUNT", views["patwari_by_count"], "patwari-count")]:
        if not _include(key):
            continue
        show_effort = (key == "patwari-pct")
        ws = wb.create_sheet(title)
        headers = ["S.NO", "NAME OF PATWARI", "VILLAGES", "TOTAL SURVEY NOS",
                   "SUBMITTED", "% COMPLETION"]
        if show_effort:
            headers.append("RELATIVE EFFORT")
        if additions_hdr:
            headers.append(additions_hdr)
        write_hdr(ws, headers)
        effort_col = 7 if show_effort else None
        additions_col = len(headers) if additions_hdr else None
        total_cols = len(headers)
        for i, row in enumerate(rows, start=1):
            r = i + 1
            vals = [i, row["patwari"], row["villages_list"], row["total"],
                    row["submitted"], round(row["pct"], 2)]
            aligns = [center, left, left, center, center, center]
            if show_effort:
                re_v = row.get("relative_effort")
                vals.append(round(re_v, 2) if re_v is not None else "—")
                aligns.append(center)
            if additions_hdr:
                vals.append(row["additions"] if row["additions"] is not None else "—")
                aligns.append(center)
            for ci, (v, a) in enumerate(zip(vals, aligns), start=1):
                c = ws.cell(row=r, column=ci, value=v)
                c.alignment = a; c.font = body_font; c.border = border
                if ci in (4, 5): c.number_format = "#,##0"
                if ci == 6: c.number_format = '0.00"%"'
                if effort_col and ci == effort_col and isinstance(v, (int, float)):
                    c.number_format = '0.00"%"'
                if additions_col and ci == additions_col and isinstance(v, int):
                    c.number_format = "+#,##0;-#,##0;0"
            apply_band_fill(ws, r, total_cols, row.get("band"))
        # TOTAL
        pt = len(rows) + 2
        pt_tot = views.get("patwari_totals")
        use_filtered = pt_tot is not None and any(r.get("band") for r in rows)
        ws.cell(row=pt, column=2, value=("TOTAL (Filtered)" if use_filtered else "TOTAL")).alignment = left
        if use_filtered:
            ws.cell(row=pt, column=4, value=pt_tot["total"]).alignment = center
            ws.cell(row=pt, column=5, value=pt_tot["submitted"]).alignment = center
            ws.cell(row=pt, column=6, value=round(pt_tot["pct"], 2)).alignment = center
            if additions_col:
                ws.cell(row=pt, column=additions_col, value=pt_tot["additions"] if pt_tot["additions"] is not None else "—").alignment = center
        else:
            ws.cell(row=pt, column=4, value=grand["total_khasras"]).alignment = center
            ws.cell(row=pt, column=5, value=grand["submitted"]).alignment = center
            ws.cell(row=pt, column=6, value=round(grand["overall_pct"], 2)).alignment = center
            if additions_col:
                ws.cell(row=pt, column=additions_col, value=grand["additions"] if grand["additions"] is not None else "—").alignment = center
        for c in range(1, total_cols + 1):
            cc = ws.cell(row=pt, column=c)
            cc.font = tot_font; cc.fill = tot_fill; cc.border = border
            if c in (4, 5): cc.number_format = "#,##0"
            if c == 6: cc.number_format = '0.00"%"'
            if show_effort and c == effort_col: cc.number_format = '0.00"%"'
            if additions_col and c == additions_col and isinstance(cc.value, int):
                cc.number_format = "+#,##0;-#,##0;0"
        widths = [7, 26, 50, 18, 14, 14]
        if show_effort: widths.append(16)
        if additions_hdr: widths.append(22)
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"

    # === CHECKER WISE ===
    if _include("checker") and views.get("checker_rows"):
        ws = wb.create_sheet("CHECKER WISE")
        headers = ["S.NO", "CHECKER", "SUB-DIVISION", "VILLAGES", "TOTAL SURVEY NOS",
                   "SUBMITTED", "VERIFIED + APPROVED", "% COMPLETION"]
        if additions_hdr:
            headers.append(additions_hdr)
        write_hdr(ws, headers)
        additions_col = len(headers) if additions_hdr else None
        # Formula note as row 2 (comment above table)
        note_row = 2
        ws.cell(row=note_row, column=1, value="% Completion = (Verified + Approved + Seek Clarification) ÷ Submitted × 100").alignment = left
        ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=len(headers))
        ws.cell(row=note_row, column=1).font = Font(name="Arial", size=10, italic=True, color="4A6B4E")
        data_start = 3
        for i, row in enumerate(views["checker_rows"], start=1):
            r = i + data_start - 1
            vals = [i, row["name"], row["subdivision"], row["villages_list"],
                    row["total"], row["submitted"], row["verified_plus_approved"], round(row["pct"], 2)]
            aligns = [center, left, left, left, center, center, center, center]
            if additions_hdr:
                vals.append(row["additions"] if row["additions"] is not None else "—")
                aligns.append(center)
            for ci, (v, a) in enumerate(zip(vals, aligns), start=1):
                c = ws.cell(row=r, column=ci, value=v)
                c.alignment = a; c.font = body_font; c.border = border
                if ci in (5, 6, 7): c.number_format = "#,##0"
                if ci == 8: c.number_format = '0.00"%"'
                if additions_col and ci == additions_col and isinstance(v, int):
                    c.number_format = "+#,##0;-#,##0;0"
        ct = views.get("checker_totals")
        if ct:
            pt = data_start + len(views["checker_rows"])
            ws.cell(row=pt, column=2, value="TOTAL").alignment = left
            ws.cell(row=pt, column=5, value=ct["total"]).alignment = center
            ws.cell(row=pt, column=6, value=ct["submitted"]).alignment = center
            ws.cell(row=pt, column=7, value=ct["verified_plus_approved"]).alignment = center
            ws.cell(row=pt, column=8, value=round(ct["pct"], 2)).alignment = center
            if additions_col:
                ws.cell(row=pt, column=additions_col, value=ct["additions"] if ct["additions"] is not None else "—").alignment = center
            for c in range(1, len(headers) + 1):
                cc = ws.cell(row=pt, column=c)
                cc.font = tot_font; cc.fill = tot_fill; cc.border = border
                if c in (5, 6, 7): cc.number_format = "#,##0"
                if c == 8: cc.number_format = '0.00"%"'
                if additions_col and c == additions_col and isinstance(cc.value, int):
                    cc.number_format = "+#,##0;-#,##0;0"
        widths = [7, 32, 14, 46, 16, 14, 18, 14]
        if additions_hdr: widths.append(22)
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = f"A{data_start}"

    # === VILLAGE WISE ===
    if _include("village"):
        ws = wb.create_sheet("VILLAGE WISE")
        headers = ["S.NO", "TEHSIL", "VILLAGE", "TOTAL SURVEY NOS", "NAME OF PATWARI",
                   "SUBMITTED", "VERIFIED", "APPROVED"]
        if additions_hdr:
            headers.append(additions_hdr)
        write_hdr(ws, headers)
        additions_col = len(headers) if additions_hdr else None
        total_cols = len(headers)
        for i, row in enumerate(views["village_rows"], start=1):
            r = i + 1
            vals = [i, row["tehsil"], row["village"], row["total"], row["patwari"],
                    row["submitted"], row["verified"], row["approved"]]
            aligns = [center, left, left, center, left, center, center, center]
            if additions_hdr:
                vals.append(row["additions"] if row["additions"] is not None else "—")
                aligns.append(center)
            for ci, (v, a) in enumerate(zip(vals, aligns), start=1):
                c = ws.cell(row=r, column=ci, value=v)
                c.alignment = a; c.font = body_font; c.border = border
                if ci in (4, 6, 7, 8): c.number_format = "#,##0"
                if additions_col and ci == additions_col and isinstance(v, int):
                    c.number_format = "+#,##0;-#,##0;0"
            apply_band_fill(ws, r, total_cols, row.get("band"))
        vt = len(views["village_rows"]) + 2
        v_tot = views.get("village_totals")
        use_filtered = v_tot is not None and any(r.get("band") for r in views["village_rows"])
        ws.cell(row=vt, column=2, value=("TOTAL (Filtered)" if use_filtered else "TOTAL")).alignment = left
        if use_filtered:
            ws.cell(row=vt, column=4, value=v_tot["total"]).alignment = center
            ws.cell(row=vt, column=6, value=v_tot["submitted"]).alignment = center
            ws.cell(row=vt, column=7, value=v_tot["verified"]).alignment = center
            ws.cell(row=vt, column=8, value=v_tot["approved"]).alignment = center
            if additions_col:
                ws.cell(row=vt, column=additions_col, value=v_tot["additions"] if v_tot["additions"] is not None else "—").alignment = center
        else:
            ws.cell(row=vt, column=4, value=grand["total_khasras"]).alignment = center
            ws.cell(row=vt, column=6, value=grand["submitted"]).alignment = center
            ov = views["overall_totals"]
            ws.cell(row=vt, column=7, value=ov["verified"]).alignment = center
            ws.cell(row=vt, column=8, value=ov["approved"]).alignment = center
            if additions_col:
                ws.cell(row=vt, column=additions_col, value=grand["additions"] if grand["additions"] is not None else "—").alignment = center
        for c in range(1, total_cols + 1):
            cc = ws.cell(row=vt, column=c)
            cc.font = tot_font; cc.fill = tot_fill; cc.border = border
            if c in (4, 6, 7, 8): cc.number_format = "#,##0"
            if additions_col and c == additions_col and isinstance(cc.value, int):
                cc.number_format = "+#,##0;-#,##0;0"
        widths = [7, 14, 28, 20, 28, 14, 14, 14]
        if additions_hdr: widths.append(22)
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"

    # === META (always) ===
    ws4 = wb.create_sheet("META")
    ws4.cell(row=1, column=1, value="Data as of").font = tot_font
    ws4.cell(row=1, column=2, value=to_date.strftime("%d %b %Y"))
    ws4.cell(row=2, column=1, value="Additions from").font = tot_font
    ws4.cell(row=2, column=2, value=from_date.strftime("%d %b %Y") if from_date else "—")
    ws4.column_dimensions["A"].width = 18
    ws4.column_dimensions["B"].width = 20

    if len(wb.sheetnames) == 1:
        info = wb.create_sheet("INFO", 0)
        info.cell(row=1, column=1, value=f"No data for view: {view_key}")

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


@app.route("/healthz")
def healthz():
    return "ok", 200


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
