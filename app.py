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
from xhtml2pdf import pisa

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


def build_views(current_df, from_df=None, daily_target=DEFAULT_DAILY_TARGET, tehsils_filter=None):
    """Build all dashboard views.
    from_df: optional older snapshot for 'Additions'.
    daily_target: district-level daily target.
    tehsils_filter: None (no filter, no coloring), 'all' (all tehsils, coloring on),
                    or set of tehsil names (filter + coloring)."""
    apply_coloring = tehsils_filter is not None
    active_tehsils = tehsils_filter if isinstance(tehsils_filter, set) else None
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

    # --- Patwari-wise (two sorts) — filtered by selected tehsils when picker is set ---
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
        old_by_p_filtered = p_old_source.groupby("patwari")["submitted"].sum().to_dict()
        p["additions"] = p["patwari"].map(lambda n: int(old_by_p_filtered.get(n, 0)))
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

    # Relative Effort: each patwari's submitted count as a % of the district-wide top
    # submitter. Only attached to patwari_by_pct — that's where % Completion misleads.
    # District-wide means the reference doesn't change when tehsil picker is on.
    district_p = (current_df.groupby("patwari", as_index=False)["submitted"].sum())
    if len(district_p) and district_p["submitted"].max() > 0:
        top_district_submitted = int(district_p["submitted"].max())
    else:
        top_district_submitted = 0
    for r in patwari_by_pct:
        if top_district_submitted > 0:
            r["relative_effort"] = (r["submitted"] / top_district_submitted) * 100
        else:
            r["relative_effort"] = None

    if apply_coloring:
        apply_bands(patwari_by_pct, "pct")
        apply_bands(patwari_by_count, "submitted")

    # Totals used in the Patwari Wise TOTAL row — always sum from displayed rows,
    # so they honestly reflect what's shown (filtered when picker set, district when not)
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

    # --- Village-wise ---
    # Default (no picker): grouped by tehsil then village (alphabetical)
    # When picker is set (either 'all' or a set of tehsils): sorted strictly by submitted desc,
    # filtered to the selected tehsils when a set is given.
    if apply_coloring:
        if active_tehsils is not None:
            v_df = current_df[current_df["tehsil"].str.upper().isin(active_tehsils)][
                ["village", "tehsil", "patwari", "khasras", "submitted"]].copy()
        else:
            v_df = current_df[["village", "tehsil", "patwari", "khasras", "submitted"]].copy()
        v_df = v_df.sort_values("submitted", ascending=False, kind="mergesort").reset_index(drop=True)
    else:
        v_df = current_df[["village", "tehsil", "patwari", "khasras", "submitted"]].copy()
        v_df = v_df.sort_values(["tehsil", "village"], kind="mergesort").reset_index(drop=True)
    # Village-level additions: match on (tehsil, village) — some village names appear in multiple tehsils
    old_by_village = {}
    if has_additions:
        old_by_village = from_df.groupby(["tehsil", "village"])["submitted"].sum().to_dict()
    village_rows = []
    for _, r in v_df.iterrows():
        additions = None
        if has_additions:
            key = (r["tehsil"], r["village"])
            additions = int(r["submitted"]) - int(old_by_village.get(key, 0))
        village_rows.append({
            "village": r["village"],
            "tehsil": r["tehsil"],
            "patwari": r["patwari"],
            "total": int(r["khasras"]),
            "submitted": int(r["submitted"]),
            "additions": additions,
        })
    if apply_coloring:
        apply_bands(village_rows, "submitted")

    # Village Wise TOTAL row — sum from displayed rows for honest filtered totals
    if village_rows:
        v_tot = sum(r["total"] for r in village_rows)
        v_sub = sum(r["submitted"] for r in village_rows)
        village_totals = {
            "total": v_tot,
            "submitted": v_sub,
            "pct": _pct(v_sub, v_tot),
            "additions": sum(r["additions"] for r in village_rows) if has_additions else None,
        }
    else:
        village_totals = None

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

    # --- Sumbal Patwaris — patwari-wise filtered to SUMBAL tehsil, sorted by submitted desc ---
    sumbal_patwari_rows = None
    sumbal_patwari_totals = None
    sumbal_df = current_df[current_df["tehsil"].str.upper() == "SUMBAL"]
    if len(sumbal_df):
        sp = (sumbal_df.groupby("patwari", as_index=False)
                       .agg(villages_list=("village", lambda s: ", ".join(sorted(s.tolist()))),
                            villages=("village", "count"),
                            total=("khasras", "sum"),
                            submitted=("submitted", "sum")))
        sp["pct"] = sp.apply(lambda r: _pct(int(r["submitted"]), int(r["total"])), axis=1)
        # Additions — compute against SUMBAL rows in old snapshot only (avoids cross-tehsil name collision)
        old_by_patwari_sumbal = {}
        if has_additions:
            old_sumbal = from_df[from_df["tehsil"].str.upper() == "SUMBAL"]
            old_by_patwari_sumbal = old_sumbal.groupby("patwari")["submitted"].sum().to_dict()
        sp = sp.sort_values(["submitted", "pct"], ascending=[False, False])
        sumbal_patwari_rows = []
        for _, r in sp.iterrows():
            additions = None
            if has_additions:
                additions = int(r["submitted"]) - int(old_by_patwari_sumbal.get(r["patwari"], 0))
            sumbal_patwari_rows.append({
                "patwari": r["patwari"],
                "villages_list": r["villages_list"],
                "villages": int(r["villages"]),
                "total": int(r["total"]),
                "submitted": int(r["submitted"]),
                "pct": float(r["pct"]),
                "additions": additions,
            })
        # Totals for the Sumbal-only footer
        subs = sum(r["submitted"] for r in sumbal_patwari_rows)
        tot = sum(r["total"] for r in sumbal_patwari_rows)
        sumbal_patwari_totals = {
            "villages": sum(r["villages"] for r in sumbal_patwari_rows),
            "total": tot,
            "submitted": subs,
            "pct": _pct(subs, tot),
            "additions": sum(r["additions"] for r in sumbal_patwari_rows) if has_additions else None,
        }
        # Sumbal Patwaris is always color-banded (by Submitted, high→low)
        apply_bands(sumbal_patwari_rows, "submitted")

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
        "patwari_totals": patwari_totals,
        "checker_rows": checker_rows,
        "checker_totals": checker_totals,
        "subdiv_rows": subdiv_rows,
        "subdiv_totals": subdiv_totals,
        "village_rows": village_rows,
        "village_totals": village_totals,
        "sumbal_patwari_rows": sumbal_patwari_rows,
        "sumbal_patwari_totals": sumbal_patwari_totals,
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


def parse_tehsils_param(request_args):
    """Read tehsils checkbox selection from query args.
    Returns None (no filter, no coloring), 'all' (all tehsils, coloring on),
    or a set of tehsil names (filter + coloring)."""
    values = request_args.getlist("tehsils")
    if not values:
        return None
    cleaned = [v.strip() for v in values if v and v.strip()]
    if not cleaned:
        return None
    if any(v.lower() == "all" for v in cleaned):
        return "all"
    return set(v.upper() for v in cleaned)


def apply_bands(rows, sort_key):
    """Attach a 'band' key ('green'|'yellow'|'red') to each row.
    Assumes rows are already sorted high-to-low by sort_key.
    Top 30% green, bottom 30% red, middle 40% yellow (by row count)."""
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

    tehsils_filter = parse_tehsils_param(request.args)

    daily_target = read_daily_target()
    views = build_views(current_df, from_df, daily_target, tehsils_filter)

    override = read_as_of_override()
    as_of_display = override if override else format_date(to_date)

    date_options = [(d.isoformat(), format_date(d)) for d, _ in snapshots]

    additions_label = None
    if views["has_additions"]:
        additions_label = f"Additions ({format_date_short(from_date)} → {format_date_short(to_date)})"

    # Picker state for the template
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
    """Common setup shared by download.xlsx and download.pdf."""
    snapshots = all_snapshots()
    if not snapshots:
        return None
    from_date, to_date = resolve_dates(snapshots,
                                        request.args.get("from"),
                                        request.args.get("to"))
    to_path = snapshot_for_date(snapshots, to_date)
    current_df = load_snapshot(to_path)
    from_df = load_snapshot(snapshot_for_date(snapshots, from_date)) if from_date else None
    tehsils_filter = parse_tehsils_param(request.args)
    daily_target = read_daily_target()
    views = build_views(current_df, from_df, daily_target, tehsils_filter)
    return {
        "views": views,
        "to_date": to_date,
        "from_date": from_date,
        "gap_days": (to_date - from_date).days if from_date else 0,
        "tehsils_filter": tehsils_filter,
    }


def _panels_for_view(views, view_key):
    """Return a list of {title, type, rows, totals} for the given view_key.
    'all' returns every populated panel in dashboard order.
    totals is None for panels that use district-wide grand totals."""
    all_panels = [
        ("Tehsil Wise", "tehsil", views.get("tehsil_rows"), None, "tehsil"),
        ("Patwari Wise — By % Completion", "patwari", views.get("patwari_by_pct"), None, "patwari-pct"),
        ("Patwari Wise — By Total Submissions", "patwari", views.get("patwari_by_count"), None, "patwari-count"),
        ("Checker Wise", "checker", views.get("checker_rows"), views.get("checker_totals"), "checker"),
        ("Sub-Division Wise", "subdivision", views.get("subdiv_rows"), views.get("subdiv_totals"), "subdivision"),
        ("Village Wise", "village", views.get("village_rows"), None, "village"),
        ("Not Started Villages", "not-started", views.get("not_started"), None, "not-started"),
        ("Sumbal Patwaris", "patwari", views.get("sumbal_patwari_rows"), views.get("sumbal_patwari_totals"), "sumbal-patwaris"),
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
        return Response(
            "Authentication required",
            401,
            {"WWW-Authenticate": 'Basic realm="AgriStack Download"'},
        )
    ctx = _resolve_download_context()
    if not ctx:
        return "No snapshots found.", 500
    view_key = request.args.get("view", "all")
    buf = _generate_workbook(ctx["views"], ctx["to_date"], ctx["from_date"], view_key)
    tail = "" if view_key == "all" else f"_{view_key}"
    filename = f"AGRISTACK_Dashboard_{ctx['to_date'].strftime('%Y-%m-%d')}{tail}.xlsx"
    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename,
    )


@app.route("/download.pdf")
def download_pdf():
    if not _check_download_password():
        return Response(
            "Authentication required",
            401,
            {"WWW-Authenticate": 'Basic realm="AgriStack Download"'},
        )
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

    # Pick orientation based on what fits the content best.
    # Wide tables (many columns / villages list) need landscape; narrow ones look better in portrait.
    LANDSCAPE_VIEWS = {"tehsil", "patwari-pct", "patwari-count", "checker", "village", "sumbal-patwaris", "all"}
    orientation = "landscape" if view_key in LANDSCAPE_VIEWS else "portrait"

    html = render_template(
        "print.html",
        panels=panels,
        as_of=format_date(ctx["to_date"]),
        from_date_display=format_date(ctx["from_date"]) if ctx["from_date"] else None,
        gap_days=ctx["gap_days"],
        has_additions=ctx["views"]["has_additions"],
        additions_label=additions_label,
        grand=ctx["views"]["grand"],
        checker_totals=ctx["views"].get("checker_totals"),
        subdiv_totals=ctx["views"].get("subdiv_totals"),
        orientation=orientation,
    )
    buf = io.BytesIO()
    result = pisa.CreatePDF(html, dest=buf)
    if result.err:
        return "PDF generation failed", 500
    buf.seek(0)
    tail = "all" if view_key == "all" else view_key
    filename = f"AGRISTACK_Dashboard_{ctx['to_date'].strftime('%Y-%m-%d')}_{tail}.pdf"
    return send_file(buf, mimetype="application/pdf",
                     as_attachment=True, download_name=filename)


def _generate_workbook(views, to_date, from_date, view_key="all"):
    """Generate xlsx. view_key='all' includes every sheet; otherwise just the selected one plus META."""
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
    ns_fill = PatternFill(start_color="A54B2A", end_color="A54B2A", fill_type="solid")
    thin = Side(border_style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Band fills for filtered/coloring views — top 30% green, mid 40% yellow, bot 30% red
    band_fills = {
        "green":  PatternFill(start_color="DDF0D8", end_color="DDF0D8", fill_type="solid"),
        "yellow": PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid"),
        "red":    PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid"),
    }

    def apply_band_fill(ws, row_num, num_cols, band):
        """Apply band fill to all cells in a row (call after normal styling)."""
        if not band or band not in band_fills:
            return
        for ci in range(1, num_cols + 1):
            ws.cell(row=row_num, column=ci).fill = band_fills[band]

    def write_hdr(ws, headers, fill=hdr_fill):
        for i, h in enumerate(headers, start=1):
            c = ws.cell(row=1, column=i, value=h)
            c.font = hdr_font; c.fill = fill; c.alignment = hdr_align; c.border = border
        ws.row_dimensions[1].height = 38

    wb = Workbook()
    wb.remove(wb.active)  # drop default sheet — we create in order
    grand = views["grand"]

    additions_hdr = None
    if views["has_additions"]:
        additions_hdr = f"Additions ({from_date.strftime('%d %b')} to {to_date.strftime('%d %b')})"

    # === TEHSIL WISE ===
    if _include("tehsil"):
        ws1 = wb.create_sheet("TEHSIL WISE")
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
    for title, rows, key in [("PATWARI BY %", views["patwari_by_pct"], "patwari-pct"),
                              ("PATWARI BY COUNT", views["patwari_by_count"], "patwari-count")]:
        if not _include(key):
            continue
        show_effort = (key == "patwari-pct")
        ws = wb.create_sheet(title)
        p_headers = ["S.NO", "NAME OF PATWARI", "VILLAGES", "TOTAL SURVEY NOS",
                     "SUBMITTED", "% COMPLETION"]
        if show_effort:
            p_headers.append("RELATIVE EFFORT")
        if additions_hdr:
            p_headers.append(additions_hdr)
        write_hdr(ws, p_headers)
        # Column indexes shift when Relative Effort is present
        effort_col = 7 if show_effort else None
        additions_col = 8 if show_effort and additions_hdr else (7 if additions_hdr else None)
        total_cols = len(p_headers)
        for i, row in enumerate(rows, start=1):
            r = i + 1
            vals = [i, row["patwari"], row["villages_list"], row["total"],
                    row["submitted"], round(row["pct"], 2)]
            aligns = [center, left, left, center, center, center]
            if show_effort:
                re_val = row.get("relative_effort")
                vals.append(round(re_val, 2) if re_val is not None else "—")
                aligns.append(center)
            if additions_hdr:
                vals.append(row["additions"] if row["additions"] is not None else "—")
                aligns.append(center)
            for ci, (v, a) in enumerate(zip(vals, aligns), start=1):
                c = ws.cell(row=r, column=ci, value=v)
                c.alignment = a; c.font = body_font; c.border = border
                if ci in (4, 5): c.number_format = "#,##0"
                if ci == 6: c.number_format = '0.00"%"'
                if show_effort and ci == effort_col and isinstance(v, (int, float)):
                    c.number_format = '0.00"%"'
                if additions_col and ci == additions_col and isinstance(v, int):
                    c.number_format = "+#,##0;-#,##0;0"
            apply_band_fill(ws, r, total_cols, row.get("band"))
        # TOTAL row
        pt = len(rows) + 2
        pt_totals = views.get("patwari_totals") if views.get("patwari_totals") else None
        use_filtered = pt_totals is not None and any(r.get("band") for r in rows)
        tot_label = "TOTAL (Filtered)" if use_filtered else "TOTAL"
        ws.cell(row=pt, column=2, value=tot_label).alignment = left
        if use_filtered:
            ws.cell(row=pt, column=4, value=pt_totals["total"]).alignment = center
            ws.cell(row=pt, column=5, value=pt_totals["submitted"]).alignment = center
            ws.cell(row=pt, column=6, value=round(pt_totals["pct"], 2)).alignment = center
            if additions_col:
                ws.cell(row=pt, column=additions_col, value=pt_totals["additions"] if pt_totals["additions"] is not None else "—").alignment = center
        else:
            ws.cell(row=pt, column=4, value=grand["total_khasras"]).alignment = center
            ws.cell(row=pt, column=5, value=grand["submitted"]).alignment = center
            ws.cell(row=pt, column=6, value=round(grand["overall_pct"], 2)).alignment = center
            if additions_col:
                ws.cell(row=pt, column=additions_col, value=grand["additions"] if grand["additions"] is not None else "—").alignment = center
        # Style the TOTAL row across all columns present
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
        widths = [7, 32, 14, 46, 16, 14, 14, 14]
        if additions_hdr: widths.append(22)
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"

    # === SUB-DIVISION WISE ===
    if _include("subdivision") and views.get("subdiv_rows"):
        ws = wb.create_sheet("SUB-DIVISION WISE")
        headers = ["S.NO", "SUB-DIVISION", "NUMBER OF VILLAGES", "TOTAL SURVEY NOS",
                   "DAILY TARGET", "SUBMITTED", "% COMPLETION"]
        if additions_hdr:
            headers.append(additions_hdr)
        write_hdr(ws, headers)
        for i, row in enumerate(views["subdiv_rows"], start=1):
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
        st = views.get("subdiv_totals")
        if st:
            pt = len(views["subdiv_rows"]) + 2
            ws.cell(row=pt, column=2, value="TOTAL").alignment = left
            ws.cell(row=pt, column=3, value=st["villages"]).alignment = center
            ws.cell(row=pt, column=4, value=st["total"]).alignment = center
            ws.cell(row=pt, column=5, value=st["daily_target"]).alignment = center
            ws.cell(row=pt, column=6, value=st["submitted"]).alignment = center
            ws.cell(row=pt, column=7, value=round(st["pct"], 2)).alignment = center
            if additions_hdr:
                ws.cell(row=pt, column=8, value=st["additions"] if st["additions"] is not None else "—").alignment = center
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

    # === VILLAGE WISE ===
    if _include("village"):
        ws_v = wb.create_sheet("VILLAGE WISE")
        v_headers = ["S.NO", "TEHSIL", "VILLAGE", "TOTAL SURVEY NOS", "NAME OF PATWARI", "SUBMITTED"]
        if additions_hdr:
            v_headers.append(additions_hdr)
        write_hdr(ws_v, v_headers)
        for i, row in enumerate(views["village_rows"], start=1):
            r = i + 1
            vals = [i, row["tehsil"], row["village"], row["total"], row["patwari"], row["submitted"]]
            aligns = [center, left, left, center, left, center]
            if additions_hdr:
                vals.append(row["additions"] if row["additions"] is not None else "—")
                aligns.append(center)
            for ci, (v, a) in enumerate(zip(vals, aligns), start=1):
                c = ws_v.cell(row=r, column=ci, value=v)
                c.alignment = a; c.font = body_font; c.border = border
                if ci in (4, 6): c.number_format = "#,##0"
                if additions_hdr and ci == 7 and isinstance(v, int):
                    c.number_format = "+#,##0;-#,##0;0"
            apply_band_fill(ws_v, r, 7 if additions_hdr else 6, row.get("band"))
        vt = len(views["village_rows"]) + 2
        v_totals = views.get("village_totals")
        v_use_filtered = v_totals is not None and any(r.get("band") for r in views["village_rows"])
        v_label = "TOTAL (Filtered)" if v_use_filtered else "TOTAL"
        ws_v.cell(row=vt, column=2, value=v_label).alignment = left
        if v_use_filtered:
            ws_v.cell(row=vt, column=4, value=v_totals["total"]).alignment = center
            ws_v.cell(row=vt, column=6, value=v_totals["submitted"]).alignment = center
            if additions_hdr:
                ws_v.cell(row=vt, column=7, value=v_totals["additions"] if v_totals["additions"] is not None else "—").alignment = center
        else:
            ws_v.cell(row=vt, column=4, value=grand["total_khasras"]).alignment = center
            ws_v.cell(row=vt, column=6, value=grand["submitted"]).alignment = center
            if additions_hdr:
                ws_v.cell(row=vt, column=7, value=grand["additions"] if grand["additions"] is not None else "—").alignment = center
        tot_col_max = 7 if additions_hdr else 6
        for c in range(1, tot_col_max + 1):
            cc = ws_v.cell(row=vt, column=c)
            cc.font = tot_font; cc.fill = tot_fill; cc.border = border
            if c in (4, 6): cc.number_format = "#,##0"
            if additions_hdr and c == 7 and isinstance(cc.value, int):
                cc.number_format = "+#,##0;-#,##0;0"
        widths = [7, 14, 28, 20, 28, 16]
        if additions_hdr: widths.append(22)
        for i, w in enumerate(widths, start=1):
            ws_v.column_dimensions[get_column_letter(i)].width = w
        ws_v.freeze_panes = "A2"

    # === NOT STARTED ===
    if _include("not-started"):
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

    # === SUMBAL PATWARIS ===
    if _include("sumbal-patwaris") and views.get("sumbal_patwari_rows"):
        ws = wb.create_sheet("SUMBAL PATWARIS")
        headers = ["S.NO", "NAME OF PATWARI", "VILLAGES", "TOTAL SURVEY NOS",
                   "SUBMITTED", "% COMPLETION"]
        if additions_hdr:
            headers.append(additions_hdr)
        write_hdr(ws, headers)
        for i, row in enumerate(views["sumbal_patwari_rows"], start=1):
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
            # Sumbal Patwaris is always color-banded
            apply_band_fill(ws, r, 7 if additions_hdr else 6, row.get("band"))
        st = views.get("sumbal_patwari_totals")
        if st:
            pt = len(views["sumbal_patwari_rows"]) + 2
            ws.cell(row=pt, column=2, value="TOTAL").alignment = left
            ws.cell(row=pt, column=4, value=st["total"]).alignment = center
            ws.cell(row=pt, column=5, value=st["submitted"]).alignment = center
            ws.cell(row=pt, column=6, value=round(st["pct"], 2)).alignment = center
            if additions_hdr:
                ws.cell(row=pt, column=7, value=st["additions"] if st["additions"] is not None else "—").alignment = center
            tot_col_max = 7 if additions_hdr else 6
            for c in range(1, tot_col_max + 1):
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

    # === META (always) ===
    ws4 = wb.create_sheet("META")
    ws4.cell(row=1, column=1, value="Data as of").font = tot_font
    ws4.cell(row=1, column=2, value=to_date.strftime("%d %b %Y"))
    ws4.cell(row=2, column=1, value="Additions from").font = tot_font
    ws4.cell(row=2, column=2, value=from_date.strftime("%d %b %Y") if from_date else "—")
    ws4.column_dimensions["A"].width = 18
    ws4.column_dimensions["B"].width = 20

    # Safety: if no sheets got created (unlikely, but possible with bad view_key),
    # add an info sheet so the workbook is valid.
    if len(wb.sheetnames) == 1:  # only META
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
