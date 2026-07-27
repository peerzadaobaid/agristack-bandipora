"""
AGRISTACK Dashboard — reads AGRISTACK.xlsx from the repo and serves
tehsil-wise, patwari-wise, and not-started views as a single dashboard.

To update the numbers: edit AGRISTACK.xlsx, commit, push. Render redeploys
automatically and the dashboard picks up the new data.
"""
import os
import re
from datetime import datetime

import pandas as pd
from flask import Flask, render_template

app = Flask(__name__)

EXCEL_PATH = os.environ.get("EXCEL_PATH", "AGRISTACK.xlsx")

# Tolerant column matchers — same logic as the browser generator
COL_MATCHERS = {
    "tehsil":    re.compile(r"tehsil", re.I),
    "village":   re.compile(r"^village$", re.I),
    "patwari":   re.compile(r"patwari", re.I),
    "khasras":   re.compile(r"khasra|survey", re.I),
    "submitted": re.compile(r"^submitted$", re.I),
}

# Simple mtime-based cache so we don't re-read the file on every request
_cache = {"mtime": None, "df": None, "sheet": None}


def _pick_sheet(xl: pd.ExcelFile) -> str:
    """Score each sheet by how many required columns it has.
    Prefer Sheet2 > MAIN > others when tied."""
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
        raise RuntimeError("No usable sheets in the Excel file.")
    return scored[0][2]


def _detect_columns(df: pd.DataFrame) -> dict:
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


def load_data():
    """Load and normalise the Excel. Cached on file mtime."""
    mtime = os.path.getmtime(EXCEL_PATH)
    if _cache["mtime"] == mtime and _cache["df"] is not None:
        return _cache["df"], _cache["sheet"]

    xl = pd.ExcelFile(EXCEL_PATH)
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

    _cache["mtime"] = mtime
    _cache["df"] = df
    _cache["sheet"] = sheet
    return df, sheet


def _pct(submitted: int, total: int) -> float:
    return (submitted / total * 100) if total > 0 else 0.0


def build_views(df: pd.DataFrame) -> dict:
    # Tehsil-wise — keep sort by absolute submitted, but add % completion
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

    # Patwari-wise — build once, produce two sorted views
    p = (df.groupby("patwari", as_index=False)
           .agg(villages_list=("village", lambda s: ", ".join(sorted(s.tolist()))),
                total=("khasras", "sum"),
                submitted=("submitted", "sum")))
    p["pct"] = p.apply(lambda r: _pct(int(r["submitted"]), int(r["total"])), axis=1)

    def _to_records(frame):
        out = []
        for _, row in frame.iterrows():
            out.append({
                "patwari": row["patwari"],
                "villages_list": row["villages_list"],
                "total": int(row["total"]),
                "submitted": int(row["submitted"]),
                "pct": float(row["pct"]),
            })
        return out

    patwari_by_pct = _to_records(
        p.sort_values(["pct", "submitted"], ascending=[False, False])
    )
    patwari_by_count = _to_records(
        p.sort_values(["submitted", "pct"], ascending=[False, False])
    )

    # Not started
    ns = df[df["submitted"] == 0].sort_values(["tehsil", "village"])
    not_started = ns.to_dict("records")

    # Top performer = highest % completion
    top = patwari_by_pct[0] if patwari_by_pct else None

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
        "top_performer": top,
        "grand": grand,
    }


def read_as_of() -> str:
    """Read the data-as-of date from AS_OF.txt if present.
    Falls back to Excel file mtime."""
    path = "AS_OF.txt"
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            text = f.read().strip()
        if text:
            return text
    mtime = datetime.fromtimestamp(os.path.getmtime(EXCEL_PATH))
    return mtime.strftime("%d %b %Y, %I:%M %p")


@app.route("/")
def index():
    df, _ = load_data()
    views = build_views(df)
    return render_template(
        "index.html",
        as_of=read_as_of(),
        **views,
    )


@app.route("/healthz")
def healthz():
    return "ok", 200


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
