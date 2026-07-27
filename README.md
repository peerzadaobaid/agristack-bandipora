# AgriStack Dashboard — Bandipora

Web dashboard that reads `AGRISTACK.xlsx` from this repo and shows three views:

- **Tehsil Wise** — tehsil, villages count, total survey numbers, submitted
- **Patwari Wise** — patwari, their villages, total survey numbers, submitted
- **Not Started** — villages with 0 submissions, sorted by tehsil

## Updating the numbers

1. Edit `AGRISTACK.xlsx` (add / update the `submitted` column, patwari names, etc.)
2. Commit and push to GitHub
3. Render auto-redeploys within 1–2 minutes
4. Refresh the dashboard URL

The app prefers a sheet named `Sheet2`, falls back to `MAIN`, then any sheet that has the five required columns (Tehsil, Village, Patwari, Total Khasras/Survey, Submitted). Column matching is case-insensitive and tolerant to minor renaming.

## Local run

```bash
pip install -r requirements.txt
python app.py
# open http://localhost:5000
```

## Deploy on Render (first time)

1. Push this folder to a new GitHub repo (public or private is fine)
2. On Render dashboard → **New +** → **Blueprint**
3. Connect your GitHub, pick the repo → Render reads `render.yaml` and provisions the web service
4. Wait ~2 min for the first build. URL will be shown on the service page.

Alternative (no blueprint):

- **New +** → **Web Service** → connect repo
- Runtime: **Python 3**
- Build command: `pip install -r requirements.txt`
- Start command: `gunicorn app:app`
- Plan: Free

## Files

| File | Purpose |
|---|---|
| `app.py` | Flask app — reads Excel, aggregates, serves the dashboard |
| `templates/index.html` | Single-page dashboard with three tabs |
| `AGRISTACK.xlsx` | The source data (edit and push to update) |
| `requirements.txt` | Python dependencies |
| `render.yaml` | Render blueprint (one-click deploy config) |

## Notes

- Free tier on Render spins down after 15 min idle — the first request after that takes ~30 sec to wake up. Every subsequent request is instant.
- Data is cached in memory based on file mtime, so repeated views don't re-parse the Excel.
- No auth is built in. If the URL should be private, either add HTTP Basic Auth via a Flask decorator, or use Render's paid IP-allowlist feature.
