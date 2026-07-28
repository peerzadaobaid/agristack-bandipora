# AgriStack Dashboard — District Bandipora

Dated-snapshot dashboard showing khasra submission progress. Reads dated Excel files from `snapshots/YYYY-MM-DD.xlsx`, compares the newest to the oldest within a 7-day rolling window, and surfaces:

- **Tehsil Wise** — count, total, submitted, % completion, change since last snapshot
- **Patwari Wise** — three sorts: by % completion, by total submissions, by recent activity
- **Not Started** — villages with zero submissions
- **Top Patwari** and **Top Tehsil** cards in the header
- **Download Excel** button — password-gated, generates a fresh xlsx of the current state

## Updating the data

Every update is a **new file**, never an overwrite.

1. Rename the day's Excel to `YYYY-MM-DD.xlsx` (e.g. `2026-08-03.xlsx`)
2. In GitHub → your repo → `snapshots/` folder → **Add file → Upload files** → drag it in
3. Commit
4. Render auto-redeploys within ~1 minute
5. Refresh the dashboard

**Naming rule:** filename must be `YYYY-MM-DD.xlsx`. Any file that doesn't match this pattern is silently ignored.

**Sheet detection:** the app prefers `Sheet2`, falls back to `MAIN`, then any sheet with the 5 required columns (Tehsil, Village, Patwari, Total Khasras/Survey, Submitted).

## The history window

- The dashboard always compares the **newest** snapshot to the **oldest snapshot within the last 7 calendar days**.
- If you upload every day, the change column will say "last 1 day", then "last 2 days", ... up to "last 7 days", then it locks at 7.
- If you skip days, the label reflects the real calendar gap. Uploading 27 July then 2 Aug gives "last 6 days" — because 2 Aug's numbers are cumulative and already contain everything since 27 July.
- Day one (only one snapshot exists) shows `—` in the change column and hides the "By Recent Activity" sub-tab.

## `AS_OF.txt` (optional override)

If you want a human-formatted "as of" line (e.g. `27 July 2026, 4:30 PM` instead of just the date), put that text in `AS_OF.txt` at the root. If the file is empty or missing, the newest snapshot's date is used.

## Password-gated Excel download

The **Download Excel** button on the dashboard is visible to everyone, but the file is only served to requests with the correct password. Non-authorized requests get a 401.

**Setting the password on Render:**

1. On the Render dashboard, open the service
2. Left sidebar → **Environment**
3. **Add Environment Variable**
4. Key: `DOWNLOAD_PASSWORD`
5. Value: whatever password you want (something long and random)
6. Save — the service will restart automatically (~30 sec)

**To download:**

Click the button on the dashboard. Browser will pop up a login prompt. Username: anything (leave blank if you like). Password: whatever you set. Browser remembers it for the session.

To change the password later, edit the env var in Render — no code change needed.

If `DOWNLOAD_PASSWORD` is not set, all download requests fail with 401. This is intentional — no accidental public downloads.

## Local run

```bash
pip install -r requirements.txt
export DOWNLOAD_PASSWORD=test123
python app.py
# open http://localhost:5000
```

## Files

| Path | Purpose |
|---|---|
| `app.py` | Flask app |
| `templates/index.html` | Dashboard page |
| `snapshots/YYYY-MM-DD.xlsx` | Dated data files (one per upload) |
| `AS_OF.txt` | Optional override for the "as of" text |
| `requirements.txt` | Python dependencies |
| `render.yaml` | Render blueprint |

## Free-tier note

Render's free web-service tier spins down after 15 min of no traffic — the first visit after a lull takes ~30 seconds. Everything after is instant. Upgrade to always-on ($7/mo) if this becomes an issue for reviewers.
