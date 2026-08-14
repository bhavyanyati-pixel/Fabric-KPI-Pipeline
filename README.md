# Fabric KPI Pipeline — GitHub Actions Setup

This repo runs `fabric_kpi_pipeline.py` automatically on a schedule using
GitHub Actions. It pulls from your source Google Sheet, calculates
feasibility, optionally syncs results back to another Google Sheet, and
emails you an HTML success/failure report each run.

The script works unchanged both locally (Windows, using your original
`D:\...` paths and Gmail credentials CSV) and in GitHub Actions (using
environment variables) — whichever is set wins.

## 0. Rotate any exposed credentials

If you ever committed or shared the actual contents of your
`Gmail Credentials.csv` or your Google service account JSON, treat those as
compromised:
- Gmail app password: https://myaccount.google.com/apppasswords — delete the old one, generate a new one.
- Service account key: in Google Cloud Console → IAM & Admin → Service Accounts → your service account → Keys → delete the old key, add a new one.

## 1. Create the GitHub repository

1. Go to https://github.com/new — name it (e.g. `fabric-kpi-pipeline`), set it **Private**, create it.
2. In the folder with these files, run:
   ```bash
   git init
   git add .
   git commit -m "Initial commit: Fabric KPI pipeline"
   git branch -M main
   git remote add origin https://github.com/<your-username>/fabric-kpi-pipeline.git
   git push -u origin main
   ```

## 2. Google Sheets access

Your service account's email (`...@project-id.iam.gserviceaccount.com`) needs:
- **Viewer** access on the source spreadsheet (`SOURCE_SPREADSHEET_ID_OR_URL`)
- **Editor** access on the sync-target spreadsheet (`SYNC_SPREADSHEET_ID`), since stage 9 writes to it — this is new compared to before, since the original script only read data.

Open the service account JSON file in a text editor — you'll paste its full contents into a secret next.

## 3. Add GitHub Secrets

Repo → **Settings → Secrets and variables → Actions → New repository secret**:

| Secret name | Value |
|---|---|
| `SERVICE_ACCOUNT_JSON` | Entire contents of your service account `.json` file |
| `SOURCE_SPREADSHEET_ID_OR_URL` | `1rLcgwVua5XXyY7lZx6G16xqmNvbwyp_BWS_y61TBssA` (or its URL) |
| `SYNC_SPREADSHEET_ID` | `1oKefmjFRJI0N_CniKjqLNzJ0C3qqbS0XBT1BP0nA31g` (or its URL) — only needed if sync-to-sheets stays enabled |
| `GMAIL_SENDER_EMAIL` | The Gmail address that sends the report |
| `GMAIL_APP_PASSWORD` | Its (new) Gmail App Password |
| `EMAIL_TO` | *(optional)* Where to send the report, if different from the sender |

## 4. The workflow file

`.github/workflows/fabric-kpi.yml`:
- Runs daily at 03:00 UTC (edit the `cron` line — use https://crontab.guru — to change it)
- Can also be triggered manually: **Actions** tab → **Fabric KPI Pipeline** → **Run workflow**
- Sets `PIPELINE_BASE_DIR=output`, so all the working folders your script normally creates under `D:\Fabric Inventory\...` are instead created under `output/` inside the runner — then uploaded as a downloadable artifact at the end of each run (the runner itself is wiped after every run)
- Sets `ENABLE_SYNC_TO_SHEETS=true` — change to `"false"` in the workflow file if you want to disable the sheet sync-back stage for automated runs

## 5. Test it

1. Push everything to GitHub.
2. **Actions** tab → **Fabric KPI Pipeline** → **Run workflow** → **Run workflow**.
3. Click into the run to watch each stage's logs live.
4. Download the **fabric-kpi-output** artifact at the bottom of the run page for the CSVs.
5. Check your inbox for the HTML success/failure email.

## Notes

- Locally on Windows, nothing changes — don't set any of these env vars and the script behaves exactly as before, reading `D:\Credentials\Gmail Credentials.csv` and your local service account file.
- If a stage fails, you get a detailed failure email and the Actions run shows ❌.
- Want the CSVs committed into the repo permanently instead of just a temporary artifact? Say the word and I'll add that step.
