# Madeleine Home — Monthly Targets Dashboard

Live site showing target vs. actual revenue/units by channel and region, updated
automatically every day from the Monthly Targets and Sales Master files in Google Drive.

## What's in this repo

- `index.html` — the dashboard itself (static page, no build step)
- `sku_images.json` — SKU → product image lookup
- `data/` — one JSON file per month (`YYYY-MM.json`) plus `manifest.json` listing
  available months. The site reads these with `fetch()`.
- `scripts/build_dashboard_data.py` — pulls the Targets/Sales Master files from
  Google Drive and regenerates the `data/` files
- `.github/workflows/update-dashboard.yml` — runs the script once a day and
  commits any changes automatically
- `local/` — the local sync script (`Monthly_Target_Update_Code.py`) and a Drive
  access sanity check (`check_drive_access.py`). These run on your own machine,
  not in GitHub Actions — they're included here just for safekeeping.

## One-time setup

### 1. Create the GitHub repo

Create a new repository on GitHub (public or private — private is fine, GitHub
Pages still works on private repos as long as you have the plan tier that
supports it, or make it public if you want it visible without a repo invite).
Push everything in this `repo` folder as the contents of that repository's
root (so `index.html` sits at the repo root, not inside a subfolder).

**Do not commit the file `madeleine-targets-dashboard-00138fd3c845.json`.**
The included `.gitignore` already blocks it, but double check it's not staged
before your first commit — that key must only ever go into a GitHub Secret
(step 3 below), never into the repo itself.

### 2. Enable GitHub Pages

In the repo: **Settings → Pages → Build and deployment → Source: Deploy from a
branch → Branch: main, folder: / (root) → Save**.

GitHub will give you a URL like `https://<your-username>.github.io/<repo-name>/`.
That's the link to share with your organization.

### 3. Add repository secrets

**Settings → Secrets and variables → Actions → New repository secret.** Add
three secrets:

| Secret name | Value |
|---|---|
| `GOOGLE_SERVICE_ACCOUNT_JSON` | The entire contents of `madeleine-targets-dashboard-00138fd3c845.json`, pasted as-is |
| `MONTHLY_TARGET_EMAIL_FROM` | The Gmail address that should send error-notification emails |
| `MONTHLY_TARGET_EMAIL_APP_PASSWORD` | The Gmail App Password for that address (not your regular password) |

To get a Gmail App Password: Google Account → Security → 2-Step Verification
must be on → App passwords → generate one for "Mail".

### 4. Confirm both Drive folders are shared with the service account

The service account email is:

```
targets-dashboard-reader@madeleine-targets-dashboard.iam.gserviceaccount.com
```

Both the **Monthly Targets** folder and the **Sales Master** folder in Google
Drive need to be shared with this email (Viewer access is enough). You already
confirmed this works locally via `check_drive_access.py`.

### 5. Run the workflow once manually

**Actions tab → "Update dashboard data" → Run workflow.** This does a first
live pull from Drive and commits the result to `data/`. Watch the run — if it
fails, the error will show in the Actions log, and (once secrets are set) also
in the email step in the local script logic.

Once that run finishes successfully, open your Pages URL — the dashboard
should load with real data.

## How it stays up to date

The workflow in `.github/workflows/update-dashboard.yml` runs automatically
every day at 12:00 UTC (`cron: "0 12 * * *"`), regenerates `data/*.json` from
Drive, and commits any changes. Adjust the cron expression if you want a
different time — GitHub Actions cron is always in UTC, so convert from your
local time zone accordingly.

You can also trigger it manually any time from the Actions tab.

## Local workflow (unchanged)

`local/Monthly_Target_Update_Code.py` is the script that syncs Sales Master
data into the per-channel Monthly Targets workbooks in Drive. This keeps
running locally on your machine on its existing schedule — it's what feeds the
data the GitHub Action later reads. It's included in this repo for reference
only; it is not run by GitHub Actions.
