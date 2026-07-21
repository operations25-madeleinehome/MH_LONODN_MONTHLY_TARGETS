# Madeleine Home — Monthly Targets Dashboard

Live site showing target vs. actual revenue/units by channel and region, updated
automatically every day from the Monthly Targets and Sales Master files in Google Drive.

## What's in this repo

- `index.html` — the dashboard itself (static page, no build step)
- `sku_images.json` — SKU → product image lookup
- `data/` — one JSON file per month (`YYYY-MM.json`) plus `manifest.json` listing
  available months. The site reads these with `fetch()`.
- `scripts/sync_sales_to_channels.py` — reads Sales Master and appends new
  sales rows into each channel's own workbook in Drive (cloud version of what
  `local/Monthly_Target_Update_Code.py` used to do on your PC)
- `scripts/build_dashboard_data.py` — reads the (now up to date) Targets and
  channel files from Drive and regenerates the `data/` files
- `.github/workflows/update-dashboard.yml` — runs both scripts once a day, in
  that order, and commits any changed `data/` files automatically
- `local/` — the original local scripts (`Monthly_Target_Update_Code.py`,
  `check_drive_access.py`). **No longer run automatically** — kept here for
  reference/backup only, now that everything runs in GitHub Actions instead.

## IMPORTANT: this repo now WRITES to your Drive files

Previously, everything here only *read* from Drive. As of the
`sync_sales_to_channels.py` script, the daily run now also *writes* new sales
rows into your channel workbooks (the same append-only, dedupe-by-PO+SKU
behavior your local script always used — it never clears or overwrites
existing rows, only adds new ones).

Because of this:

1. **The service account needs Editor access on the "Monthly Targets" folder
   only** — that's where the per-channel workbooks it writes new rows into
   live. Re-share that folder with
   `targets-dashboard-reader@madeleine-targets-dashboard.iam.gserviceaccount.com`
   and change its permission from Viewer to **Editor**. Without this, every
   write will fail with a permission error (you'll get an email about it).
   **Leave the "Sales Master - New" folder at Viewer** — the sync script only
   ever reads Sales Master (opened read-only, never saved back to), so it
   never needs write access there, and keeping it at Viewer means the
   credentials themselves can't touch Sales Master even if a bug existed.
2. **Run the workflow manually once first** (Actions tab → Run workflow) and
   open the affected files in Drive afterward to confirm the new rows look
   right, before fully trusting the daily 12 PM IST run unattended.
3. Google Drive keeps automatic version history on every file. If a run ever
   writes something wrong, open the affected file → File → Version history →
   See version history, and restore the version from just before the bad run.

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

The **Monthly Targets** folder needs to be shared with this email at
**Editor** access (the sync step writes new rows into your channel files —
see the warning above). The **Sales Master** folder only needs **Viewer**
access — it's never written to. You already confirmed read access to both
works via `check_drive_access.py`; just bump Monthly Targets up to Editor.

### 5. Run the workflow once manually

**Actions tab → "Update dashboard data" → Run workflow.** This does a first
live pull from Drive and commits the result to `data/`. Watch the run — if it
fails, the error will show in the Actions log, and (once secrets are set) also
in the email step in the local script logic.

Once that run finishes successfully, open your Pages URL — the dashboard
should load with real data.

## Channel-only targets (no SKU breakdown)

Some channels only get a single monthly revenue target rather than a per-SKU
breakdown (e.g. Walmart, Home Depot). To track these, add a sheet named
exactly **`Channel-Only Targets`** to the relevant Targets workbook (MH US or
MH EU, matching the channel's region) with these columns:

| Channel | Month | Year | Revenue Target | Units Target |
|---|---|---|---|---|
| Walmart | June | 2026 | 4000 | (optional, can leave blank) |

Add one row per channel per month as needed. The sheet is entirely optional —
if it doesn't exist, nothing changes. The dashboard will show a "channel-level
target only" badge on that channel's card so it's clear there's no per-SKU
breakdown behind the number.

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
