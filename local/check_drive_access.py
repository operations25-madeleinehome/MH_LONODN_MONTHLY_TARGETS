"""
One-time sanity check: confirms the service account can see the shared
Google Drive folders before we build the full dashboard data-pull script.

Setup (run once):
    pip install google-api-python-client google-auth

Usage:
    python check_drive_access.py

Make sure the service account JSON key file (madeleine-targets-dashboard-...json)
is in the same folder as this script, or edit KEY_FILE below to point at it.
"""

from pathlib import Path

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

KEY_FILE = "madeleine-targets-dashboard-00138fd3c845.json"
SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]


def main():
    key_path = Path(KEY_FILE)
    if not key_path.exists():
        print(f"Can't find '{KEY_FILE}' next to this script. "
              f"Move the key file here or edit KEY_FILE at the top of this file.")
        return

    creds = service_account.Credentials.from_service_account_file(str(key_path), scopes=SCOPES)
    service = build("drive", "v3", credentials=creds)

    try:
        results = service.files().list(
            q="name contains 'Monthly Targets' or name contains 'Sales Master'",
            fields="files(id, name, mimeType)",
            pageSize=20,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
    except HttpError as exc:
        print(f"FAILED: Drive API returned an error: {exc}")
        return

    files = results.get("files", [])
    if not files:
        print(
            "Connected to Drive successfully, but the service account can't see "
            "'Monthly Targets' or 'Sales Master' yet. Double-check both folders "
            "were shared with:\n  targets-dashboard-reader@madeleine-targets-dashboard.iam.gserviceaccount.com\n"
            "with at least Viewer access."
        )
        return

    print(f"SUCCESS -- the service account can see {len(files)} matching item(s):\n")
    for f in files:
        kind = "folder" if f["mimeType"] == "application/vnd.google-apps.folder" else "file"
        print(f"  [{kind}] {f['name']}  (id: {f['id']})")


if __name__ == "__main__":
    main()
