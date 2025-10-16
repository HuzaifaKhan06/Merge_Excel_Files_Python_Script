# Excel Merge Tool README

**Author:** Muhamad Huzaifa Khan

A desktop Tkinter application that merges multiple Excel files into a single master file, previews data, saves backups, sends email (Gmail/yagmail or generic SMTP), and optionally creates/updates Google Sheets. This repository contains `excel_merge_app.py` — a single-file GUI app — and supporting files.

---

## Quick summary / TL;DR

1. Create a Python venv and install requirements: `python -m venv venv && source venv/bin/activate && pip install -r requirements.txt`.
2. Obtain `credentials.json` from Google Cloud (Desktop OAuth client), place it in the project folder.
3. (Optional) Create a Gmail "App password" if using Gmail sending and set it in Settings (do NOT commit it to GitHub).
4. Run: `python excel_merge_app.py`.

Full step-by-step instructions, screenshots and details below.

---

## Features (analyzed from `excel_merge_app.py`)

* Select multiple `.xlsx` / `.xls` files and merge them into one `pandas.DataFrame`.
* Preview first N rows (default 10) and view full merged data with pagination (configurable page size).
* Save merged master to `.xlsx` or `.csv` and automatically create timestamped backups in `backups/`.
* Send merged file by email using:

  * Gmail (via `yagmail` and a Gmail App Password) **or**
  * Generic SMTP host (host, port, TLS/SSL, username/password).
* Create a Google Sheets spreadsheet from the merged DataFrame using OAuth (Sheets API + Drive).
* Update an existing Google Sheet by providing its spreadsheet ID (requires the authenticated user to have edit access).
* Watch a folder (requires `watchdog`) to automatically merge new/modified Excel files and optionally auto-email the result.
* Keeps a session list of Google Sheets created during the running session and can open them in your default browser.
* Logging to `excel_merge_app.log` with the ability to open `backups/` folder from the UI.

---

## Requirements (suggested `requirements.txt`)

```
pandas
openpyxl
xlrd
yagmail
google-api-python-client
google-auth-httplib2
google-auth-oauthlib
watchdog
```

> Note: `tkinter` is needed by the GUI but is usually provided by your OS Python package. On Debian/Ubuntu install `python3-tk`.

## System & Python prerequisites

* Python 3.9+ (3.10/3.11 recommended)
* `tkinter` (GUI toolkit)
* Internet access for Google OAuth flows and for sending email (unless using local SMTP on the network).

### Platform notes

* **Linux (Debian/Ubuntu)**: `sudo apt install python3 python3-venv python3-pip python3-tk`.
* **Windows**: Install Python from python.org (include "Add to PATH"). `tkinter` comes with the official installer.
* **macOS**: Use the official Python installer or Homebrew; ensure `tkinter` is available (python.org builds include it).

---

## Google Cloud (Sheets + Drive API) — step-by-step

You must create OAuth credentials so the app can create/update Google Sheets for *your* Google account.

1. Go to [https://console.cloud.google.com/](https://console.cloud.google.com/)
2. Create a new Project (or select an existing one).
3. In the left menu go to **APIs & Services → Enabled APIs & services → + ENABLE APIS AND SERVICES**.
4. Find and enable both **Google Sheets API** and **Google Drive API**.
5. In **APIs & Services → OAuth consent screen**:

   * Choose **External** if this is for your account and test users (easiest).
   * App name: e.g. `Excel Merge Tool`.
   * Add an email address and any optional branding.
   * **Scopes:** you can leave default; the code uses `https://www.googleapis.com/auth/spreadsheets` and `https://www.googleapis.com/auth/drive.file` (Drive.file allows creating files the app creates).
   * Save and continue. For testing/personal use you can keep the app in testing mode; add your Google account under **Test users**.
6. In **Credentials → Create Credentials → OAuth client ID**:

   * Choose **Desktop app**.
   * Name it (e.g. `ExcelMerge Desktop`), then click Create.
   * Download the JSON and **save as** `credentials.json` into the project root (the same folder as `excel_merge_app.py`).

**Important**: keep `credentials.json` out of version control. Add it to `.gitignore`.

When you run the app and use the Google Sheets features for the first time, a browser will open and ask you to sign in and grant permissions. After the flow completes the app will save a `token.json` file — do not commit this file either.

---

## Gmail & App Password (for `yagmail` option)

> Google no longer allows plain account password sign-in for third-party apps unless you use OAuth or App Passwords with 2-Step Verification enabled.

1. Make sure the Gmail account you want to send from has **2-Step Verification** enabled.
2. Visit [https://myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords) and create an **App password** for the app (select Mail / Other → name it `ExcelMerge`).
3. Copy the 16-character app password and paste it into the app Settings (Gmail section → App password). If you don't want to store it, leave the checkbox **Save password** unchecked; the app will require it each session.

Alternatively, if you prefer not to use Gmail, configure **Generic SMTP** settings in Settings (host, port, username, password, TLS/SSL).

---

## Config file and settings

The app writes/reads `config.json` in the working directory. You can either use the Settings UI inside the app to set these values or create `config.json` manually. Example template (do NOT include secrets in the repo):

```json
{
  "use_gmail": true,
  "sender_email": "your.address@gmail.com",
  "app_password": "",
  "save_password": false,
  "smtp": {
    "host": "",
    "port": 587,
    "username": "",
    "password": "",
    "use_tls": true,
    "use_ssl": false
  },
  "backups_enabled": true,
  "watch_folder": {
    "enabled": false,
    "path": "",
    "auto_email_on_watch": false,
    "auto_email_recipients": ""
  },
  "page_size_default": 100
}
```

* `use_gmail`: `true` uses yagmail with `sender_email` + `app_password`.
* `smtp.*`: configure if `use_gmail` is false.
* `watch_folder.path`: folder path to watch; enable watch to activate.

---

## Running the app locally

1. Create and activate a virtual environment:

**Linux / macOS**

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

**Windows (PowerShell)**

```powershell
python -m venv venv
venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

2. Ensure `credentials.json` (Google OAuth) is in the same folder.
3. Run the app:

```bash
python excel_merge_app.py
```

4. The GUI opens. Use **Select Files** → pick `.xlsx` files → **Merge Now** → review preview → **Save** or **Save & Email**.

---

## Using the Google Sheets features

* **Save → Google Sheets**: prompts for a spreadsheet title and creates a new Google Sheet in your Drive.
* The first time you call this, a browser will open to complete the OAuth flow. The resulting `token.json` will be saved.
* **Update existing Google Sheet**: supply the `spreadsheetId` (the long id present in the sheet URL) — the authenticated user must have edit access.

Permissions note: the app uses installed-app OAuth credentials, so spreadsheets created belong to the signed-in user.

---

## Watch folder / Auto-email

If you enable folder watching (requires `watchdog`), the app will monitor the configured folder for new or modified `.xlsx`/`.xls` files. When a change is detected:

* The app will collect all Excel files in that folder and run a merge.
* If `auto_email_on_watch` is enabled and recipients are set, the app will save a timestamped file in `backups/` and attempt to email the merged file.

Be mindful: the watch logic uses a naive polling/wait approach and waits up to 60s for merging to complete before attempting the email — adjust/watch for large datasets.

---

## Backups & logs

* Backups are saved under `backups/` with a timestamped filename when saving or emailing (if `backups_enabled` is true).
* Runtime logs are written to `excel_merge_app.log`.

---

## Security & best practices

* **Never commit** `credentials.json`, `token.json`, `config.json` (if it contains secrets), or `backups/` to GitHub. Use `.gitignore`.
* Treat `app_password` or SMTP passwords as secrets.
* For sharing the repo, include a `config_template.json` (without secrets) to show users which keys to set.

### Suggested `.gitignore`

```
credentials.json
token.json
config.json
backups/
*.log
venv/
__pycache__/
```

---

## Packaging (optional)

You can create a standalone executable (Windows/macOS/Linux) using PyInstaller:

```bash
pip install pyinstaller
pyinstaller --onefile --name excel_merge_app excel_merge_app.py
```

Note: Tkinter GUI + resources + OAuth browser flow may require additional flags and testing per OS.

---

## Troubleshooting

* **tkinter errors on Linux**: install `python3-tk` (Debian/Ubuntu).
* **Google auth: credentials.json not found**: make sure file exists and is valid OAuth client for Desktop.
* **Gmail sending fails**: ensure 2-Step Verification is ON and an App Password was created and pasted into Settings.
* **watchdog not installed**: either `pip install watchdog` or disable watch folder.

If you see stack traces in `excel_merge_app.log`, attach them when asking for help.

---

## How to push to GitHub (brief)

1. Create a new repo on GitHub.
2. In your project folder:

```bash
git init
git add .
git commit -m "Initial commit: Excel Merge Tool"
git branch -M main
git remote add origin https://github.com/<youruser>/<repo>.git
git push -u origin main
```

**Before** committing, ensure `.gitignore` is set to exclude secrets and backups.

---

## Example `config_template.json` (drop-in for users)

```json
{
  "use_gmail": true,
  "sender_email": "",
  "app_password": "",
  "save_password": false,
  "smtp": {
    "host": "",
    "port": 587,
    "username": "",
    "password": "",
    "use_tls": true,
    "use_ssl": false
  },
  "backups_enabled": true,
  "watch_folder": {
    "enabled": false,
    "path": "",
    "auto_email_on_watch": false,
    "auto_email_recipients": ""
  },
  "page_size_default": 100
}
```

---

## License & Contribution

Include a license (MIT recommended for utilities). Add a `CONTRIBUTING.md` if you want pull requests.

---

## Contact / Support

If you want me to prepare a release build (PyInstaller) or add features (column-matching rules, scheduled merging, background service), reply here with what OS and how you'd like it distributed.

---

*End of README*
