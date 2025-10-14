# Excel Merge Tool by Huzaifa

A lightweight desktop GUI (Tkinter) application to merge multiple Excel files with the same headers, preview and paginate the merged data, save a master file with backups, and email the master file to one or more recipients. Optional Watch Folder mode monitors a folder and auto-merges when Excel files are added/changed.

---

## Files included

- `excel_merge_app.py` — Main single-file Python app (runnable).
- `requirements.txt` — Python dependencies.
- `config_template.json` — Example config file.
- `excel_merge_app.log` — (Created at runtime) log file.
- `backups/` — folder (created at runtime) to store timestamped backups.

---

## System requirements

- Python 3.9+ recommended.
- Windows / macOS / Linux desktop with GUI.
- Internet connection for emailing (if emailing).

---

## Installation (recommended)

1. Create a virtualenv (optional but recommended):
   ```bash
   python -m venv venv
   source venv/bin/activate     # macOS/Linux
   venv\Scripts\activate        # Windows
