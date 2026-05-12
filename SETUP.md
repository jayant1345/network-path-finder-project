# Network Path Finder — Setup Guide

This guide explains how to copy this project to a new PC and get it running from scratch.

---

## Prerequisites

Install these on the new PC before anything else:

1. **Python 3.10 or later** — download from https://www.python.org/downloads/
   - During installation, check **"Add Python to PATH"**
   - Verify: open Command Prompt and run `python --version`

2. **Git** (optional, only needed if cloning from a repo) — https://git-scm.com/

---

## Step 1 — Copy the Project Files

Copy the entire project folder to the new PC. The minimum required files are:

```
network-path-finder-project/
├── server.py
├── cpan.py
├── requirements.txt
├── links.csv              ← default network data
└── templates/
    └── index.html
```

Do **not** copy the `venv/` folder — it contains compiled files tied to the original machine.

---

## Step 2 — Open a Terminal in the Project Folder

**Option A — Command Prompt / PowerShell:**
```
cd C:\path\to\network-path-finder-project
```

**Option B — File Explorer:**
Right-click inside the project folder → "Open in Terminal" (Windows 11)

---

## Step 3 — Create a Virtual Environment

Run this once on the new PC:

```cmd
python -m venv venv
```

This creates a `venv/` folder inside the project directory.

---

## Step 4 — Activate the Virtual Environment

**PowerShell:**
```powershell
venv\Scripts\Activate.ps1
```

If you see a permissions error in PowerShell, run this first:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```
Then activate again.

**Command Prompt (cmd.exe):**
```cmd
venv\Scripts\activate.bat
```

**Git Bash / WSL:**
```bash
source venv/Scripts/activate
```

After activation, your prompt will show `(venv)` at the start — this confirms it is active.

---

## Step 5 — Install Dependencies

With the virtual environment active, run:

```cmd
pip install -r requirements.txt
```

This installs: Flask, pandas, networkx, openpyxl, pywin32.

Expected output ends with something like:
```
Successfully installed flask-x.x pandas-x.x networkx-x.x ...
```

---

## Step 6 — Run the Application

```cmd
python server.py
```

You should see:
```
 * Running on http://127.0.0.1:5001
```

Open your browser and go to: **http://127.0.0.1:5001**

---

## Step 7 — Load Network Data

- The app automatically loads `links.csv` from the project folder on startup.
- To use a different file, click the **Upload** button in the web UI and select a CSV or Excel file.

### Required columns in the data file:

| Column | Example |
|---|---|
| A End | `10.121.24.167_SiteName_...` |
| Z End | `10.121.24.168_SiteName_...` |
| Bandwidth | `10GE` or `GE` |
| CIR Utilization Ratio(%) | `45.2` |
| Alarm Status | `Critical`, `Warning`, or blank |
| Name | Link name/label |
| Update Time | Timestamp |
| Bandwidth Utilization Ratio(%) | `30.5` |

---

## Daily Usage (after first-time setup)

Every time you want to use the app:

1. Open a terminal in the project folder.
2. Activate the virtual environment:
   - PowerShell: `venv\Scripts\Activate.ps1`
   - CMD: `venv\Scripts\activate.bat`
3. Run: `python server.py`
4. Open browser: http://127.0.0.1:5001

---

## Troubleshooting

| Problem | Fix |
|---|---|
| `python` not found | Python is not in PATH — reinstall Python and check "Add to PATH" |
| `pip install` fails with SSL error | Run `python -m pip install --trusted-host pypi.org --trusted-host files.pythonhosted.org -r requirements.txt` |
| Port 5001 already in use | Change the port at the bottom of `server.py` |
| `pywin32` install error | Run `pip install pywin32` separately, then `python venv\Scripts\pywin32_postinstall.py -install` |
| App loads but no data | Make sure `links.csv` is in the same folder as `server.py` |

---

## Stopping the Server

Press `Ctrl + C` in the terminal where `python server.py` is running.

To deactivate the virtual environment when done:
```
deactivate
```
