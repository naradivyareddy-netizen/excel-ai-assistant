# Excel AI Assistant Pro

A professional, AI-powered Excel troubleshooting and productivity assistant. Upload any Excel or CSV file and the app detects errors, explains them in plain English, and fixes them with one click — plus generates formulas, pivot tables, dashboards, VBA macros, and reports.

Available as both a **Windows desktop app** (PyQt5) and a **web app** (Docker-ready, deployable to Hugging Face Spaces).

---

## Features

- **Formula Error Fixer** — detects and auto-repairs `#VALUE!`, `#DIV/0!`, `#N/A`, `#REF!`, `#NAME?`, `#NUM!`, `#NULL!`, circular references, and broken links.
- **Data Cleaner** — removes duplicates and blank rows, trims spaces, standardizes dates, text, and headers.
- **AI Formula Generator** — describe a calculation in plain English, get the exact Excel formula plus an explanation.
- **Pivot Table Assistant** — auto-builds grouped summary tables.
- **Dashboard Builder** — instant bar, pie, and trend charts with KPI cards.
- **VBA / Macro Generator** — produces ready-to-paste VBA code from natural-language prompts.
- **AI Chat** — ask questions about your workbook.
- **Health Check & Reports** — workbook-wide quality report with exportable output.

---

## Tech Stack

- **Language:** Python 3
- **Desktop UI:** PyQt5
- **Data:** pandas, numpy, openpyxl, xlrd
- **Charts:** matplotlib
- **Voice:** pyttsx3
- **Web:** Docker (Hugging Face Spaces compatible, port 7860)
- **Packaging:** PyInstaller (single-file `.exe`)

---

## Repository Layout

```
Excel_AI_Assistant_Pro.py    Desktop app source (single-file, auto-installs deps)
WebApp/                      Web version
  app.py                       Web server
  Dockerfile                   Container definition
  requirements.txt             Python dependencies
Sample Files/                Demo Excel files
Excel_Error_Sample.xlsx      Demo file with common errors
Excel_Error_Sample_FIXED.xlsx  Same file after auto-fix
```

---

## Run the Desktop App

```bash
python Excel_AI_Assistant_Pro.py
```

On first run the script auto-installs its dependencies. No manual `pip install` needed.

## Run the Web App (Docker)

```bash
cd WebApp
docker build -t excel-ai-assistant .
docker run -p 7860:7860 excel-ai-assistant
```

Then open http://localhost:7860.

## Deploy the Web App to Hugging Face Spaces

1. Create a new **Docker** Space at https://huggingface.co/spaces.
2. Upload the four files from `WebApp/` (`app.py`, `Dockerfile`, `requirements.txt`, `README.md`).
3. Wait for the build to finish — your app is live at `https://huggingface.co/spaces/<user>/<space>`.

---

## How to Use

1. Click **Upload Excel** and choose a `.xlsx`, `.xls`, or `.csv` file.
2. Open **Formula Fixer** → **Scan Workbook** to see every error with an explanation and suggested fix.
3. Click **Auto-Fix & Save File** to download a corrected copy.
4. Explore the other tabs: Data Cleaner, Pivot, VBA, Dashboard, AI Chat, Health Check, Reports.

Try it with the included `Excel_Error_Sample.xlsx` to see the fixer in action.

---

## Author

**Nara Divya Reddy** — [GitHub](https://github.com/naradivyareddy-netizen)
