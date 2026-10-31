# Bangladesh SSC / HSC Board Result Analyzer & Data Extractor

An automated Data Extraction & Result Analytics Web Application built with **FastAPI**, **Playwright (Headless)**, **Pandas**, **OpenPyXL**, and **Bootstrap 5** to fetch, analyze, and audit Bangladesh Secondary (SSC) and Higher Secondary (HSC) Education Board results in bulk.

---

## 🌟 Key Features

1. **Dual Extraction Modes (Bootstrap 5 UI)**:
   - **Bulk CSV / Excel Upload**: Drag-and-drop `.csv` or `.xlsx` files with `Roll` (and optional `Registration`) numbers to scrape an entire class or institution in one click.
   - **Individual Student Check**: Directly input a single student's Roll to instantly view their detailed result and marksheet.
   - **Dedicated Cumilla Fast Server**: Direct high-speed integration with `hscresult.comillaboard.gov.bd` with zero captcha and roll-only lookup.
   - Live **real-time progress bar** and dynamic results preview streamed via Server-Sent Events (SSE).

2. **Official Marksheet Card with Print Support**:
   - Generates a styled, printable marksheet displaying Student Name, Parents' names, Roll, Reg, Group, Type, Overall GPA, and individual subject grades (e.g. `101 BANGLA: A+`).

3. **Comprehensive Result Analysis Engine**:
   - **4th Subject Mode Analysis**: Toggle between **Official Board GPA (With 4th Subject Bonus)** and **Golden A+ (Without 4th Subject Pure GPA)**.
   - **Absent Student Analysis**: Tracks absent students who did not sit for the examination, displays absent rates, and provides a filterable student roster.
   - **Subject Wise Grade Analysis**: Detailed grade distribution (`A+`, `A`, `A-`, `B`, `C`, `D`, `F`), student attendance, and pass rates for each subject with visual failure alerts.
   - **Candidate Type Breakdown**: Detailed performance matrix comparing Regular fresh batch vs. Irregular/Improvement candidates.
   - **Failure Intervention Roster**: Dedicated tracking of 1-Subject Fail (supplementary eligible), 2-Subject Fail, and Multi-Subject Fail candidates.

4. **Multi-Sheet Executive Excel Workbook**:
   - Generates professional 7-sheet Excel reports with custom styling, borders, and auto-adjusted column widths:
     1. `Executive Summary`
     2. `Regular vs Irregular`
     3. `Subject Wise Grade Analysis`
     4. `1 Subject Fail`
     5. `2 Subject Fail`
     6. `3+ Subject Fail`
     7. `Absent Students`

5. **Publication-Ready Printable Audit / PDF Report**:
   - A4-optimized printable layout featuring an executive KPI dashboard, explanatory reading guide, subject grade tables, failure rosters, absent list, and official sign-off signatures.

6. **Automated Captcha Solving & Ethical Scraping**:
   - Automatically parses and solves arithmetic captchas on standard board servers.
   - Features polite request spacing and exponential retry handlers.

---

## 📁 Project Structure

```text
Board Result Analyzer/
│
├── app.py                  # FastAPI application, REST endpoints, SSE stream
├── scraper.py              # Playwright headless engine, captcha solver, table parser
├── analysis.py             # Result analytics engine & 7-sheet Excel workbook generator
├── requirements.txt        # Python library dependencies
├── sample_students.csv     # Sample CSV with Roll and Registration columns
├── run_app.bat             # 1-click Windows startup script
├── templates/
│   └── index.html          # Frontend dashboard UI (Bootstrap 5 + Icons)
├── static/
│   ├── css/
│   │   └── style.css       # Custom styles, responsive tweaks, animations
│   └── js/
│       └── main.js         # Frontend SSE handler, modal analytics, print preview
├── uploads/                # Temporary directory for uploaded files (.gitkeep)
└── exports/                # Generated Excel and CSV reports for download (.gitkeep)
```

---

## 🚀 Setup & Local Installation

### Prerequisites
- Python 3.9, 3.10, 3.11, or 3.12 installed on your system.

### Step 1: Clone / Navigate to the Project Directory
```bash
git clone <YOUR_GITHUB_REPO_URL>
cd "Board Result Analyzer"
```

### Step 2: Create and Activate a Python Virtual Environment

**On Windows (PowerShell):**
```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
```

**On Windows (Command Prompt):**
```cmd
python -m venv venv
venv\Scripts\activate.bat
```

**On Linux / macOS:**
```bash
python3 -m venv venv
source venv/bin/activate
```

### Step 3: Install Required Dependencies
```bash
pip install -r requirements.txt
```

### Step 4: Install Playwright Chromium Browser
```bash
playwright install chromium
```

---

## 🏃 Running the Application

### Option A: Using the 1-Click Batch File (Windows)
Double-click `run_app.bat` or run:
```cmd
run_app.bat
```

### Option B: Using Python / Uvicorn
```bash
python app.py
```
*(Or alternatively: `uvicorn app:app --host 127.0.0.1 --port 8000 --reload`)*

Once started, open your web browser and navigate to:
👉 **[http://localhost:8000](http://localhost:8000)**

---

## 📋 Input File Format

The uploaded file can be either `.csv` or `.xlsx`. It should contain at least:
- **`Roll`**: The 6-digit candidate roll number (e.g., `100101`).
- **`Registration`**: *(Optional for Cumilla Board)* The 10-digit candidate registration number.

---

## 📄 License
MIT License. Open source for educational and institutional research purposes.
