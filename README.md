# 📊 Paystub Analyzer — Automated ETL Pipeline

![CI](https://github.com/moisesvivass/paystub-analyzer/actions/workflows/ci.yml/badge.svg)

**The problem:** Workers with multiple employers lose track of total earnings, deductions, and tax contributions across pay periods. Manually compiling paystubs into spreadsheets is time-consuming and error-prone.

**What this solves:** An automated pipeline that answers: How much did I earn this year? Are my deductions correct? How do my earnings compare across employers?

Connects to Gmail, downloads encrypted PDF paystubs, extracts structured payroll data using Claude AI, stores everything in SQLite, and generates a professional multi-sheet Excel report, automatically, every Thursday at 6:30 AM. Already running in production with 120+ real paystubs across multiple years and employers.

> 📥 Try it with anonymized data: [demo/paystubs_DEMO.xlsx](demo/paystubs_DEMO.xlsx)

## 🚀 What it does

- Connects to Gmail via Google API and finds paystub emails using a configurable search query
- Downloads and decrypts AES-256 password-protected PDF paystubs
- Uses Claude AI (Haiku) to extract structured payroll data from raw PDF text
- Validates extracted data with Pydantic — formula injection prevention, numeric parsing, math check
- Tracks processed emails in SQLite to avoid duplicate processing across runs
- Stores all paystub data in a local SQLite database
- Generates a professional 9-sheet Excel report with year-by-year personal summaries
- Runs on a weekly schedule (every Thursday at 6:30 AM) using APScheduler

## 📊 Excel Report

The workbook generates **one personal sheet per year automatically** — no code changes needed when a new year starts or when you change jobs.

| Sheet | Description |
|-------|-------------|
| ⭐ {current year} Personal | YTD summary cards + full paystub detail |
| 📅 {previous year} Personal | Full earnings breakdown |
| … one sheet per year in DB | Auto-generated, newest first |
| 🏠 Dashboard | Lifetime totals and averages at a glance |
| 📋 Raw Data | Complete history — one row per pay period |
| 📊 Annual Summary | Totals aggregated by year |
| 📅 Monthly Summary | Earnings over time with net rate % |
| 🏢 By Company | Comparison across employers — works with multiple companies simultaneously |
| 💰 Deductions | Federal tax, provincial tax, CPP, EI breakdown |
| 📖 Glossary | Canadian paystub terms explained |

Each personal year sheet includes: Pay Periods, Gross Pay, Net Pay, Income Tax, CPP, EI, Avg Net/Period, Net Rate %, Vacation Pay, Hours Worked, and a totals row.

## 🛠️ Tech Stack

| Layer | Technology |
|-------|-----------|
| Language | Python 3 |
| AI Extraction | Claude Haiku (claude-haiku-4-5-20251001) via Anthropic API |
| Email | Gmail API (OAuth 2.0, read-only) |
| PDF | PyPDF2 + pycryptodome (AES-256 decryption) |
| Validation | Pydantic v2 |
| Database | SQLite |
| Excel | OpenPyXL |
| Scheduler | APScheduler (CronTrigger) |
| Testing | pytest + pytest-cov |
| CI/CD | GitHub Actions |
| Config | python-dotenv |

## ⚙️ Setup

### 1. Clone the repository
```bash
git clone https://github.com/moisesvivass/paystub-analyzer.git
cd paystub-analyzer
```

### 2. Create virtual environment
```bash
# Recommended: create outside OneDrive to avoid cloud-sync issues on Windows
python -m venv C:/venvs/paystub-analyzer
C:/venvs/paystub-analyzer/Scripts/activate    # Windows
source C:/venvs/paystub-analyzer/bin/activate  # Mac/Linux
```

### 3. Install dependencies
```bash
pip install -r requirements.txt
pip install -r requirements-dev.txt  # only needed to run tests
```

### 4. Set up Google Cloud
1. Go to [console.cloud.google.com](https://console.cloud.google.com)
2. Create a new project and enable the **Gmail API**
3. Create **OAuth 2.0 credentials** (Desktop app type)
4. Download the credentials JSON file

### 5. Create your `.env` file
```env
ANTHROPIC_API_KEY=your_anthropic_api_key
PDF_PASSWORD=your_pdf_password
CREDENTIALS_FILE=client_secret.json
OUTPUT_EXCEL=paystubs.xlsx
DB_FILE=paystubs.db
LOG_FILE=process.log
EMAIL_QUERY=subject:"Pay Stub" OR subject:"Paystub" OR subject:"payslip"

# Scheduler (default: every Thursday at 6:30 AM in your local timezone, auto-detected)
SCHEDULE_DAY_OF_WEEK=thu
SCHEDULE_HOUR=6
SCHEDULE_MINUTE=30
# SCHEDULE_TIMEZONE=America/Toronto  # optional — omit to use system timezone
SCHEDULE_MAX_EMAILS=10
```

### 6. Authenticate Gmail (first run only)
```bash
python main.py --mode update --limit 1
```
A browser window will open. Sign in with your Google account and grant read-only Gmail access. A `token.json` file is saved locally for future runs.

### 7. Run the pipeline
```bash
# Process new emails only (recommended for regular use)
python main.py --mode update

# Process all emails with a limit
python main.py --mode update --limit 50

# Start the weekly scheduler (runs every Thursday at 6:30 AM)
python main.py --schedule
```

## 💻 CLI Reference

| Flag | Description |
|------|-------------|
| `--mode update` | Only process emails not yet in the database |
| `--mode full` | Re-evaluate all emails (respects deduplication) |
| `--limit N` | Cap the number of emails fetched from Gmail |
| `--schedule` | Start the APScheduler weekly background job |

## 📅 Scheduler

When started with `--schedule`, the pipeline runs automatically every **Thursday at 6:30 AM** (configurable via `.env`). Run this once and leave the terminal open:

```bash
python main.py --schedule
```

Features:

- **Cross-platform lock** prevents two instances from running simultaneously
- **Run history** logs every scheduled run in the `run_history` DB table with start time, finish time, emails processed, and any errors
- **Misfire tolerance** if the machine is off at 6:30 AM, the job fires within 1 hour of coming back online
- **Timezone** auto-detected from your system, no configuration needed. Override with `SCHEDULE_TIMEZONE` in `.env` if needed

## 🗄️ Database

All data is stored in `paystubs.db` (SQLite). Three tables:

**`paystubs`** — one row per pay period  
**`processed_emails`** — tracks which Gmail message IDs have been handled  
**`run_history`** — log of every scheduled pipeline run

```sql
-- Total earnings by year
SELECT strftime('%Y', pay_period_end) AS year,
       COUNT(*) AS periods,
       SUM(gross_pay) AS total_gross,
       SUM(net_pay) AS total_net
FROM paystubs
GROUP BY year;

-- Earnings by company
SELECT company, COUNT(*) AS periods, SUM(gross_pay) AS total
FROM paystubs
GROUP BY company;
```

## 🔒 Security

- All secrets in `.env` — never committed to GitHub
- Gmail access is **read-only** (no send, no delete)
- Formula injection prevention on all text fields extracted from PDFs
- `*.xlsx`, `*.db`, `token.json`, `client_secret.json` all excluded via `.gitignore`
- Parameterized SQL queries throughout — no string interpolation

## 🧪 Tests

```bash
pytest tests/ -v
pytest tests/ --cov=paystub_analyzer --cov-report=term-missing
```

28 tests across 6 files. No real API calls — all external dependencies are mocked.

| File | What it tests |
|------|--------------|
| `test_models.py` | Pydantic validation, formula injection, numeric parsing |
| `test_database.py` | Insert, deduplication, run_history lifecycle |
| `test_excel_report.py` | Dedup logic, sorting, file generation |
| `test_tracker.py` | DB-backed email tracking |
| `test_claude_extractor.py` | Retry logic, JSON parsing, API error handling |
| `test_pdf_processor.py` | PDF decryption and text extraction |

## 📁 Project Structure

```
paystub-analyzer/
├── main.py                         # Entry point + CLI + run_pipeline()
├── requirements.txt                # Runtime dependencies
├── requirements-dev.txt            # Dev/test dependencies
├── .env                            # Secrets (never committed)
├── .gitignore
├── README.md
├── demo/
│   └── paystubs_DEMO.xlsx          # Anonymized sample report
├── docs/images/                    # Screenshots for README
├── .github/workflows/ci.yml        # GitHub Actions CI
├── tests/
│   ├── test_claude_extractor.py
│   ├── test_database.py
│   ├── test_excel_report.py
│   ├── test_models.py
│   ├── test_pdf_processor.py
│   └── test_tracker.py
└── paystub_analyzer/
    ├── config.py                   # Loads .env, validate_config()
    ├── logger.py                   # RotatingFileHandler (10 MB × 5)
    ├── models.py                   # Pydantic PaystubData model
    ├── gmail_client.py             # Gmail OAuth + paginated search
    ├── pdf_processor.py            # AES-256 PDF decrypt + text extract
    ├── claude_extractor.py         # Claude AI extraction + retry
    ├── database.py                 # SQLite layer (context manager)
    ├── tracker.py                  # Processed email ID tracking
    ├── excel_report.py             # 9-sheet Excel report builder
    └── scheduler.py               # APScheduler weekly job
```

## ✅ Roadmap

- ✅ Modular architecture with clean separation of concerns
- ✅ CLI with `--mode` and `--limit` flags
- ✅ Incremental processing — only new emails per run
- ✅ Deduplication — never double-counts a paystub
- ✅ Pydantic validation + math consistency check
- ✅ AES-256 PDF decryption (pycryptodome)
- ✅ SQLite database with run history
- ✅ Dynamic Excel — one personal sheet per year, auto-generated, no code changes needed
- ✅ Multi-employer support — works across job changes simultaneously
- ✅ Weekly APScheduler with cross-platform lock and run history
- ✅ Rotating log file (10 MB × 5 backups)
- ✅ 28 automated tests (pytest)
- ✅ GitHub Actions CI/CD
- ⬜ Web dashboard — view analytics in browser (FastAPI + React)

## 👨‍💻 Author

**Moises Vivas** — AI Application Developer · Python · FastAPI · React · PostgreSQL · Claude API · Toronto, Canada

- GitHub: [github.com/moisesvivass](https://github.com/moisesvivass)
- LinkedIn: [linkedin.com/in/moisesvivas](https://linkedin.com/in/moisesvivas)
