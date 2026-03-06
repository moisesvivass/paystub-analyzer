# 📊 Paystub Analyzer — Automated Earnings Report

Automatically extracts paystub data from Gmail, processes encrypted PDFs using Claude AI, and generates a professional Excel report with 7 sheets and analytics.

## 🚀 What it does

- Connects to Gmail via Google API and finds paystub emails automatically
- Downloads and decrypts password-protected PDF paystubs
- Uses Claude AI (Anthropic) to extract structured payroll data from each PDF
- Tracks already-processed emails to avoid duplicates
- Generates a professional Excel report with 7 sheets:
  - 🏠 Dashboard — key metrics and totals at a glance
  - 📋 Raw Data — full earnings history, one row per pay period
  - 📊 Annual Summary — totals by year
  - 📅 Monthly Summary — earnings over time
  - 🏢 By Company — comparison across employers
  - 💰 Deductions — tax, CPP, EI breakdown
  - 📖 Glossary — Canadian paystub terms explained

## 🛠️ Technologies Used

- Python 3
- Gmail API (Google Cloud)
- Claude AI API (Anthropic)
- PyPDF2 — PDF decryption and text extraction
- OpenPyXL — Excel generation
- OAuth 2.0 — Secure Gmail authentication
- python-dotenv — Environment variable management

## ⚙️ Setup Instructions

### 1. Clone the repository
```bash
git clone https://github.com/moisesvivass/paystub-analyzer.git
cd paystub-analyzer
```

### 2. Create and activate virtual environment
```bash
python -m venv .venv
.venv\Scripts\activate        # Windows
source .venv/bin/activate     # Mac/Linux
```

### 3. Install dependencies
```bash
pip install -r requirements.txt
```

### 4. Set up Google Cloud
- Go to [console.cloud.google.com](https://console.cloud.google.com)
- Create a new project
- Enable the Gmail API
- Create OAuth 2.0 credentials (Desktop app)
- Download the credentials JSON file

### 5. Create your `.env` file
```
ANTHROPIC_API_KEY=your_anthropic_api_key
PDF_PASSWORD=your_pdf_password
CREDENTIALS_FILE=path/to/your/client_secret.json
OUTPUT_EXCEL=path/to/output/paystubs.xlsx
EMAIL_QUERY=subject:"Pay Stub" OR subject:"Paystub" OR subject:"payslip"
```

### 6. Run the script
```bash
# Process all emails (limit to 5 for testing)
python main.py --mode full --limit 5

# Process all emails
python main.py --mode full --limit 100

# Only process new emails since last run
python main.py --mode update
```

## 📁 Project Structure
```
paystub-analyzer/
├── main.py                        # Entry point + CLI
├── requirements.txt               # Dependencies
├── .env                           # Environment variables (never committed)
├── .gitignore                     # Files excluded from GitHub
├── README.md                      # This file
├── processed_ids.json             # Tracks processed email IDs (auto-generated)
└── paystub_analyzer/
    ├── gmail_client.py            # Gmail connection + PDF download
    ├── pdf_processor.py           # PDF decryption + text extraction
    ├── claude_extractor.py        # Claude AI data extraction
    ├── excel_report.py            # Excel report generation (7 sheets)
    ├── tracker.py                 # Processed email ID tracker
    ├── config.py                  # Loads .env variables
    └── logger.py                  # Logging setup
```

## 💻 CLI Usage

```bash
# Full mode — process all emails
python main.py --mode full

# Update mode — only new emails since last run
python main.py --mode update

# Limit number of emails (recommended for testing)
python main.py --mode full --limit 10
```

## 🔒 Security

- All credentials stored in `.env` file (never committed to GitHub)
- Gmail access is read-only
- No personal data is uploaded or shared externally
- `processed_ids.json` contains only Gmail message IDs (no personal data)

## 💡 Roadmap

- [x] Modular architecture
- [x] CLI with --mode and --limit
- [x] Incremental mode — only process new paystubs
- [x] Deduplication — never double-count a pay period
- [x] Professional Excel report with 7 sheets
- [x] Generic support for any company paystubs
- [ ] Automated tests
- [ ] Scheduled automation — run automatically every month
- [ ] Web dashboard — view analytics in browser

## 👤 Author

**Moises Vivas** — Built as a personal finance + data analytics project.  
GitHub: [github.com/moisesvivass](https://github.com/moisesvivass)