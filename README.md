# 📊 Paystub Analyzer — Automated Earnings Report

Automatically extracts paystub data from Gmail, processes encrypted PDFs using AI, and generates a professional Excel report with charts and analytics.

## 🚀 What it does

- Connects to Gmail via Google API and finds paystub emails automatically
- Downloads and decrypts password-protected PDF paystubs
- Uses Claude AI (Anthropic) to extract structured data from each PDF
- Generates a professional Excel report with:
  - Full earnings history
  - Annual and monthly summaries
  - Tax breakdown (Federal, Provincial, CPP, EI)
  - Visual charts and graphs
  - Company comparisons

## 🛠️ Technologies Used

- Python 3
- Gmail API (Google Cloud)
- Claude AI API (Anthropic)
- PyPDF2 — PDF processing
- OpenPyXL — Excel generation
- OAuth 2.0 — Secure Gmail authentication

## ⚙️ Setup Instructions

### 1. Clone the repository
```
git clone https://github.com/yourusername/paystub-analyzer.git
cd paystub-analyzer
```

### 2. Install dependencies
```
pip install -r requirements.txt
```

### 3. Set up Google Cloud
- Go to console.cloud.google.com
- Create a new project
- Enable the Gmail API
- Create OAuth 2.0 credentials (Desktop app)
- Download the credentials JSON file

### 4. Set up environment variables
Create a `.env` file in the root folder:
```
ANTHROPIC_API_KEY=your_anthropic_api_key
PDF_PASSWORD=your_pdf_password
CREDENTIALS_FILE=path/to/your/client_secret.json
OUTPUT_EXCEL=path/to/output/paystubs.xlsx
```

### 5. Update email search query
In `paystubs.py`, update the Gmail search query to match your paystub email subject:
```python
q='subject:Your Paystub Email Subject Here'
```

### 6. Run the script
```
python paystubs.py
```

## 📁 Project Structure
```
paystub-analyzer/
  ├── paystubs.py          # Main script
  ├── requirements.txt     # Dependencies
  ├── .env.example         # Environment variables template
  ├── .gitignore           # Files excluded from GitHub
  ├── README.md            # This file
  └── demo/
      └── sample_report.xlsx  # Sample Excel output (fake data)
```

## 🔒 Security

- All credentials stored in `.env` file (never committed to GitHub)
- Gmail access is read-only
- No personal data is uploaded or shared

## 📊 Sample Output

See `demo/sample_report.xlsx` for an example of the generated report with anonymized data.

## 💡 Future Improvements

- [ ] Incremental mode — only process new paystubs
- [ ] Scheduling — run automatically every month
- [ ] Web dashboard — view analytics in browser
- [ ] Support for multiple employees

## 👤 Author

Built as a data analytics portfolio project.