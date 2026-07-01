# ComplaintGuard 🛡️

SLA compliance analyser for IOCL vendor complaints.  
Takes the raw export from the complaint management system, figures out who was late, and calculates penalties.

Built originally as an Excel VBA macro. This is the grown-up version — runs in a browser, stores history, shows charts, and doesn't crash on Mac.

## What it does

- Upload the IOCL complaint export (.xlsx / .xls / .xlsm)
- Automatically calculates due time from SLA + assignment time
- Flags each complaint as Early / On Time / Delayed / Pending
- Applies penalty: ₹1,000 per full day of delay
- Detects auto-closed complaints (system close, not actual resolution)
- Shows charts and a filterable data table
- Export results as CSV or Excel

## Quick start

```bash
# install dependencies
pip install -r requirements.txt

# run the app
streamlit run app/main.py
```

Or with Docker:

```bash
docker compose up
```

Open http://localhost:8501 in your browser.

## Expected columns

The file needs these columns (exact names from the IOCL export):

| Column | Description |
|---|---|
| Complaint ID | Unique ID |
| Complaint Resolution Time | SLA as text: "48 hours", "24 hours" |
| Complaint DateTime | When it was logged |
| Vendor Close DateTime | When vendor closed it |
| Vendor Code | Vendor name |
| RO Code / RO Name | Dealer info |
| Vendor Remarks | Used for auto-close detection |

If your export has different column names, edit the constants at the top of `app/processor.py`.

## Project structure

```
complaintguard/
├── app/
│   ├── main.py          # streamlit dashboard
│   ├── processor.py     # core analysis engine
│   ├── database.py      # sqlite persistence
│   └── models.py        # data structures
├── tests/
│   └── test_processor.py
├── data/
│   └── sample/          # sample input files
├── requirements.txt
├── Dockerfile
├── docker-compose.yml
└── README.md
```

## Tests

```bash
python -m pytest tests/
```

## Why this exists

Before this, someone at IOCL was opening the export, sorting by date, manually checking each row, and typing penalty amounts into a spreadsheet. That takes hours and has errors. This does it in 30 seconds with a paper trail.

## Tech stack

Python, Pandas, Streamlit, Plotly, SQLite, Docker.
