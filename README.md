# ComplaintGuard

IOCL gets complaint exports from vendors — rows and rows of tickets with SLA times, close dates, remarks. Somebody used to open this in Excel, sort by date, check each row manually, and type penalty amounts into a spreadsheet. Hours of work, and mistakes happen.

This does the same thing in about 30 seconds.

## What it does

You upload the Excel export. The app reads the SLA ("48 hours", "24 hours"), works out when each complaint should have been resolved, compares it to when it actually was, and flags every row:

- **Early** — closed before due
- **On Time** — closed within an hour of due
- **Delayed** — late, with penalty at ₹1,000 per full day
- **Pending** — not yet closed, but shows the expected due time

It also spots auto-closed complaints (the system closed the ticket, the vendor didn't actually fix it) and flags repeat visits to the same equipment within 30 days.

The dashboard shows all this in charts — which complaint types cause the most delays, which engineers are fastest, which vendors rack up the most penalties, which retail outlets have the most issues. You can filter by date range, status, vendor, RO, or engineer. Click any row in the analysis tables to see the individual complaints behind the number.

Export the results as CSV or Excel when you're done.

## Quick start

```bash
pip install -r requirements.txt
streamlit run app/main.py
```

Or with Docker:

```bash
docker compose up
```

Open http://localhost:8501.

## Expected columns

The file comes from IOCL's system, so the column names are fixed. If your export calls them something different, you don't need to edit code anymore — just add the aliases in `config/settings.yaml`.

| Column | What it is |
|---|---|
| Complaint ID | Unique row ID |
| Complaint Resolution Time | Looks like "48 hours" or "24 hours" |
| Complaint DateTime | When the ticket came in |
| Vendor Close DateTime | When the vendor closed the ticket |
| Vendor Code | Who the vendor is |
| RO Code / RO Name | Which retail outlet |
| Vendor Remarks | Free text — used for auto-close detection |
| Engineer Name | Who worked on it |
| Nature of Complaint | What went wrong |
| DU serial No | Which equipment |
| Comp Mode | WEB or SYSTEM |

## Project structure

```
complaintguard/
├── app/
│   ├── main.py          # the dashboard
│   ├── processor.py     # all the logic
│   ├── config_loader.py # reads settings.yaml
│   ├── database.py      # upload history in sqlite
│   └── models.py        # just data classes
├── config/
│   └── settings.yaml    # penalty rate, revisit window, column aliases
├── tests/
│   ├── test_processor.py
│   └── test_config.py
├── data/
│   └── sample/          # sample file to try it
├── requirements.txt
├── Dockerfile
├── docker-compose.yml
└── README.md
```

## Settings

Everything in `config/settings.yaml`:

- `penalty.per_block` — how much per full day delay (default 1000)
- `revisit.window_days` — how many days for same-equipment repeat (default 30)
- `columns.*` — the column name aliases it tries when matching your file

Change these without touching any Python.

## Tests

```bash
python -m pytest tests/
```

## How it's different from the VBA macro

The original was an Excel VBA macro that broke on Mac because of AutoFilter and FreezePanes. This version:

- Works on any OS
- Has a proper UI with tabs and charts
- Persists upload history
- Detects column names flexibly
- Lets you configure penalty rate and revisit window
- Shows vendor and RO level breakdowns
- Compares two uploads side by side
