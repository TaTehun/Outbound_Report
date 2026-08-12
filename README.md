# Daily Outbound Report · Automation Pipeline

An end-to-end Python automation pipeline that downloads daily outbound report attachments from email, consolidates them into a single Excel workbook, and distributes the final report via email, all without manual intervention.

> **Visual overview:** [project-overview.html](project-overview.html)

---

## Overview

Each business day, multiple carriers and logistics partners send their outbound report files to a shared mailbox. This pipeline:

1. **Downloads** today's report attachments from the POP3 mailbox
2. **Transforms** each source file into a standardized format
3. **Consolidates** all sources into a single workbook with a refreshed pivot table
4. **Distributes** the final `.xlsb` report via SMTP, with the pivot table embedded in the email body
5. **Alerts** stakeholders if any files were not received
6. **Detects updates**: runs up to 3 times per day; if new files arrived since the last run, re-sends with an `(Updated)` subject prefix; if nothing changed, sends a "No Update" notification instead

---

## Project Structure

```
outbound_report/
├── main.py                        # Data processing, Excel consolidation & run logic
├── main.example.py                # Simplified reference version of main.py
├── project-overview.html          # Visual project overview (portfolio)
├── run_initial.bat                # 1st daily run (17:00)
├── run_second.bat                 # 2nd daily run (18:00)
├── run_third.bat                  # 3rd daily run (19:00)
├── email_pipeline/
│   ├── email_setup.py             # All settings (paths, servers, targets, recipients)
│   ├── downloader.py              # POP3 attachment downloader
│   └── sender.py                  # Email sender (report + failure + skip alerts)
└── utils/
    ├── date_utils.py              # Date formatting utilities
    └── pivot_utils.py             # Pivot refresh (xlwings) + HTML rendering
```

---

## Data Sources

| Source | Description |
|--------|-------------|
| Carrier A | IOD outbound report (Excel) |
| Carrier B | LTL daily exception list (Excel) |
| Carrier B | Milk Run daily exception list (Excel) |
| Carrier C | Daily manifest report (Excel) |
| Carrier D | Daily manifest report (CSV) |

> If a file is not received on a given day, that source is skipped and the remaining files are still consolidated.

---

## Output

All files are organized into date-stamped folders:

```
Outbound/
├── Data/
│   ├── template.xlsx              # Master template (column headers, pivot, formatting)
│   └── 05-07-2026/                # Today's downloaded attachments
└── Output/
    └── 05-07-2026/
        └── Consolidated Daily Manifest - 05.07.2026.xlsb
```

---

## How It Works

### 1. Download
The downloader connects to the POP3 server and scans incoming emails **newest-first**, stopping as soon as it reaches emails older than today. For each email, it checks the subject and sender against the configured targets before downloading the attachment, which minimizes unnecessary data transfer. Each saved file is timestamped (`filename_HHMM.ext`) so subsequent runs can detect whether new files arrived.

### 2. Transform
Each source file goes through source-specific transformations (column reordering, address splitting, date formatting, blank column insertion) to conform to a standard 19-column layout. Blank shipment status values are filled with `SHIPMENT EN-ROUTE-TO-DEST` so the pivot table shows a meaningful label instead of `(Blank)`.

### 3. Consolidate
All transformed DataFrames are concatenated and written into the `Daily Manifest` sheet of the template workbook, starting at row 2 to preserve the pre-formatted header. The pivot table is refreshed via `pivot_utils.read_pivot()` and the resulting data is captured before saving.

### 4. Distribute
The final `.xlsb` file is attached to a report email. The email body includes the refreshed pivot table rendered as an HTML table for quick in-email review. If any files were missing, a separate failure alert is sent detailing which files were not received.

### 5. Multi-Run Logic
The pipeline runs up to 3 times per day via separate bat files. Each subsequent run checks whether new files arrived since the previous run using the file timestamp (`_HHMM`) embedded in the filename:

| Run | Bat file | Condition | Action |
|-----|----------|-----------|--------|
| 1st | `run_initial.bat` | Output file doesn't exist | Consolidate and send report |
| 2nd | `run_second.bat` | Files with timestamp ≥ 17:00 exist | Re-consolidate and send `(Updated)` report |
| 2nd | `run_second.bat` | No new files | Send "No Update" notification |
| 3rd | `run_third.bat` | Files with timestamp ≥ 18:00 exist | Re-consolidate and send `(Updated)` report |
| 3rd | `run_third.bat` | No new files | Send "No Update" notification |

---

## Configuration

All settings are managed in `email_pipeline/email_setup.py`:

| Setting | Description |
|---------|-------------|
| `DATA_DIR` / `OUTPUT_DIR` | Date-stamped output folders |
| `TEMPLATE_PATH` | Path to `template.xlsx` |
| `EMAIL_TARGETS` | List of expected email sources with sender/subject criteria |
| `EMAIL_SEARCH_COUNT` | Max number of recent emails to scan |
| `REPORT_TO` / `REPORT_CC` | Report email recipients |
| `NOTIFICATION_TO` | Failure alert recipients |

Credentials (POP3/SMTP username and password) are stored in a `.env` file and loaded at runtime.

---

## Requirements

```
pandas
xlwings
openpyxl
pytz
python-dotenv
```

---

## Running

```bash
# 1st run (17:00)
python main.py initial
# or double-click run_initial.bat

# 2nd run (18:00)
python main.py second
# or double-click run_second.bat

# 3rd run (19:00)
python main.py third
# or double-click run_third.bat
```

---

## Changelog

### v1.3 · 2026-05-07 · Multi-Run Schedule & Module Split
- 3-run daily schedule with stateless update detection (`_HHMM` suffix)
- `email_pipeline` split into `downloader.py` / `sender.py` / `email_setup.py`
- Fault-tolerant source dispatch: alert and report paths fully decoupled
- Added `project-overview.html`

### v1.2 · 2026-04-XX · Email Pipeline Refactor
- Restructured `email_pipeline` into `email_setup.py`, `downloader.py`, `sender.py`

### v1.1 · 2026-04-XX · 2nd-Run Detection & Pivot Email
- 2nd-run update detection via `_HHMM` filename timestamp
- Pivot table rendered as inline HTML in email body

### v1.0 · Initial Release
- POP3 downloader, multi-source Excel consolidation, SMTP distribution
