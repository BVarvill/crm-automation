# CRM Automation

Automates contact data entry into a web-based CRM using Selenium WebDriver.

Reads a spreadsheet of contacts, searches the CRM by name and organisation, opens
the matching record, adds a personalised note, and schedules a follow-up call at 9am
in the contact's local timezone.

## What It Does

For each row in the input spreadsheet:

1. Logs into the CRM
2. Searches by last name, then filters results by first name + organisation
3. Handles nickname variants (e.g. Mike/Michael, Dan/Daniel)
4. Opens the matching contact record
5. Injects a pre-written note into the CRM note editor
6. Schedules a "New Call" at 9am in the contact's local timezone (converted to UK time for the CRM)

## Setup

```bash
pip install selenium pandas openpyxl
```

Firefox and [geckodriver](https://github.com/mozilla/geckodriver/releases) must be installed and on your PATH.

Copy `.env.example` to `.env` and fill in your credentials:

```bash
cp .env.example .env
```

Set the environment variables:

```bash
export CRM_URL=https://your-crm-url
export CRM_USERNAME=your_username
export CRM_PASSWORD=your_password
export EXCEL_PATH=crm_entry_sheet.xlsx
```

## Input Spreadsheet Format

The Excel file should have these columns:

| Column | Description |
|--------|-------------|
| `firstName` | Contact first name |
| `lastName` | Contact last name |
| `ParentOrg` | Organisation name |
| `Country` | Contact country (for timezone) |
| `State` | US state (for timezone) |
| `Paste_in_CRM` | Full note text to add to the contact record |

## Usage

```bash
python crm_automation.py
```

The script logs progress for each contact and skips any row with incomplete data.
