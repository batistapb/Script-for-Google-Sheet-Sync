# Porter Images — Google Sheets Sync (Apps Script)

Automatically syncs image request data from the **Requests** spreadsheet into the **Control** spreadsheet, writing only to the `Imagens Yunique` sheet. Prevents duplicate entries by validating the `Client-Protocol` field before inserting.

---

## Overview

This Google Apps Script reads all rows from the `Página1` sheet of the **Requests** spreadsheet and appends any new entries to the `Imagens Yunique` sheet of the **Control** spreadsheet — skipping rows that already exist (deduplication by `Client-Protocol`).

| Feature | Details |
|---|---|
| Optimized | Single batch read + single batch write per execution |
| No duplicates | Deduplication via `Client-Protocol` number |
| Non-destructive | Never modifies existing rows |
| Status untouched | `STATUS` column is always inserted blank |
| Target sheet only | Writes exclusively to `Imagens Yunique` |

---

## Spreadsheet Structure

### Source — Requests (`Página1`)

| Column | Field |
|---|---|
| A | `Protocol-Yunique` (full ID string) |
| B | `Client-Protocol` (numeric — used as dedup key) |
| C | `Requester` |
| D | `Condominium` |
| E | `Partner` |
| F | `Email` |
| G | `Occurrence-Type` |
| H | `SLA` |
| I | `Syndic` |
| J | `Requested-Camera-Location` |
| K | `Start-Date` |
| L | `End-Date` |
| M | `Request-Date` |
| N | `Request-Time` |

### Destination — Control (`Imagens Yunique`)

Same columns as above (A–N), plus:

| Column | Field |
|---|---|
| O | `STATUS` — always inserted **blank** |
| P | `Send-Date` — blank |
| Q | `Notes` — blank |
| R–T | Extra columns — blank |

---

## Setup

### 1. Open Apps Script

In the **Control** spreadsheet, go to:

```
Extensions → Apps Script
```

### 2. Paste the code

Create a new script file (or replace the default `Code.gs`) and paste the full script below.

### 3. Configure the Spreadsheet IDs

At the top of the script, update the two constants with your actual spreadsheet IDs:

```javascript
const SOURCE_SPREADSHEET_ID = 'YOUR_REQUESTS_SPREADSHEET_ID';
const TARGET_SPREADSHEET_ID = 'YOUR_CONTROL_SPREADSHEET_ID';
```

You can find the ID in the spreadsheet URL:
```
https://docs.google.com/spreadsheets/d/SPREADSHEET_ID_IS_HERE/edit
```

### 4. Run for the first time

- Select the function `syncRequestsToControl` in the toolbar
- Click **Run**
- Accept the permissions prompt (required only once)

---

## Automating with a Trigger (optional)

To run the sync automatically on a schedule:

1. In Apps Script, click **Triggers** (clock icon in the left sidebar)
2. Click **+ Add Trigger**
3. Configure:
   - Function: `syncRequestsToControl`
   - Event source: `Time-driven`
   - Type: `Hour timer` or `Day timer` (as needed)
4. Save

---

## Requirements & Permissions

- The Google account running the script must have **Editor** access to **both** spreadsheets.
- If the target sheet has **protected ranges**, the account must be added as an **exception** in the protection settings, otherwise writes will silently fail.

To check/fix protection in Google Sheets:
1. Right-click the `Imagens Yunique` tab → **Protect sheet**
2. In the panel, open the existing protection rule
3. Under **Except certain users**, add the editor's email
4. Save

---

## Troubleshooting

| Symptom | Likely Cause | Fix |
|---|---|---|
| Script runs but no rows are added | Sheet has protected ranges | Add your account as a protection exception (see above) |
| `Sheet "Página1" not found` | Sheet name mismatch | Check the exact tab name and update `SOURCE_SHEET_NAME` |
| `Sheet "Imagens Yunique" not found` | Sheet name mismatch | Check the exact tab name and update `TARGET_SHEET_NAME` |
| Dates appear in wrong format | Timezone mismatch in Apps Script | Check project timezone under **Project Settings** |
| Duplicate rows inserted | Protocol column is empty in source | Ensure column B in source always has a value |

---

## Repository Structure

```
.
├── Code.gs       # Main Apps Script sync function
└── README.md     # This file
```

