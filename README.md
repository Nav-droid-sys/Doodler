# Doodler — Harman → 2026 Billing Automation

Automates the monthly billing reconciliation between the **Harman** project-delivery
dump (Mobily side) and the **2026.xlsx** billing tracker (Wipro side).

Instead of manually copy-pasting hundreds of task entries each month, this tool reads
every task from the Harman file and appends them into `2026.xlsx` in the exact same
format, structure, and business logic as the existing entries.

---

## What It Does

1. **Reads** all tasks from the `Harman Cumulative 2025` sheet of the Harman file.
2. **Skips duplicates** — any Task ID already present in `2026.xlsx` is ignored, so
   the tool only adds new months of data.
3. **Transforms each task into 3 rows** following the milestone split:
   - **UAT** = 60% of Estimator Points
   - **RFS** = 20% of Estimator Points
   - **PAC** = 20% of Estimator Points
4. **Fills the Comments + Invoice Due columns** using the milestone-timing rule:
   | Milestone | Comment            | Invoice Due        |
   |-----------|--------------------|--------------------|
   | UAT       | `UAT - {DAC}`      | DAC Month          |
   | RFS       | `RFS - {DAC + 1}`  | DAC Month + 1      |
   | PAC       | `PAC - {DAC + 3}`  | DAC Month + 3      |

   *Example:* if UAT = `Jul '26` → RFS = `Aug '26`, PAC = `Oct '26`.
5. **Writes in the 2026.xlsx format** — preserving the two colour zones
   (blue cols A–P, orange cols R–AF), fonts, alignment, and the monthly SP columns
   (Jan–Dec '26).
6. **Milestone cell (col F)** is always filled **green** (`#00B050`) for now. The
   conditional green/red colour coding is intentionally **not implemented yet** — it
   depends on the Invoice Due logic which is still being finalised.
7. **Creates a timestamped backup** of `2026.xlsx` before writing anything, and
   handles OneDrive file locks gracefully (retries, then falls back to
   `2026_UPDATED.xlsx` if the original is open).

---

## Project Structure

```
Customer reports/
├── Automation/
│   ├── harman_to_2026.py        ← MAIN script (current, use this one)
│   ├── billing_automation.py    ← earlier version (kept for reference)
│   ├── backups/                 ← auto-generated backups (not in git)
│   └── _tmp/                    ← temp working copies (not in git)
├── Dataset/                     ← customer Excel files (NOT in git)
│   ├── 2026.xlsx
│   └── Harman_Aug'26_MIT_Abhinav.xlsx
├── .gitignore
└── README.md
```

> **Note:** All Excel files (`Dataset/`, `Automation/backups/`) are excluded from
> version control via `.gitignore`. Only the source code is tracked.

---

## Requirements

- Python 3.9+
- Dependencies:
  ```bash
  pip install openpyxl python-dateutil
  ```

---

## How to Run

1. Place the latest Harman file and the `2026.xlsx` tracker inside the `Dataset/` folder.
2. From the `Automation/` folder, run:
   ```bash
   python harman_to_2026.py
   ```
3. The script will:
   - Create a backup in `Automation/backups/`
   - Read Harman + existing 2026 data
   - Append only the new tasks
   - Save the result back to `Dataset/2026.xlsx`
     (or to `2026_UPDATED.xlsx` if the original is open in Excel)

### If the file is locked by Excel / OneDrive
Close `2026.xlsx` in Excel and let OneDrive finish syncing, then re-run. If the tool
still can't overwrite, it saves the output as `2026_UPDATED.xlsx` — just rename that to
`2026.xlsx` once Excel is closed.

---

## Business Column Mapping (Harman → 2026)

| 2026.xlsx Column        | Source in Harman file                  |
|-------------------------|----------------------------------------|
| Resource / Employee     | `Current Assignee: Name`               |
| Task ID                 | `Task ID`                              |
| Estimator Points        | `Project Total Points`                 |
| Current Milestone       | `Milestone Reached` (UAT/RFS/PAC)      |
| Date onboarded          | `Project Delivered Date`               |
| Invoicing calendar (G)  | Estimator Points × milestone % (60/20/20) |
| Invoice INR (H)         | SP fraction × billing rate             |
| Comments (I)            | Milestone + month string               |
| Invoice Due (J)         | DAC-month-based timing rule            |
| Billing rate (S)        | `Cost PPP`                             |
| Monthly SP (T–AE)       | `Actual Month` → mapped to Jan–Dec     |

---

## Current Status

### Done
- File analysis and column mapping between Harman and 2026.xlsx
- Core transformation script (`harman_to_2026.py`)
- 3-row milestone split (UAT/RFS/PAC) with SP percentages
- Comments + Invoice Due month logic
- Exact 2026.xlsx formatting (zones, fonts, fills, monthly columns)
- Duplicate detection via Task ID
- Automatic backups + OneDrive lock handling

### In Progress / Planned
- **Invoice Due column** — finalising the exact timing logic (main focus)
- **Green/Red conditional colour coding** — driven by Invoice Due (after it's final)
- **Columns K, L, M** (Earned / Billed / Balance SP) — currently placeholders
- **Columns O, P** (Unbilled from Mobily dump) — currently blank
- **Desktop UI + `.exe` packaging** — autonomous weekly check with manual trigger

---

## Roadmap

1. Finalise Invoice Due logic
2. Implement conditional milestone colour coding
3. Populate Earned/Billed/Balance and Unbilled columns
4. Build a simple desktop UI (tkinter)
5. Package as a standalone `.exe` with a weekly auto-check
