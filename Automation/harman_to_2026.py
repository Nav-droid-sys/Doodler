"""
Harman -> 2026.xlsx Billing Automation
=======================================
Reads all entries from Harman Cumulative 2025 sheet and appends them
to the 'Mobily billing actuals' sheet in 2026.xlsx, matching the exact
structure, formatting and business logic of existing entries.

Business Logic:
- Each task gets 3 rows: UAT (60%), RFS (20%), PAC (20%) of Estimator Points
- Comments logic:
    UAT comment = "UAT - {DAC Month}"    | Invoice due = DAC Month
    RFS comment = "RFS - {DAC Month+1}"  | Invoice due = DAC Month+1
    PAC comment = "PAC - {DAC Month+3}"  | Invoice due = DAC Month+3
- Milestone fill color: Green (#00B050) when milestone value present
- Header Zone 1 (A-P): Light Blue (#BDD7EE)
- Header Zone 2 (R-AF): Orange/Amber (#FFC000)
- Col Q is a blank separator column

Usage:
    python harman_to_2026.py
"""

import os
import shutil
import datetime
from collections import defaultdict, OrderedDict

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────
#  FILE PATHS  (edit these if files move)
# ─────────────────────────────────────────────
BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR  = os.path.dirname(BASE_DIR)          # Customer reports folder
DATASET_DIR = os.path.join(PARENT_DIR, "Dataset")  # main data files live here

HARMAN_FILE = os.path.join(DATASET_DIR, "Harman_Aug'26_MIT_Abhinav.xlsx")
OUTPUT_FILE = os.path.join(DATASET_DIR, "2026.xlsx")
BACKUP_DIR  = os.path.join(BASE_DIR, "backups")

# ─────────────────────────────────────────────
#  CONSTANTS
# ─────────────────────────────────────────────
MONTHS_ORDER = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
]

# Column indices in 2026.xlsx (1-based)
COL_SNO          = 1   # A
COL_RESOURCE     = 2   # B
COL_DATE_ONBOARD = 3   # C
COL_TASK_ID      = 4   # D
COL_EST_POINTS   = 5   # E
COL_MILESTONE    = 6   # F
COL_INV_CALENDAR = 7   # G  (SP fraction per milestone)
COL_INVOICE_INR  = 8   # H
COL_COMMENTS     = 9   # I
COL_INVOICE_DUE  = 10  # J
COL_EARNED_SP    = 11  # K
COL_BILLED_SP    = 12  # L
COL_BALANCE_SP   = 13  # M
COL_UNBILLED_INR = 14  # N
COL_UNBILLED_PTS = 15  # O
COL_UNBILLED_AMT = 16  # P
# Q (17) = blank separator
COL_EMP_NAME     = 18  # R
COL_BILLING_RATE = 19  # S
COL_JAN          = 20  # T  (Jan '26)
COL_FEB          = 21  # U
COL_MAR          = 22  # V
COL_APR          = 23  # W
COL_MAY          = 24  # X
COL_JUN          = 25  # Y
COL_JUL          = 26  # Z
COL_AUG          = 27  # AA
COL_SEP          = 28  # AB
COL_OCT          = 29  # AC
COL_NOV          = 30  # AD
COL_DEC          = 31  # AE
COL_TOTAL        = 32  # AF

MONTH_COL_MAP = {
    "Jan": COL_JAN, "Feb": COL_FEB, "Mar": COL_MAR,
    "Apr": COL_APR, "May": COL_MAY, "Jun": COL_JUN,
    "Jul": COL_JUL, "Aug": COL_AUG, "Sep": COL_SEP,
    "Oct": COL_OCT, "Nov": COL_NOV, "Dec": COL_DEC
}

# ─────────────────────────────────────────────
#  FILL / FONT STYLES
# ─────────────────────────────────────────────
FILL_HEADER_BLUE   = PatternFill("solid", fgColor="FFBDD7EE")
FILL_HEADER_ORANGE = PatternFill("solid", fgColor="FFFFC000")
FILL_MILESTONE_GRN = PatternFill("solid", fgColor="FF00B050")
FILL_EMP_YELLOW    = PatternFill("solid", fgColor="FFFFFF00")
FILL_NONE          = PatternFill(fill_type=None)

FONT_HEADER = Font(name="Calibri", size=10, bold=True)
FONT_DATA   = Font(name="Calibri", size=10, bold=False)

ALIGN_WRAP_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_WRAP_TOP    = Alignment(horizontal="center", vertical="top",    wrap_text=True)
ALIGN_CENTER      = Alignment(horizontal="center", vertical="center")
ALIGN_DEFAULT     = Alignment()


# ─────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────
def fmt_month(dt):
    """Convert datetime to string like "Jul '26"."""
    if dt is None:
        return None
    if isinstance(dt, datetime.datetime) or isinstance(dt, datetime.date):
        return "%s '%s" % (MONTHS_ORDER[dt.month - 1], str(dt.year)[-2:])
    return str(dt)


def add_months(dt, n):
    """Add n months to a datetime, return last day of that month."""
    if dt is None:
        return None
    month = dt.month - 1 + n
    year  = dt.year + month // 12
    month = month % 12 + 1
    # last day of new month
    import calendar
    last_day = calendar.monthrange(year, month)[1]
    return datetime.datetime(year, month, last_day)


def month_str_to_col(month_str):
    """'Jul '26' -> COL_JUL (26). Returns None if not 2026."""
    if not month_str:
        return None
    parts = str(month_str).split("'")
    if len(parts) < 2:
        return None
    mon = parts[0].strip()[:3]
    yr  = parts[1].strip()
    if yr != "26":
        return None
    return MONTH_COL_MAP.get(mon)


def normalize_name(name):
    """Strip and title-case a name for matching."""
    if not name:
        return ""
    return " ".join(str(name).strip().split())


# ─────────────────────────────────────────────
#  STEP 1 – READ HARMAN DATA
# ─────────────────────────────────────────────
def read_harman(path):
    """
    Returns list of dicts, one per task row from 'Harman Cumulative 2025'.
    Only rows that have a Task ID and Assignee are included.
    """
    print("[INFO] Reading Harman file: %s" % path)
    wb  = openpyxl.load_workbook(path, data_only=True)
    ws  = wb["Harman Cumulative 2025"]
    raw = [r for r in ws.iter_rows(values_only=True) if any(c is not None for c in r)]
    wb.close()

    # Row 4 (index 3) is the real header
    tasks = []
    for row in raw[4:]:
        task_id  = row[4]
        assignee = row[3]
        if not task_id or not assignee:
            continue

        total_pts    = row[1]   # Project Total Points
        milestone    = str(row[10]).strip() if row[10] else None
        dac_raw      = row[9]   # DAC Month (datetime or "2024 PO")
        rfsit_month  = row[17]  # RFSIT Month (datetime)
        rfs_month    = row[18]  # RFS Month (datetime)
        pac_month    = row[19]  # PAC Month (datetime)
        actual_month = row[31]  # Actual Month (datetime) – for monthly SP col
        cost_ppp     = row[32]  # Cost PPP / billing rate
        deliv_date   = row[6]   # Project Delivered Date

        # Normalise DAC: if it's a datetime use it; if text like "2024 PO" treat as None
        dac_dt = dac_raw if isinstance(dac_raw, (datetime.datetime, datetime.date)) else None

        tasks.append({
            "task_id":      str(task_id).strip(),
            "assignee":     normalize_name(assignee),
            "total_pts":    float(total_pts) if total_pts else 0.0,
            "milestone":    milestone,
            "dac_dt":       dac_dt,
            "dac_raw":      dac_raw,
            "rfsit_month":  rfsit_month,
            "rfs_month":    rfs_month,
            "pac_month":    pac_month,
            "actual_month": actual_month,
            "cost_ppp":     cost_ppp,
            "deliv_date":   deliv_date,
        })

    print("[INFO] Harman tasks loaded: %d rows" % len(tasks))
    return tasks


# ─────────────────────────────────────────────
#  STEP 2 – READ EXISTING 2026 DATA
# ─────────────────────────────────────────────
def read_existing_2026(path):
    """
    Returns:
      - last_sno (int): highest SNO in sheet
      - last_data_row (int): row number of last written row
      - existing_task_ids (set): all task IDs already in the sheet
      - employee_rate_map (dict): name -> billing_rate
      - employee_onboard_map (dict): name -> onboard_date
    """
    print("[INFO] Reading existing 2026.xlsx: %s" % path)
    wb  = openpyxl.load_workbook(path, data_only=True)
    ws  = wb["Mobily billing actuals"]

    last_sno      = 0
    last_data_row = 1
    existing_ids  = set()
    emp_rate      = {}
    emp_onboard   = {}

    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        # SNO column
        if row[0] is not None:
            try:
                last_sno = int(row[0])
            except (ValueError, TypeError):
                pass
            last_data_row = i
            # record onboard date and resource name
            if row[1] and row[2]:
                emp_onboard[normalize_name(row[1])] = row[2]

        # Task ID
        if row[3]:
            existing_ids.add(str(row[3]).strip())

        # Employee name + billing rate (cols R=17, S=18 in 0-based)
        if row[17] and row[18]:
            emp_rate[normalize_name(row[17])] = row[18]

        if i > 1:
            last_data_row = max(last_data_row, i)

    wb.close()
    print("[INFO] Last SNO=%d, last data row=%d, existing tasks=%d" % (
        last_sno, last_data_row, len(existing_ids)))
    return last_sno, last_data_row, existing_ids, emp_rate, emp_onboard


# ─────────────────────────────────────────────
#  STEP 3 – GROUP HARMAN TASKS BY EMPLOYEE
# ─────────────────────────────────────────────
def group_by_employee(tasks, existing_ids):
    """
    Groups tasks by assignee. Skips tasks already in 2026.
    Returns OrderedDict: assignee_name -> [task, ...]
    """
    grouped = OrderedDict()
    skipped = 0
    for t in tasks:
        if t["task_id"] in existing_ids:
            skipped += 1
            continue
        name = t["assignee"]
        if name not in grouped:
            grouped[name] = []
        grouped[name].append(t)

    print("[INFO] New tasks to add: %d (skipped %d already existing)" % (
        sum(len(v) for v in grouped.values()), skipped))
    return grouped


# ─────────────────────────────────────────────
#  STEP 4 – COMPUTE INVOICE ROWS PER TASK
# ─────────────────────────────────────────────
def compute_invoice_rows(task):
    """
    Returns list of 3 dicts (UAT, RFS, PAC rows) with all column values.

    SP split: UAT=60%, RFS=20%, PAC=20%

    Comments & Invoice Due logic:
        UAT: comment="UAT - {DAC}"      | due = DAC month
        RFS: comment="RFS - {DAC+1}"   | due = DAC+1 month
        PAC: comment="PAC - {DAC+3}"   | due = DAC+3 month
    (If DAC is unavailable, fall back to rfsit/rfs/pac months from Harman)
    """
    pts   = task["total_pts"]
    uat_sp = round(pts * 0.60, 4)
    rfs_sp = round(pts * 0.20, 4)
    pac_sp = round(pts * 0.20, 4)

    rate   = task["cost_ppp"] or 0

    dac = task["dac_dt"]

    # Derive invoice due dates
    if dac:
        uat_due_dt = dac
        rfs_due_dt = add_months(dac, 1)
        pac_due_dt = add_months(dac, 3)
    else:
        # fallback: use what Harman provides
        uat_due_dt = task["rfsit_month"]
        rfs_due_dt = task["rfs_month"]
        pac_due_dt = task["pac_month"]

    uat_due = fmt_month(uat_due_dt)
    rfs_due = fmt_month(rfs_due_dt)
    pac_due = fmt_month(pac_due_dt)

    rows = [
        {
            "milestone":    "UAT",
            "inv_calendar": uat_sp,
            "invoice_inr":  round(uat_sp * rate, 2),
            "comments":     "UAT - %s" % uat_due if uat_due else "UAT",
            "invoice_due":  uat_due,
        },
        {
            "milestone":    "RFS",
            "inv_calendar": rfs_sp,
            "invoice_inr":  round(rfs_sp * rate, 2),
            "comments":     "RFS - %s" % rfs_due if rfs_due else "RFS",
            "invoice_due":  rfs_due,
        },
        {
            "milestone":    "PAC",
            "inv_calendar": pac_sp,
            "invoice_inr":  round(pac_sp * rate, 2),
            "comments":     "PAC - %s" % pac_due if pac_due else "PAC",
            "invoice_due":  pac_due,
        },
    ]
    return rows


# ─────────────────────────────────────────────
#  STEP 5 – APPLY CELL STYLE
# ─────────────────────────────────────────────
def style_cell(cell, value, numfmt=None, font=None, fill=None, align=None):
    cell.value = value
    if numfmt:
        cell.number_format = numfmt
    if font:
        cell.font = font
    if fill:
        cell.fill = fill
    if align:
        cell.alignment = align


# ─────────────────────────────────────────────
#  STEP 6 – WRITE ROWS TO 2026 SHEET
# ─────────────────────────────────────────────
def write_employee_block(ws, start_row, sno, employee, tasks_data,
                         emp_rate_map, emp_onboard_map):
    """
    Writes a full employee block starting at start_row.
    Each task produces 3 rows (UAT, RFS, PAC).
    Returns the next available row number after this block.
    """
    current_row  = start_row
    first_task   = tasks_data[0]
    emp_name     = employee
    rate         = emp_rate_map.get(emp_name) or first_task.get("cost_ppp") or ""
    onboard_date = emp_onboard_map.get(emp_name) or first_task.get("deliv_date")

    # Compute totals for cols K, L, M (Earned/Billed/Balance SP)
    total_pts_all = sum(t["total_pts"] for t in tasks_data)

    is_first_row_of_employee = True

    for task in tasks_data:
        invoice_rows = compute_invoice_rows(task)
        actual_mon   = task["actual_month"]
        mon_col      = month_str_to_col(fmt_month(actual_mon)) if actual_mon else None
        task_pts     = task["total_pts"]

        for ridx, inv in enumerate(invoice_rows):
            r = current_row

            # ── Zone 1: Billing data (A-P) ──────────────────────
            # A: SNO – only on very first row of entire employee block
            if is_first_row_of_employee and ridx == 0:
                c = ws.cell(row=r, column=COL_SNO, value=sno)
                c.font      = FONT_DATA
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            # B: Resource – only on first row of employee block
            if is_first_row_of_employee and ridx == 0:
                c = ws.cell(row=r, column=COL_RESOURCE, value=emp_name)
                c.font      = FONT_DATA
                c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

            # C: Date onboarded – only on first row of employee block
            if is_first_row_of_employee and ridx == 0:
                c = ws.cell(row=r, column=COL_DATE_ONBOARD, value=onboard_date)
                c.font          = FONT_DATA
                c.number_format = "d-mmm-yy"
                c.alignment     = ALIGN_CENTER

            # D: Task ID – only on UAT row (first of 3)
            if ridx == 0:
                c = ws.cell(row=r, column=COL_TASK_ID, value=task["task_id"])
                c.font      = FONT_DATA
                c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

            # E: Estimator Points – only on UAT row
            if ridx == 0:
                c = ws.cell(row=r, column=COL_EST_POINTS, value=task_pts)
                c.font      = FONT_DATA
                c.alignment = ALIGN_CENTER

            # F: Current Milestone – only on UAT row, green fill
            if ridx == 0:
                c = ws.cell(row=r, column=COL_MILESTONE, value=task["milestone"])
                c.font      = FONT_DATA
                c.fill      = FILL_MILESTONE_GRN
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            # G: Invoicing calendar (SP fraction)
            c = ws.cell(row=r, column=COL_INV_CALENDAR, value=inv["inv_calendar"])
            c.font      = FONT_DATA
            c.alignment = ALIGN_CENTER

            # H: Invoice INR
            c = ws.cell(row=r, column=COL_INVOICE_INR, value=inv["invoice_inr"])
            c.font          = FONT_DATA
            c.number_format = "0"
            c.alignment     = ALIGN_CENTER

            # I: Comments
            c = ws.cell(row=r, column=COL_COMMENTS, value=inv["comments"])
            c.font      = FONT_DATA
            c.alignment = Alignment(horizontal="left", vertical="center")

            # J: Invoice due
            c = ws.cell(row=r, column=COL_INVOICE_DUE, value=inv["invoice_due"])
            c.font      = FONT_DATA
            c.alignment = ALIGN_CENTER

            # K/L/M: Earned/Billed/Balance – only on first UAT row of first task
            if is_first_row_of_employee and ridx == 0:
                c = ws.cell(row=r, column=COL_EARNED_SP, value=total_pts_all)
                c.font      = FONT_DATA
                c.alignment = ALIGN_CENTER

                c = ws.cell(row=r, column=COL_BILLED_SP, value=0)
                c.font      = FONT_DATA
                c.alignment = ALIGN_CENTER

                c = ws.cell(row=r, column=COL_BALANCE_SP, value=total_pts_all)
                c.font      = FONT_DATA
                c.alignment = ALIGN_CENTER

            # O, P: Unbilled from Mobily dump – leave blank (to be filled later)
            # N: Unbilled INR – leave blank

            # ── Zone 2: Employee / monthly SP (R-AF) ────────────
            # R: Employee name – only UAT row of first task
            if is_first_row_of_employee and ridx == 0:
                c = ws.cell(row=r, column=COL_EMP_NAME, value=emp_name)
                c.font      = FONT_DATA
                c.fill      = FILL_EMP_YELLOW
                c.alignment = Alignment(horizontal="left", vertical="center")

            # S: Billing rate – only UAT row of first task
            if is_first_row_of_employee and ridx == 0:
                c = ws.cell(row=r, column=COL_BILLING_RATE, value=rate)
                c.font      = FONT_DATA
                c.fill      = FILL_EMP_YELLOW
                c.alignment = ALIGN_CENTER

            # Monthly SP: put task SP in the actual month column (UAT row only)
            if ridx == 0 and mon_col:
                c = ws.cell(row=r, column=mon_col, value=task_pts)
                c.font      = FONT_DATA
                c.alignment = ALIGN_CENTER

            current_row += 1

        is_first_row_of_employee = False

    # Total column (AF) for this employee block:
    # Sum all task pts in the Total column on the first row
    total_row = start_row
    c = ws.cell(row=total_row, column=COL_TOTAL, value=total_pts_all)
    c.font      = FONT_DATA
    c.alignment = ALIGN_CENTER

    return current_row  # next available row


# ─────────────────────────────────────────────
#  STEP 7 – BACKUP + MAIN WRITE
# ─────────────────────────────────────────────
def backup_file(path, backup_dir):
    os.makedirs(backup_dir, exist_ok=True)
    ts   = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    name = os.path.splitext(os.path.basename(path))[0]
    dst  = os.path.join(backup_dir, "%s_backup_%s.xlsx" % (name, ts))
    shutil.copy2(path, dst)
    print("[INFO] Backup saved: %s" % dst)
    return dst


def copy_to_local(src, tmp_dir):
    """
    Copy a OneDrive-locked file to a local temp folder so openpyxl can open it.
    Returns the local copy path.
    """
    import tempfile
    os.makedirs(tmp_dir, exist_ok=True)
    dst = os.path.join(tmp_dir, os.path.basename(src))
    shutil.copy2(src, dst)
    print("[INFO] Local copy: %s" % dst)
    return dst


def run():
    print("=" * 60)
    print("  Harman -> 2026.xlsx Billing Automation")
    print("=" * 60)

    # Local temp dir to avoid OneDrive permission locks
    TMP_DIR = os.path.join(BASE_DIR, "_tmp")

    # 1. Read Harman
    if not os.path.exists(HARMAN_FILE):
        print("[ERROR] Harman file not found: %s" % HARMAN_FILE)
        return

    harman_local = copy_to_local(HARMAN_FILE, TMP_DIR)
    tasks = read_harman(harman_local)

    # 2. Read existing 2026
    if not os.path.exists(OUTPUT_FILE):
        print("[ERROR] 2026.xlsx not found: %s" % OUTPUT_FILE)
        return

    output_local = copy_to_local(OUTPUT_FILE, TMP_DIR)
    last_sno, last_data_row, existing_ids, emp_rate_map, emp_onboard_map = \
        read_existing_2026(output_local)

    # 3. Filter + group new tasks
    grouped = group_by_employee(tasks, existing_ids)

    if not grouped:
        print("[INFO] No new tasks to add. Everything is already up to date.")
        # Clean up tmp
        shutil.rmtree(TMP_DIR, ignore_errors=True)
        return

    # 4. Backup original file
    backup_file(OUTPUT_FILE, BACKUP_DIR)

    # 5. Open LOCAL copy for writing (avoids OneDrive lock)
    print("[INFO] Opening local copy of 2026.xlsx for writing...")
    wb = openpyxl.load_workbook(output_local)
    ws = wb["Mobily billing actuals"]

    next_row = last_data_row + 1
    next_sno = last_sno + 1

    total_employees_added = 0
    total_tasks_added     = 0

    for emp_name, emp_tasks in grouped.items():
        print("[ADD] SNO=%d | %s | %d tasks" % (next_sno, emp_name, len(emp_tasks)))
        next_row = write_employee_block(
            ws, next_row, next_sno, emp_name, emp_tasks,
            emp_rate_map, emp_onboard_map
        )
        next_sno            += 1
        total_employees_added += 1
        total_tasks_added     += len(emp_tasks)

    # 6. Save local copy then move back to OneDrive
    print("[INFO] Saving updated file locally...")
    wb.save(output_local)
    wb.close()

    print("[INFO] Copying updated file back to OneDrive: %s" % OUTPUT_FILE)
    # Retry loop in case OneDrive is syncing / file is briefly locked
    max_retries = 10
    for attempt in range(1, max_retries + 1):
        try:
            shutil.copy2(output_local, OUTPUT_FILE)
            print("[INFO] File saved successfully to OneDrive.")
            break
        except PermissionError:
            if attempt == 1:
                print("[WARN] 2026.xlsx is locked (OneDrive sync or Excel open).")
                print("[WARN] Please CLOSE 2026.xlsx in Excel if it is open, then wait...")
            print("[WAIT] Retrying in 5 seconds... (attempt %d/%d)" % (attempt, max_retries))
            import time
            time.sleep(5)
    else:
        # All retries failed — save to a new file instead
        fallback = os.path.join(PARENT_DIR, "2026_UPDATED.xlsx")
        shutil.copy2(output_local, fallback)
        print("[WARN] Could not overwrite original. Saved updated file as:")
        print("[WARN] %s" % fallback)
        print("[WARN] Manually replace 2026.xlsx with this file when Excel is closed.")

    # 7. Clean up temp folder
    shutil.rmtree(TMP_DIR, ignore_errors=True)
    print("[INFO] Temp files cleaned up.")

    print("")
    print("=" * 60)
    print("  DONE!")
    print("  Employees added : %d" % total_employees_added)
    print("  Tasks added     : %d" % total_tasks_added)
    print("  Rows written    : ~%d" % (total_tasks_added * 3))
    print("  Output file     : %s" % OUTPUT_FILE)
    print("=" * 60)


if __name__ == "__main__":
    run()
