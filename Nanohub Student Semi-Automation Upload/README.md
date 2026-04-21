# Nanohub Student Semi-Automation Upload

Compares the SCALE master student list against a current system data export and produces a ready-to-upload Excel import file. The output contains four tabs — **Users**, **WorkExperience**, **Degrees**, and **Mentoring** — formatted to match the Nanohub import template exactly.

---

## Requirements

| Requirement | Version |
|---|---|
| Python | 3.9 + |
| pandas | any recent |
| openpyxl | any recent |
| numpy | any recent |

Install dependencies (first time only):

```bash
pip install pandas openpyxl numpy
```

---

## File Overview

```
Nanohub Student Semi-Automation Upload/
├── Nanohub Student Semi-Automation Upload.py   # Main script
├── Outputs/                                    # Auto-created on first run
│   └── SCALE New Student Import MM.DD.YYYY ... .xlsx
└── README.md
```

---

## One-Time Setup

Open `Nanohub Student Semi-Automation Upload.py` in any text editor and update the two values near the top of the file:

```python
# ---------------------------------------------------------------------------
# Configuration – update these values before each run
# ---------------------------------------------------------------------------

CURRENT_SCALE_SEMESTER = "Spring 2026"   # ← change each semester

# In main():
MASTER_LIST_PATH = r"C:\path\to\your\Updated_Student_List.xlsm"  # ← your local path
```

| Variable | What to set it to |
|---|---|
| `MASTER_LIST_PATH` | Full path to the master student list `.xlsm` workbook (the file that has the **"New Student List"** tab). |
| `CURRENT_SCALE_SEMESTER` | The current semester label as it should appear in the import file, e.g. `"Fall 2026"`. Update this at the start of each semester. |

---

## How to Run

### Step 1 — Export system data from Nanohub

Log into the Nanohub admin panel and export the current student data. The file will typically be named `SCALE_Student_Data_<date>.xlsx` and must contain a tab named **`students`**.

### Step 2 — Run the script

Open a terminal, navigate to this folder, and run:

```bash
python "Nanohub Student Semi-Automation Upload.py"
```

Or on Windows you can double-click the file if Python is associated with `.py` files.

### Step 3 — Provide the export file path

The script will prompt:

```
Please provide the path to the SCALE_Student_Data Excel file.
(You can drag and drop the file into this window)

SCALE_Student_Data file path:
```

**Drag and drop** the exported file directly into the terminal window (Windows/macOS both support this), or type/paste the full path. Single or double quotes around the path are handled automatically.

### Step 4 — Collect the output

The script prints a live progress log and saves the result to:

```
Outputs/SCALE New Student Import MM.DD.YYYY HH.MM AM.xlsx
```

The `Outputs/` folder is created automatically next to the script if it does not already exist.

---

## What the Script Does

```
Master list  ──┐
               ├──► Match by Name + DOB + Email ──► New students only
System export ─┘         (normalized, case-insensitive)
                                    │
                    Filter: Current / No Mentor status only
                                    │
                    Map to Nanohub import column schema
                    (translates Vertical, Gender, Degree Type, Dates)
                                    │
                    Write Excel: Users | WorkExperience | Degrees | Mentoring
```

### Matching logic

A student is considered **already in the system** (and therefore suppressed) if **either**:

- Their normalized **First Name + Last Name + DOB** matches a system record, **or**
- Their **email address** matches any email column in the system export.

### Status filtering

Only students with a status of **Current** or **No Mentor** (including variants like `Current - No Mentor`) are included in the output.

### Suppression report

The terminal prints a **"Suppressed Recently (Last 3 Weeks)"** table showing any students who applied within the past 21 days but were suppressed because they already exist in the system. Use this to quickly catch edge cases.

---

## Output Tab Reference

| Tab | Contents |
|---|---|
| **Users** | One row per new student. All required Nanohub import columns, styled with alternating-row formatting and frozen header. |
| **WorkExperience** | Empty template with correct column headers. Fill in manually if needed before upload. |
| **Degrees** | One row per new student with degree type, major, university, and back-calculated start date. |
| **Mentoring** | Empty template with correct column headers. Fill in manually if needed before upload. |

---

## Value Translations Applied Automatically

### Vertical (Technical Area)

| Master list value | Import value |
|---|---|
| System-on-Chip (SoC) | `SoC` |
| CSME (graduate students only) | `SoC` |
| Radiation Hardening (RH) | `RH` |
| Embedded Security Systems/Trusted AI (ESS/TAI) | `TAI` |
| Heterogeneous Integration and Advanced Packaging (HI/AP) | `HI/AP` |

### Gender

| Master list value | Import value |
|---|---|
| Man | `Male` |
| Woman | `Female` |

### Degree Type

| Source value | Import value |
|---|---|
| Any Ph.D / doctoral variant | `PHD` |
| Any Master's variant | `PHD` |
| Any undergraduate / BS variant | `BS` |

### Graduation / Start Dates

Graduation date text like `Spring '26`, `Fall 2026`, `Su26` is parsed into a concrete date:

| Term | Graduation date used |
|---|---|
| Spring | May 15 |
| Summer | Aug 15 |
| Fall | Dec 15 |

Start date is back-calculated from graduation date:

| Degree | Years subtracted |
|---|---|
| BS / PHD | 4 years |
| Masters | 2 years |

---

## Troubleshooting

| Symptom | Fix |
|---|---|
| `✗ Error loading master list` | Check `MASTER_LIST_PATH` and confirm the workbook has a tab named **"New Student List"**. |
| `✗ Error loading system data` | Confirm the export file has a tab named **"students"** and is `.xlsx` or `.xls`. |
| `Permission denied` while saving output | The output file is open in Excel — close it and re-run. |
| Students missing from output unexpectedly | Check the **"Suppressed Recently"** table in the terminal. They may already exist in the system. |
| Column mapping warnings | The script auto-detects column names; if a column isn't found the field is left blank. Verify the master list column names match common patterns (e.g. "First Name", "Email", "Status"). |

---

## Updating for a New Semester

1. Update `CURRENT_SCALE_SEMESTER` in the script to the new semester string.
2. Export fresh system data from Nanohub.
3. Run the script as normal.
