# Stories of Success Candidate Builder

A Python script that evaluates SCALE program graduates and identifies students who meet the **Stories of Success** criteria — students with strong defense-industry engagement, sustained SCALE participation, and advanced academic credentials.

---

## What It Does

The script reads a SCALE student-database Excel export and scores every student against four criteria. Students who meet **all four** are written to a clean, sorted **Candidates** tab in the output report.

| # | Criterion | Default Threshold |
|---|-----------|-------------------|
| 1 | Has defense-industry experience (flag on student record) | Must be set |
| 2 | Was enrolled in SCALE for at least N years before graduating | ≥ 2.0 years |
| 3 | Earned multiple degrees including at least one MS or PhD | ≥ 2 degrees, ≥ 1 graduate |
| 4 | Completed multiple work experiences, at least one defense-related | ≥ 2 total, ≥ 1 defense |

---

## Output

A timestamped Excel file is saved to the **same folder as the script** (configurable). It contains three sheets:

| Sheet | Contents |
|-------|----------|
| **Candidates** | Students meeting all criteria, sorted strongest-first, internal flags hidden |
| **AllStudentsEvaluated** | Full scoring detail for every student in the export |
| **Summary** | Run parameters and result counts |

---

## Requirements

- Python 3.9+
- pandas
- openpyxl
- numpy

Install dependencies:

```bash
pip install pandas openpyxl numpy
```

---

## Usage

Run the script from a terminal:

```bash
python "Stories of Success.py"
```

When prompted, enter (or drag-and-drop) the path to your SCALE student-database Excel file. The file must contain three sheets:

| Sheet | Required columns |
|-------|-----------------|
| **Students** | email, first/last name, SCALE start semester, defense-experience flag |
| **Degrees** | student email, degree type, graduation date |
| **WorkExperience** | student email, defense-related flag |

Column and sheet names are matched flexibly — minor naming differences between exports are handled automatically.

---

## Adjusting Thresholds

Pass keyword arguments to `StoriesOfSuccessProcessor` to override any default threshold:

```python
from pathlib import Path
from StoriesOfSuccess import StoriesOfSuccessProcessor

processor = StoriesOfSuccessProcessor(
    database_file_path=Path("path/to/SCALE_Export.xlsx"),
    output_dir=Path("path/to/output/folder"),   # default: script folder
    min_years_in_scale=2.5,                      # default: 2.0
    min_total_work_experiences=3,                # default: 2
    min_defense_work_experiences=2,              # default: 1
)
processor.run()
```

---

## Project Structure

```
Stories of Success/
├── Stories of Success.py   # Main script
└── README.md
```

---

## Notes

- The defense-experience flag on the Students sheet is expected to be a numeric code (`20` = yes) or a text boolean (`Yes`/`No`). Both formats are handled.
- Graduation dates and SCALE start semesters may be entered as term strings (`Fall 2024`, `FA 24`, `Spring '25`) or ISO date strings — all are parsed automatically.
- If the output file is open in Excel when the script runs, it will print an error and exit without overwriting.
