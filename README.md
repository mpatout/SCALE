# SCALE Automation Scripts

A collection of Python automation scripts built for the **SCALE** (Scalable Asymmetric Lifecycle Engagement) program — a defense-industry-focused STEM training program that connects universities and students with U.S. defense employers.

These scripts reduce manual, repetitive data-management work by automating the most time-intensive parts of SCALE operations: onboarding new students, processing company report emails, and identifying top candidates for recognition.

---

## Scripts

### [Nanohub Student Semi-Automation Upload](Nanohub%20Student%20Semi-Automation%20Upload/)

Compares a master student list against a live Nanohub system export to find net-new students, then generates a ready-to-upload Excel import file formatted to the Nanohub schema.

**What it automates:** Identifying which students from the master list are not yet in the system, mapping them to all required Nanohub import columns, and writing a four-tab Excel file (Users, WorkExperience, Degrees, Mentoring) that can be uploaded directly.

**Key features:**
- Fuzzy matching on name + date of birth + email to avoid false duplicates
- Automatic value translation (gender, degree type, vertical/technical area)
- Academic term date parsing (e.g. `Spring '26` → `May 15, 2026`)
- Start date back-calculation from graduation date
- Suppression report in the terminal for recently-added students

**Dependencies:** `pandas`, `openpyxl`, `numpy`

---

### [Company Report Email Automation](Company%20Report%20Email%20Automation/)

Connects to a Microsoft 365 mailbox, reads company alert emails from an Archive folder, extracts student application data, saves resume attachments, and generates one Excel employer report per company — fully deduplicated and formatted.

**What it automates:** Manually opening dozens of alert emails, copy-pasting student lists, downloading resumes, and building per-employer Excel reports. The script does all of this in a single run.

**Key features:**
- Microsoft Graph API authentication via Entra ID app registration (no credentials in code — all via environment variables)
- Employer name extraction from email body
- Student deduplication per employer (latest application per email kept)
- Resume attachment saving with normalized `FirstName_LastName_Resume.ext` filenames
- Optional Google Drive resume upload with shareable link embedding
- Clickable `HYPERLINK` formulas in Excel output

**Dependencies:** `O365`, `pandas`, `openpyxl`  
**Optional (Google Drive):** `google-api-python-client`, `google-auth`, `google-auth-oauthlib`, `google-auth-httplib2`

---

### [Stories of Success](Stories%20of%20Success/)

Reads a SCALE student-database Excel export and identifies graduates who meet all four **Stories of Success** criteria — sustained SCALE participation, advanced degrees, and strong defense-industry work experience.

**What it automates:** Manually cross-referencing student records, degree history, and work experience across three data sheets to find nomination-worthy candidates.

**Key features:**
- Four-criteria scoring engine with configurable thresholds
- Flexible column and sheet name resolution (handles export naming variations automatically)
- Academic term date parsing for SCALE start and graduation dates
- Three-tab Excel output: clean Candidates list, full AllStudentsEvaluated audit trail, and Summary

**Dependencies:** `pandas`, `openpyxl`, `numpy`

---

## Requirements

All scripts require **Python 3.9 or later**. Each script folder contains its own README with dependency details and setup instructions.

To install the common dependencies shared by most scripts:

```bash
pip install pandas openpyxl numpy
```

The Company Report Email Automation script has additional requirements — see its [README](Company%20Report%20Email%20Automation/README.md).

---

## Repository Structure

```
SCALE/
├── Nanohub Student Semi-Automation Upload/
│   ├── Nanohub Student Semi-Automation Upload.py
│   └── README.md
│
├── Company Report Email Automation/
│   ├── Company Report Email Automation.py
│   └── README.md
│
├── Stories of Success/
│   ├── Stories of Success.py
│   └── README.md
│
└── README.md
```

---

## General Notes

- **No credentials or personal data are hardcoded.** All environment-specific values (paths, secrets, API keys) are either passed at runtime via prompts or set as environment variables. See each script's README for configuration details.
- Each script is self-contained in its own folder and can be run independently.
- Output files are written relative to each script's own directory unless overridden — no absolute paths to configure.
