# Company Report Email Automation

Automatically pulls student application alert emails from a Microsoft 365 mailbox, extracts and deduplicates student data, saves resume attachments locally (or uploads them to Google Drive), and generates one Excel report per employer.

---

## How It Works

1. **Authenticate** to Microsoft 365 using an Entra ID app registration (client-credentials / application-permission flow via the [O365](https://github.com/O365/python-o365) library).
2. **Query the Archive folder** of the configured mailbox for emails from the configured alert sender.
3. **For each alert email:**
   - Extract the employer name from the email body.
   - Parse the `student_data.xlsx` attachment into a pandas DataFrame.
   - Save any resume attachments locally with normalized `FirstName_LastName_Resume.ext` filenames.
   - Optionally upload resumes to a Google Drive folder and embed shareable links in the report.
4. **Export** one `<EmployerName>_Student_Application_Report.xlsx` file per employer into the Reports folder, with:
   - Duplicate students removed (latest application kept per email address).
   - Internal-only columns stripped.
   - Resume links rendered as clickable Excel `HYPERLINK` formulas.
   - Application date formatted as `m/d/yyyy` in the first column.

---

## Requirements

### Python Version
Python 3.8 or later.

### Required packages
```
pip install O365 pandas openpyxl
```

### Optional packages (Google Drive resume upload)
```
pip install google-api-python-client google-auth google-auth-oauthlib google-auth-httplib2
```

---

## Configuration

All configuration is done via **environment variables**. No values are hardcoded in the script.

| Variable | Required | Description |
|---|---|---|
| `GRAPH_CLIENT_ID` | Yes | Entra ID application (client) ID |
| `GRAPH_CLIENT_SECRET` | Yes | Entra ID client secret |
| `GRAPH_TENANT_ID` | Yes | Azure AD tenant ID |
| `GRAPH_WORK_EMAIL` | Yes | Mailbox address to query (e.g. `team@yourorg.com`) |
| `ALERT_SENDER_EMAIL` | Yes | Email address of the alert sender to filter on |
| `REPORTS_DIR` | No | Output folder for Excel reports. Default: `<script dir>/Reports` |
| `RESUMES_DIR` | No | Local folder for saved resume files. Default: `<script dir>/Student Resumes` |
| `RESUME_PUBLIC_FILE_BASE_URL` | No | Direct per-file base URL prepended to resume filenames in exports (preferred over `RESUME_PUBLIC_BASE_URL`). Falls back to `file://` paths if not set. |
| `RESUME_PUBLIC_BASE_URL` | No | Shared-folder URL used as resume link base when `RESUME_PUBLIC_FILE_BASE_URL` is not set. |
| `GOOGLE_DRIVE_FOLDER_ID` | No | Google Drive destination folder ID. Required if `enable_google_drive = True`. |
| `GOOGLE_DRIVE_LINK_VISIBILITY` | No | `anyone_with_link` (default) or `domain` |
| `GOOGLE_DRIVE_OAUTH_CLIENT_JSON` | No | Path to the OAuth 2.0 client secrets JSON file (or its directory). Default: script directory. |
| `GOOGLE_DRIVE_OAUTH_TOKEN_JSON` | No | Path where the cached OAuth token is stored. Default: `<script dir>/google-oauth-token.json` |

### Setting environment variables (Windows)
```powershell
$env:GRAPH_CLIENT_ID     = "your-client-id"
$env:GRAPH_CLIENT_SECRET = "your-client-secret"
$env:GRAPH_TENANT_ID     = "your-tenant-id"
$env:GRAPH_WORK_EMAIL    = "mailbox@yourorg.com"
$env:ALERT_SENDER_EMAIL  = "alerts@youralertsystem.com"
```

### Setting environment variables (macOS / Linux)
```bash
export GRAPH_CLIENT_ID="your-client-id"
export GRAPH_CLIENT_SECRET="your-client-secret"
export GRAPH_TENANT_ID="your-tenant-id"
export GRAPH_WORK_EMAIL="mailbox@yourorg.com"
export ALERT_SENDER_EMAIL="alerts@youralertsystem.com"
```

---

## Entra ID App Registration Setup

1. Go to [Azure Portal](https://portal.azure.com) → **Azure Active Directory** → **App registrations** → **New registration**.
2. Note the **Application (client) ID** and **Directory (tenant) ID**.
3. Under **Certificates & secrets**, create a new **Client secret** and note its value.
4. Under **API permissions**, add the following **Application** (not Delegated) permissions for **Microsoft Graph**:
   - `Mail.Read`
   - `Mail.ReadWrite` (only needed if the script modifies messages)
5. Click **Grant admin consent**.

---

## Google Drive Resume Upload (Optional)

Google Drive upload is **disabled by default** (`enable_google_drive = False` in the script). To enable it:

1. Set `enable_google_drive = True` in the script.
2. Create a Google Cloud project and enable the **Google Drive API**.
3. Create an **OAuth 2.0 Desktop App** client ID and download the `client_secret_*.json` file.
4. Place the JSON file in the script directory (or point `GOOGLE_DRIVE_OAUTH_CLIENT_JSON` at it).
5. Set `GOOGLE_DRIVE_FOLDER_ID` to the ID of the target Drive folder.
6. On first run, a browser window will open for you to authorize access. The token is then cached locally.

> **Note:** Google Drive service accounts cannot upload to personal My Drive due to storage quota restrictions. Use a Shared Drive (Google Workspace) or the OAuth user flow described above.

---

## Running the Script

```bash
python "Company Report Email Automation.py"
```

On completion the script prints a diagnostics summary:

```
- Messages scanned: 142
- Messages with attachments: 142
- Attachments named 'student_data.xlsx' found: 142
- Attachments named 'student_data.xlsx' processed: 142
- Attachment read failures: 0
- Employers with data prepared: 18
- Reports exported: 18
- Resume links available: 94
- Unique resumes saved: 94
- Resume mappings reused: 12
- Resume mappings updated to newer version: 3
- Google Drive links active: False
```

---

## Output Structure

```
<script dir>/
├── Reports/
│   ├── Acme_Corp_Student_Application_Report.xlsx
│   ├── Wayne_Enterprises_Student_Application_Report.xlsx
│   └── ...
└── Student Resumes/
    ├── Jane_Doe_Resume.pdf
    ├── John_Smith_Resume.docx
    └── ...
```

Each Excel report contains one row per unique student (deduplicated by email address, latest application kept), with **Date of Confirmed Application** as the first column and a clickable **Resume Link** column where resumes were found.

---

## In-Script Toggles

These flags live at the top of the script and can be adjusted without changing any other logic:

| Flag | Default | Description |
|---|---|---|
| `enable_google_drive` | `False` | Enable Google Drive resume upload |
| `clear_reports_folder_each_run` | `True` | Delete existing report files before each run |
| `clear_resumes_folder_each_run` | `True` | Delete existing resume files before each run |
| `attachment_name_expected` | `'student_data.xlsx'` | Expected filename of the student-data attachment |
| `columns_to_remove_from_exports` | *(see script)* | Column names to strip from all employer exports |
