# =============================================================================
# Outlook Reports — Email Workflow Automation
# =============================================================================
#
# DEPENDENCIES
# ------------
# Required:
#   pip install O365 pandas openpyxl
#
# Optional (Google Drive resume upload):
#   pip install google-api-python-client google-auth google-auth-oauthlib google-auth-httplib2
#
# WORKFLOW
# --------
# 1. Authenticate to Microsoft 365 via an Entra ID app registration
#    (application permission / client-credentials flow).
# 2. Query the mailbox Archive folder for alert emails from a configured sender.
# 3. For each alert email:
#    a. Extract the employer name from the email body.
#    b. Parse the student-data Excel attachment (student_data.xlsx by default).
#    c. Save any resume attachments locally with normalized per-student filenames.
#    d. Optionally upload resumes to Google Drive and store shareable links.
#    e. Append the resulting DataFrame to an in-memory per-employer bucket.
# 4. For each employer bucket: concatenate DataFrames, deduplicate by student email
#    (keeping the latest application date), strip internal-only columns, and write
#    one Excel report file per employer.
#
# CONFIGURATION — environment variables (all required unless marked optional)
# --------------------------------------------------------------------------
# GRAPH_CLIENT_ID               Entra ID application (client) ID
# GRAPH_CLIENT_SECRET           Entra ID client secret
# GRAPH_TENANT_ID               Azure AD tenant ID
# GRAPH_WORK_EMAIL              Mailbox address to query (e.g. team@yourorg.com)
# ALERT_SENDER_EMAIL            Email address of the alert sender to filter on
# REPORTS_DIR                   (optional) Output folder for Excel reports
#                               Default: <script dir>/Reports
# RESUMES_DIR                   (optional) Local folder for saved resume files
#                               Default: <script dir>/Student Resumes
# RESUME_PUBLIC_BASE_URL        (optional) Shared-folder URL used as resume link base
# RESUME_PUBLIC_FILE_BASE_URL   (optional) Direct per-file base URL for resume links
#                               Preferred over RESUME_PUBLIC_BASE_URL when set
# GOOGLE_DRIVE_FOLDER_ID        (optional) Google Drive destination folder ID
#                               Required when enable_google_drive = True
# GOOGLE_DRIVE_LINK_VISIBILITY  (optional) 'anyone_with_link' or 'domain'
#                               Default: anyone_with_link
# GOOGLE_DRIVE_OAUTH_CLIENT_JSON (optional) Path to OAuth 2.0 client JSON or its directory
# GOOGLE_DRIVE_OAUTH_TOKEN_JSON  (optional) Path where the cached OAuth token is stored
#                               Default: <script dir>/google-oauth-token.json
# =============================================================================

from O365 import Account
import pandas as pd
import re
import os
import urllib.parse
import json
import mimetypes

try:
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
    from googleapiclient.http import MediaFileUpload
    google_drive_core_libs_available = True
except ImportError:
    build = None
    HttpError = Exception
    MediaFileUpload = None
    google_drive_core_libs_available = False

try:
    from google.oauth2.credentials import Credentials as UserCredentials
    from google.auth.transport.requests import Request
    from google_auth_oauthlib.flow import InstalledAppFlow
    google_drive_oauth_libs_available = True
except ImportError:
    UserCredentials = None
    Request = None
    InstalledAppFlow = None
    google_drive_oauth_libs_available = False

# --- Section 1: Configuration ---
# All credentials and identifiers are read from environment variables.
# See the header block above for the full list of supported variables.
client_id = os.getenv('GRAPH_CLIENT_ID', '').strip()
client_secret = os.getenv('GRAPH_CLIENT_SECRET', '').strip()
tenant_id = os.getenv('GRAPH_TENANT_ID', '').strip()
work_email = os.getenv('GRAPH_WORK_EMAIL', '').strip()
alert_sender_email = os.getenv('ALERT_SENDER_EMAIL', '').strip()

credentials = (client_id, client_secret)
_script_dir = os.path.dirname(os.path.abspath(__file__))
reports_dir = os.getenv('REPORTS_DIR', os.path.join(_script_dir, 'Reports'))
resumes_dir = os.getenv('RESUMES_DIR', os.path.join(_script_dir, 'Student Resumes'))
attachment_name_expected = 'student_data.xlsx'
resume_extensions = {'.pdf', '.doc', '.docx'}
clear_resumes_folder_each_run = True
clear_reports_folder_each_run = True
enable_google_drive = False  # Set to True to upload resumes to Google Drive; False for local-only saves (default)

columns_to_remove_from_exports = [
    'Microelectonics related?',
    'Microelectronics related?',
    'Microelectronics Related?',
    'Security clearance',
    'Security Clearance',
    'Citizenship',
    'SCALE Version',
    'Recruitment Ready Points',
    'Development Ready Points',
    'Knowledge Ready Points',
    'Experience Ready Points',
    'Total Percentage',
]

columns_to_remove_from_exports_normalized = {
    re.sub(r'\s+', ' ', str(col).strip()).lower()
    for col in columns_to_remove_from_exports
}

# Optional: public-facing URL base for resume links embedded in exports.
# Set RESUME_PUBLIC_FILE_BASE_URL for a direct per-file base URL (preferred).
# Set RESUME_PUBLIC_BASE_URL for a short shared-folder link (/:f:/ style).
# If neither is set, resume links fall back to local file:// paths.
resume_public_base_url = os.getenv('RESUME_PUBLIC_BASE_URL', '').strip()
resume_public_file_base_url = os.getenv('RESUME_PUBLIC_FILE_BASE_URL', '').strip()

# Google Drive configuration (OAuth user mode)
google_drive_folder_id = os.getenv('GOOGLE_DRIVE_FOLDER_ID', '').strip()
google_drive_link_visibility = os.getenv('GOOGLE_DRIVE_LINK_VISIBILITY', 'anyone_with_link').strip().lower()
google_drive_oauth_client_path_or_dir = os.getenv('GOOGLE_DRIVE_OAUTH_CLIENT_JSON', _script_dir).strip()
google_drive_oauth_token_path = os.getenv(
    'GOOGLE_DRIVE_OAUTH_TOKEN_JSON',
    os.path.join(_script_dir, 'google-oauth-token.json'),
).strip()
google_drive_scopes = ['https://www.googleapis.com/auth/drive']


def is_oauth_client_json(path):
    try:
        with open(path, 'r', encoding='utf-8') as f:
            payload = json.load(f)
        installed = payload.get('installed', {})
        return bool(installed.get('client_id') and installed.get('client_secret'))
    except Exception:
        return False


def discover_oauth_client_json(path_or_dir):
    candidates = []

    if path_or_dir:
        if os.path.isfile(path_or_dir):
            candidates.append(path_or_dir)
        elif os.path.isdir(path_or_dir):
            for name in sorted(os.listdir(path_or_dir)):
                if name.lower().endswith('.json'):
                    candidates.append(os.path.join(path_or_dir, name))

    # Deterministic fallback: project directory only.
    project_dir = os.path.dirname(os.path.abspath(__file__))
    if os.path.isdir(project_dir):
        for name in sorted(os.listdir(project_dir)):
            if name.lower().endswith('.json'):
                candidates.append(os.path.join(project_dir, name))

    seen = set()
    for candidate in candidates:
        normalized_path = os.path.normpath(candidate)
        if normalized_path in seen:
            continue
        seen.add(normalized_path)
        if is_oauth_client_json(normalized_path):
            return normalized_path
    return ''


def get_google_drive_folder_meta(service):
    return service.files().get(
        fileId=google_drive_folder_id,
        supportsAllDrives=True,
        fields='id,name,mimeType,driveId',
    ).execute()

def initialize_google_drive_oauth_user():
    if not google_drive_core_libs_available:
        return None, '', '', 'Install google-api-python-client and google-auth.'
    if not google_drive_oauth_libs_available:
        return None, '', '', 'Install google-auth-oauthlib for OAuth user auth.'

    oauth_client_path = discover_oauth_client_json(google_drive_oauth_client_path_or_dir)
    if not oauth_client_path:
        return None, '', '', 'OAuth client JSON not found.'

    creds = None

    def refresh_error_requires_reauth(exc):
        text = str(exc).lower()
        oauth_reauth_markers = [
            'disabled_client',
            'invalid_grant',
            'invalid_client',
            'account not found',
            'token has been expired or revoked',
        ]
        return any(marker in text for marker in oauth_reauth_markers)

    try:
        if os.path.exists(google_drive_oauth_token_path):
            creds = UserCredentials.from_authorized_user_file(
                google_drive_oauth_token_path,
                google_drive_scopes,
            )

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                try:
                    creds.refresh(Request())
                except Exception as refresh_exc:
                    if refresh_error_requires_reauth(refresh_exc):
                        creds = None
                    else:
                        raise
            else:
                creds = None

            if not creds or not creds.valid:
                flow = InstalledAppFlow.from_client_secrets_file(
                    oauth_client_path,
                    google_drive_scopes,
                )
                creds = flow.run_local_server(port=0)

            with open(google_drive_oauth_token_path, 'w', encoding='utf-8') as token_file:
                token_file.write(creds.to_json())

        service = build('drive', 'v3', credentials=creds, cache_discovery=False)
        folder_meta = get_google_drive_folder_meta(service)
        if folder_meta.get('mimeType') != 'application/vnd.google-apps.folder':
            return None, oauth_client_path, google_drive_oauth_token_path, 'Configured folder ID is not a folder.'

        return service, oauth_client_path, google_drive_oauth_token_path, ''
    except Exception as exc:
        return None, oauth_client_path, google_drive_oauth_token_path, f'OAuth user auth failed: {exc}'


def initialize_google_drive_service():
    if not google_drive_folder_id:
        return None, '', '', ''

    service, oauth_client_path, oauth_token_path, reason = initialize_google_drive_oauth_user()
    if service is not None:
        return service, 'oauth_user', oauth_client_path, oauth_token_path

    print(f'Google Drive disabled: {reason} Using existing resume link behavior.')
    return None, '', '', ''


def drive_query_escape(value):
    return str(value).replace('\\', '\\\\').replace("'", "\\'")


def find_drive_file_by_name(service, folder_id, file_name):
    escaped_name = drive_query_escape(file_name)
    query = (
        f"name = '{escaped_name}' and "
        f"'{folder_id}' in parents and trashed = false"
    )
    response = service.files().list(
        q=query,
        fields='files(id,name,webViewLink,modifiedTime)',
        orderBy='modifiedTime desc',
        pageSize=5,
        includeItemsFromAllDrives=True,
        supportsAllDrives=True,
    ).execute()
    files = response.get('files', [])
    return files[0] if files else None


def parse_google_http_error(exc):
    status_code = None
    if hasattr(exc, 'status_code'):
        status_code = exc.status_code
    elif hasattr(exc, 'resp') and hasattr(exc.resp, 'status'):
        status_code = exc.resp.status

    reason = ''
    details = getattr(exc, 'error_details', None)
    if isinstance(details, list):
        for item in details:
            if isinstance(item, dict) and item.get('reason'):
                reason = str(item.get('reason'))
                break

    if not reason and 'responsepreparationfailure' in str(exc).lower():
        reason = 'responsePreparationFailure'

    return status_code, reason


def is_recoverable_response_preparation_failure(exc):
    status_code, reason = parse_google_http_error(exc)
    if reason == 'responsePreparationFailure':
        return True
    return status_code == 500 and 'responsepreparationfailure' in str(exc).lower()


def get_drive_file_links(service, file_id):
    metadata = service.files().get(
        fileId=file_id,
        fields='id,webViewLink,webContentLink',
        supportsAllDrives=True,
    ).execute()
    return metadata.get('webViewLink', ''), metadata.get('webContentLink', '')


def ensure_drive_view_permission(service, file_id, visibility, permission_cache):
    if visibility != 'anyone_with_link':
        return
    if file_id in permission_cache:
        return
    try:
        service.permissions().create(
            fileId=file_id,
            body={'type': 'anyone', 'role': 'reader'},
            fields='id',
            supportsAllDrives=True,
        ).execute()
    except HttpError as exc:
        # Some organisations restrict public sharing; continue and retain whatever link was returned.
        print(f'Could not set public link permission for Google Drive file {file_id}: {exc}')
    permission_cache.add(file_id)


def upload_resume_to_google_drive(
    service,
    folder_id,
    local_file_path,
    drive_file_name,
    visibility,
    permission_cache,
    existing_file_id='',
):
    guessed_type, _ = mimetypes.guess_type(local_file_path)
    mime_type = guessed_type or 'application/octet-stream'
    media = MediaFileUpload(local_file_path, mimetype=mime_type, resumable=False)

    target_file_id = existing_file_id
    if not target_file_id:
        existing = find_drive_file_by_name(service, folder_id, drive_file_name)
        if existing:
            target_file_id = existing.get('id', '')

    try:
        if target_file_id:
            file_metadata = service.files().update(
                fileId=target_file_id,
                media_body=media,
                fields='id',
                supportsAllDrives=True,
            ).execute()
        else:
            file_metadata = service.files().create(
                body={'name': drive_file_name, 'parents': [folder_id]},
                media_body=media,
                fields='id',
                supportsAllDrives=True,
            ).execute()
    except HttpError as exc:
        if not is_recoverable_response_preparation_failure(exc):
            raise

        # Google occasionally returns a malformed response body despite a successful upload.
        recovered_file_id = target_file_id
        if not recovered_file_id:
            existing_after = find_drive_file_by_name(service, folder_id, drive_file_name)
            if existing_after:
                recovered_file_id = existing_after.get('id', '')
        if not recovered_file_id:
            raise
        file_metadata = {'id': recovered_file_id}

    file_id = file_metadata.get('id', '')
    if file_id:
        ensure_drive_view_permission(service, file_id, visibility, permission_cache)

    web_view_link = ''
    web_content_link = ''
    if file_id:
        try:
            web_view_link, web_content_link = get_drive_file_links(service, file_id)
        except Exception:
            # Link metadata fetch can fail transiently; the constructed fallback URL remains usable.
            pass

    if web_view_link:
        return web_view_link, file_id
    if web_content_link:
        return web_content_link, file_id
    return f'https://drive.google.com/file/d/{file_id}/view?usp=sharing', file_id


def safe_filename(value):
    cleaned = re.sub(r'[<>:"/\\|?*]', '_', value)
    cleaned = re.sub(r'\s+', ' ', cleaned).strip().rstrip('.')
    return cleaned or 'Unknown_Employer'


def normalize_student_key(first_name, last_name, email):
    email_clean = str(email or '').strip().lower()
    if email_clean:
        return f"email:{email_clean}"
    full_name = f"{str(first_name or '').strip()} {str(last_name or '').strip()}".strip().lower()
    return f"name:{full_name}" if full_name else ''


def normalize_resume_filename(first_name, last_name, original_extension):
    full_name = f"{str(first_name or '').strip()}_{str(last_name or '').strip()}".strip('_')
    safe_name = safe_filename(full_name or 'Unknown_Student').replace(' ', '_')
    extension = (original_extension or '').lower()
    if not extension:
        extension = '.pdf'
    # Preserve the original extension to avoid MIME-type mismatches.
    return f"{safe_name}_Resume{extension}"


def build_resume_url(saved_file_path):
    file_name = os.path.basename(saved_file_path)
    if resume_public_file_base_url:
        return f"{resume_public_file_base_url.rstrip('/')}/{urllib.parse.quote(file_name)}?web=1"
    if resume_public_base_url:
        # A short SharePoint shared-folder URL (/:f:/) cannot resolve individual file
        # paths without Graph API calls; append a filename hint instead.
        if '/:f:/' in resume_public_base_url:
            return f"{resume_public_base_url}#file={urllib.parse.quote(file_name)}"
        return f"{resume_public_base_url.rstrip('/')}/{urllib.parse.quote(file_name)}"
    return f"file:///{saved_file_path.replace(os.sep, '/')}"


def save_attachment_to_path(attachment, destination_path):
    if hasattr(attachment, 'save'):
        directory = os.path.dirname(destination_path)
        custom_name = os.path.basename(destination_path)
        return attachment.save(location=directory, custom_name=custom_name)
    if hasattr(attachment, 'download'):
        return attachment.download(custom_name=destination_path)
    raise RuntimeError(f"Attachment object does not support save/download: {type(attachment).__name__}")


def to_excel_hyperlink(url, label='View Resume'):
    if not url:
        return ''
    safe_url = str(url).replace('"', '""')
    safe_label = str(label).replace('"', '""')
    return f'=HYPERLINK("{safe_url}","{safe_label}")'


def propagate_attachment_level_date(df, column_name):
    if column_name not in df.columns:
        return

    cleaned_series = df[column_name].replace(r'^\s*$', pd.NA, regex=True)
    non_empty_values = cleaned_series.dropna()
    if non_empty_values.empty:
        df[column_name] = cleaned_series
        return

    # The upload date appears on only one metadata row; forward-fill it to all student rows.
    df[column_name] = cleaned_series.fillna(non_empty_values.iloc[0])

# --- Section 2: Authenticate with Microsoft Graph ---
account = Account(credentials, auth_flow_type='credentials', tenant_id=tenant_id)

if account.authenticate(scopes=['https://graph.microsoft.com/.default']):
    print("Successfully connected to Outlook!")
    print(f"Reports output folder: {reports_dir}")

    mailbox = account.mailbox(resource=work_email)
    archive_folder = mailbox.get_folder(folder_name='Archive')
    if archive_folder is None:
        raise RuntimeError("Archive folder not found. Verify the mailbox has an Archive folder and the app has access.")

    # --- Section 3: Pull full alert-email history from the Archive folder ---
    if not alert_sender_email:
        raise RuntimeError("ALERT_SENDER_EMAIL environment variable is not set.")
    query = archive_folder.new_query().equals('from', alert_sender_email)
    messages = archive_folder.get_messages(query=query, limit=None, download_attachments=True)

    if enable_google_drive:
        (
            google_drive_service,
            google_drive_active_auth_mode,
            google_drive_credentials_path,
            google_drive_token_path,
        ) = initialize_google_drive_service()
        use_google_drive_links = google_drive_service is not None
        google_drive_permission_cache = set()
        if use_google_drive_links:
            print(f"Google Drive upload enabled (folder: {google_drive_folder_id}).")
            print(f"Google Drive auth mode: {google_drive_active_auth_mode}")
            print(f"Google credentials file: {google_drive_credentials_path}")
            if google_drive_active_auth_mode == 'oauth_user' and google_drive_token_path:
                print(f"Google OAuth token file: {google_drive_token_path}")
        else:
            print("Google Drive upload not active; using existing resume link behavior.")
    else:
        google_drive_service = None
        google_drive_active_auth_mode = ''
        google_drive_credentials_path = ''
        google_drive_token_path = ''
        use_google_drive_links = False
        google_drive_permission_cache = set()
        print("Google Drive upload disabled by default; using local resume saves only.")
    google_drive_disable_reason = ''

    # In-memory buckets for per-employer DataFrames, resume links, and run-level counters.
    employer_data = {}
    resume_link_by_name = {}
    resume_link_by_person = {}
    resume_latest_received_by_person = {}
    resume_drive_file_id_by_person = {}
    messages_scanned = 0
    messages_with_attachments = 0
    matching_name_attachments = 0
    matched_attachments = 0
    unmatched_attachment_names = set()
    failed_attachment_reads = 0
    resumes_saved = 0
    resumes_reused = 0
    resumes_updated_with_newer_version = 0

    print("Scanning inbox and extracting data...")
    os.makedirs(reports_dir, exist_ok=True)
    if clear_reports_folder_each_run:
        for existing_name in os.listdir(reports_dir):
            existing_path = os.path.join(reports_dir, existing_name)
            if os.path.isfile(existing_path) and existing_name.endswith('_Student_Application_Report.xlsx'):
                try:
                    os.remove(existing_path)
                except PermissionError:
                    print(f"Skipped deleting locked report file: {existing_name}")

    os.makedirs(resumes_dir, exist_ok=True)
    if clear_resumes_folder_each_run:
        for existing_name in os.listdir(resumes_dir):
            existing_path = os.path.join(resumes_dir, existing_name)
            if os.path.isfile(existing_path) and os.path.splitext(existing_name)[1].lower() in resume_extensions:
                os.remove(existing_path)

    for msg in messages:
        messages_scanned += 1
        # A. Record when the application email was received.
        msg_received = msg.received
        app_date = msg_received.date()

        # B. Extract the employer name from the alert email body.
        match = re.search(r'application to (.*?) through', msg.body)
        employer = match.group(1).strip() if match else "Unknown_Employer"

        if msg.attachments:
            messages_with_attachments += 1

        # Pass 1: locate and process the student-data Excel attachment.
        for attachment in msg.attachments:
            if attachment.name != attachment_name_expected:
                continue
            matching_name_attachments += 1
            temp_path = 'temp_student_data.xlsx'
            try:
                save_attachment_to_path(attachment, temp_path)
                df = pd.read_excel(temp_path)
                df['Date of Confirmed Application'] = app_date
                propagate_attachment_level_date(df, 'Date Database Uploaded')

                resume_name_to_person = {}
                if 'Resume File Name' in df.columns:
                    temp_resume_names = df['Resume File Name'].fillna('').astype(str).str.strip()
                    first_names = df['First Name'].fillna('') if 'First Name' in df.columns else pd.Series('', index=df.index)
                    last_names = df['Last Name'].fillna('') if 'Last Name' in df.columns else pd.Series('', index=df.index)
                    emails = df['Email'].fillna('') if 'Email' in df.columns else pd.Series('', index=df.index)

                    for idx in df.index:
                        resume_name = str(temp_resume_names.loc[idx]).strip()
                        if not resume_name:
                            continue
                        person_key = normalize_student_key(first_names.loc[idx], last_names.loc[idx], emails.loc[idx])
                        if not person_key:
                            continue
                        resume_name_to_person[resume_name.lower()] = {
                            'person_key': person_key,
                            'first_name': str(first_names.loc[idx]).strip(),
                            'last_name': str(last_names.loc[idx]).strip(),
                        }

                    # Pass 2: save resume attachments with normalized per-student filenames.
                    for msg_attachment in msg.attachments:
                        if msg_attachment.name == attachment_name_expected:
                            continue
                        original_name_key = msg_attachment.name.strip().lower()
                        attachment_ext = os.path.splitext(msg_attachment.name)[1].lower()

                        if attachment_ext not in resume_extensions:
                            unmatched_attachment_names.add(msg_attachment.name)
                            continue

                        student_info = resume_name_to_person.get(original_name_key)
                        if not student_info:
                            unmatched_attachment_names.add(msg_attachment.name)
                            continue

                        person_key = student_info['person_key']
                        existing_received = resume_latest_received_by_person.get(person_key)
                        if existing_received is not None and msg_received <= existing_received:
                            resume_link_by_name[original_name_key] = resume_link_by_person[person_key]
                            resumes_reused += 1
                            continue

                        normalized_name = normalize_resume_filename(
                            student_info['first_name'],
                            student_info['last_name'],
                            attachment_ext,
                        )
                        save_path = os.path.join(resumes_dir, normalized_name)
                        try:
                            save_attachment_to_path(msg_attachment, save_path)

                            resume_url = ''
                            if use_google_drive_links:
                                try:
                                    resume_url, drive_file_id = upload_resume_to_google_drive(
                                        service=google_drive_service,
                                        folder_id=google_drive_folder_id,
                                        local_file_path=save_path,
                                        drive_file_name=normalized_name,
                                        visibility=google_drive_link_visibility,
                                        permission_cache=google_drive_permission_cache,
                                        existing_file_id=resume_drive_file_id_by_person.get(person_key, ''),
                                    )
                                    if drive_file_id:
                                        resume_drive_file_id_by_person[person_key] = drive_file_id
                                except Exception as exc:
                                    if not google_drive_disable_reason:
                                        google_drive_disable_reason = str(exc)
                                        print(
                                            'Google Drive uploads disabled for this run after first failure. '
                                            f"Falling back to existing links. Error: {exc}"
                                        )
                                    use_google_drive_links = False

                            if not resume_url:
                                resume_url = build_resume_url(save_path)

                            had_existing_person_resume = person_key in resume_link_by_person
                            resume_link_by_person[person_key] = resume_url
                            resume_link_by_name[original_name_key] = resume_url
                            resume_latest_received_by_person[person_key] = msg_received

                            if had_existing_person_resume:
                                resumes_updated_with_newer_version += 1
                            else:
                                resumes_saved += 1
                        except Exception as exc:
                            print(f"Failed to save resume attachment '{msg_attachment.name}' for {employer}: {exc}")

                if 'Resume File Name' in df.columns:
                    resume_names = df['Resume File Name'].fillna('').astype(str).str.strip()
                    first_names = df['First Name'].fillna('') if 'First Name' in df.columns else pd.Series('', index=df.index)
                    last_names = df['Last Name'].fillna('') if 'Last Name' in df.columns else pd.Series('', index=df.index)
                    emails = df['Email'].fillna('') if 'Email' in df.columns else pd.Series('', index=df.index)

                    def row_resume_link(idx):
                        name_key = str(resume_names.loc[idx]).strip().lower()
                        if name_key and name_key in resume_link_by_name:
                            return resume_link_by_name[name_key]
                        person_key = normalize_student_key(first_names.loc[idx], last_names.loc[idx], emails.loc[idx])
                        return resume_link_by_person.get(person_key, '') if person_key else ''

                    df['Resume Link'] = [row_resume_link(idx) for idx in df.index]
                    df.drop(columns=['Resume File Name'], inplace=True)

                # Accumulate this email's DataFrame into the employer's in-memory bucket.
                if employer not in employer_data:
                    employer_data[employer] = []
                employer_data[employer].append(df)
                matched_attachments += 1
            except Exception as exc:
                failed_attachment_reads += 1
                print(f"Failed to process attachment '{attachment.name}' for {employer}: {exc}")
            finally:
                if os.path.exists(temp_path):
                    os.remove(temp_path)

        print(f"Grabbed application for {employer}")

    # --- Section 4: Consolidate and export per-employer Excel reports ---
    print("\nConsolidating data and generating Excel reports...")
    exported_reports = 0
    for emp, df_list in employer_data.items():
        
        # Concatenate all per-email DataFrames for this employer into one.
        final_df = pd.concat(df_list, ignore_index=True)

        # Drop rows with no student identity (no first name, last name, or email).
        first_name_series = final_df['First Name'].fillna('').astype(str).str.strip() if 'First Name' in final_df.columns else pd.Series('', index=final_df.index)
        last_name_series = final_df['Last Name'].fillna('').astype(str).str.strip() if 'Last Name' in final_df.columns else pd.Series('', index=final_df.index)
        email_series = final_df['Email'].fillna('').astype(str).str.strip() if 'Email' in final_df.columns else pd.Series('', index=final_df.index)
        has_identity = (first_name_series != '') | (last_name_series != '') | (email_series != '')
        final_df = final_df[has_identity].copy()

        # Deduplicate by email address, keeping the row with the latest confirmed application date.
        if 'Email' in final_df.columns:
            final_df['Date of Confirmed Application'] = pd.to_datetime(
                final_df['Date of Confirmed Application'], errors='coerce'
            )
            final_df['__email_key'] = final_df['Email'].fillna('').astype(str).str.strip().str.lower()

            with_email = final_df[final_df['__email_key'] != ''].copy()
            without_email = final_df[final_df['__email_key'] == ''].copy()

            with_email.sort_values(by='Date of Confirmed Application', inplace=True)
            with_email = with_email.drop_duplicates(subset='__email_key', keep='last')

            final_df = pd.concat([with_email, without_email], ignore_index=True)
            final_df.drop(columns=['__email_key'], inplace=True)
        else:
            # Fallback: drop exact duplicate rows when the Email column is absent.
            final_df.drop_duplicates(inplace=True)

        # Format the application date as m/d/yyyy (no time component).
        if 'Date of Confirmed Application' in final_df.columns:
            final_df['Date of Confirmed Application'] = final_df['Date of Confirmed Application'].apply(
                lambda d: f"{d.month}/{d.day}/{d.year}" if pd.notna(d) else ""
            )

        # Convert resume URL strings to clickable Excel HYPERLINK formulas.
        if 'Resume Link' in final_df.columns:
            final_df['Resume Link'] = final_df['Resume Link'].fillna('').astype(str).apply(to_excel_hyperlink)

        # Strip internal-only columns that should not appear in employer exports.
        existing_columns_to_remove = [
            col
            for col in final_df.columns
            if re.sub(r'\s+', ' ', str(col).strip()).lower() in columns_to_remove_from_exports_normalized
        ]
        if existing_columns_to_remove:
            final_df.drop(columns=existing_columns_to_remove, inplace=True)

        # Reorder so Date of Confirmed Application appears as the first column.
        if 'Date of Confirmed Application' in final_df.columns:
            ordered_columns = ['Date of Confirmed Application'] + [
                col for col in final_df.columns if col != 'Date of Confirmed Application'
            ]
            final_df = final_df[ordered_columns]
        
        # Write the final DataFrame to an Excel file named after the employer.
        filename = f"{safe_filename(emp)}_Student_Application_Report.xlsx"
        output_path = os.path.join(reports_dir, filename)
        try:
            final_df.to_excel(output_path, index=False)
            exported_reports += 1
            print(f"Successfully created {output_path} with {len(final_df)} records.")
        except PermissionError:
            print(f"Skipped writing locked report file: {output_path}")

    print("\nDiagnostics summary:")
    print(f"- Messages scanned: {messages_scanned}")
    print(f"- Messages with attachments: {messages_with_attachments}")
    print(f"- Attachments named '{attachment_name_expected}' found: {matching_name_attachments}")
    print(f"- Attachments named '{attachment_name_expected}' processed: {matched_attachments}")
    print(f"- Attachment read failures: {failed_attachment_reads}")
    print(f"- Employers with data prepared: {len(employer_data)}")
    print(f"- Reports exported: {exported_reports}")
    print(f"- Resume links available: {len(resume_link_by_name)}")
    print(f"- Unique resumes saved: {resumes_saved}")
    print(f"- Resume mappings reused: {resumes_reused}")
    print(f"- Resume mappings updated to newer version: {resumes_updated_with_newer_version}")
    print(f"- Google Drive links active: {use_google_drive_links}")
    if use_google_drive_links:
        print(f"- Google Drive auth mode used: {google_drive_active_auth_mode}")
    if google_drive_disable_reason:
        print(f"- Google Drive disabled reason: {google_drive_disable_reason}")
    if unmatched_attachment_names:
        print("- Other attachment names seen:")
        for name in sorted(unmatched_attachment_names):
            print(f"  * {name}")
    if exported_reports == 0:
        print("- No report files were written. Check attachment names above and whether expected files are present.")

    print("\nAll tasks complete!")

else:
    print("Authentication failed. Verify GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET, and GRAPH_TENANT_ID.")