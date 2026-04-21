"""
SCALE Stories of Success Candidate Builder
==========================================

Dependencies
------------
    pip install pandas openpyxl numpy

Script Workflow
---------------
    1. Prompt the user (or drag-and-drop) for the SCALE student-database Excel file.
    2. Load three sheets from that workbook: Students, Degrees, Work Experience.
    3. Evaluate each student against four criteria:
         a. Has defense-industry experience (flag on student record).
         b. Was in SCALE for at least 2 years before graduating.
         c. Earned multiple degrees, including at least one graduate degree (MS or PhD).
         d. Has at least 2 total work experiences, with at least 1 defense-related.
    4. Write a timestamped Excel report to the same directory as this script containing:
         - Candidates            - students who meet ALL criteria (cleaned, sorted output)
         - AllStudentsEvaluated  - full scoring detail for every student
         - Summary               - run parameters and result counts

Arguments / Thresholds (set via StoriesOfSuccessProcessor constructor)
----------------------------------------------------------------------
    database_file_path           Path to the SCALE export Excel file (required).
    output_dir                   Directory to write the report.
                                 Defaults to the folder containing this script.
    min_years_in_scale           Minimum years from SCALE start to graduation (default 2.0).
    min_total_work_experiences   Minimum total internships/co-ops (default 2).
    min_defense_work_experiences Minimum defense-related work experiences (default 1).

Input Sheet Requirements
------------------------
    Students sheet       : email, first/last name, SCALE start semester, defense-experience flag
    Degrees sheet        : student email, degree type, graduation date
    WorkExperience sheet : student email, defense-related flag

Column and sheet names are resolved flexibly (case- and whitespace-insensitive) so
minor naming variations between exports are handled automatically.
"""

import re
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd


# Default output directory — same folder as this script.
# Override by passing output_dir= to StoriesOfSuccessProcessor.
DEFAULT_OUTPUT_DIR = Path(__file__).parent


class StoriesOfSuccessProcessor:
	"""Evaluate SCALE export data and identify Stories of Success candidates."""

	def __init__(
		self,
		database_file_path,
		output_dir=DEFAULT_OUTPUT_DIR,
		min_years_in_scale=2.0,
		min_total_work_experiences=2,
		min_defense_work_experiences=1,
	):
		self.database_file_path = Path(database_file_path)
		self.output_dir = Path(output_dir)
		self.min_years_in_scale = float(min_years_in_scale)
		self.min_total_work_experiences = int(min_total_work_experiences)
		self.min_defense_work_experiences = int(min_defense_work_experiences)

		self.students_df = None
		self.degrees_df = None
		self.work_experience_df = None

	def _normalize_key(self, value):
		"""Strip all non-alphanumeric characters and lowercase for fuzzy matching."""
		if value is None:
			return ""
		return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())

	def _normalize_email(self, value):
		"""Return a lowercase, stripped email or pd.NA for empty/null values."""
		if pd.isna(value):
			return pd.NA
		text = str(value).strip().lower()
		return text if text else pd.NA

	def _excel_col_letter_to_index(self, letter):
		"""Convert an Excel column letter (e.g. 'AD') to a 0-based integer index."""
		index = 0
		for char in str(letter).strip().upper():
			if char < "A" or char > "Z":
				raise ValueError(f"Invalid Excel column letter: {letter}")
			index = (index * 26) + (ord(char) - ord("A") + 1)
		return index - 1

	def _column_by_letter(self, df, letter):
		"""Return the column name at the given Excel column letter, or None if out of range."""
		try:
			index = self._excel_col_letter_to_index(letter)
		except ValueError:
			return None

		if 0 <= index < len(df.columns):
			return df.columns[index]
		return None

	def _resolve_sheet_name(self, workbook_sheets, candidates, required=True):
		"""
		Find a sheet name from the workbook by matching against a list of candidate names.

		Tries exact normalized match first, then partial/substring match.
		Raises ValueError if no match is found and required=True.
		"""
		normalized_lookup = {
			self._normalize_key(sheet_name): sheet_name for sheet_name in workbook_sheets
		}

		for candidate in candidates:
			key = self._normalize_key(candidate)
			if key in normalized_lookup:
				return normalized_lookup[key]

		for sheet_key, original_name in normalized_lookup.items():
			for candidate in candidates:
				candidate_key = self._normalize_key(candidate)
				if candidate_key and (candidate_key in sheet_key or sheet_key in candidate_key):
					return original_name

		if required:
			raise ValueError(f"Could not find required sheet. Tried: {candidates}")
		return None

	def _resolve_column_name(self, df, candidates, fallback_letter=None, required=False):
		"""
		Find a column in a DataFrame by matching against candidate names.

		Tries exact normalized match, then substring match, then Excel column
		letter fallback.  Raises ValueError if required and nothing matches.
		"""
		normalized_lookup = {}
		for col in df.columns:
			key = self._normalize_key(col)
			if key and key not in normalized_lookup:
				normalized_lookup[key] = col

		for candidate in candidates:
			candidate_key = self._normalize_key(candidate)
			if candidate_key in normalized_lookup:
				return normalized_lookup[candidate_key]

		for col in df.columns:
			col_key = self._normalize_key(col)
			for candidate in candidates:
				candidate_key = self._normalize_key(candidate)
				if candidate_key and (candidate_key in col_key or col_key in candidate_key):
					return col

		if fallback_letter:
			fallback_col = self._column_by_letter(df, fallback_letter)
			if fallback_col is not None:
				return fallback_col

		if required:
			attempted = ", ".join(candidates)
			raise ValueError(f"Could not find required column. Tried: {attempted}")

		return None

	def _parse_term_year(self, text_value, mode):
		"""
		Parse a SCALE academic-term string (e.g. "Fall 2024", "FA 24", "24 Spring")
		into a Timestamp.

		mode="scale_start" maps to the beginning of the term;
		mode="graduation" maps to the end of the term.
		Returns pd.NaT on failure.
		"""
		if pd.isna(text_value):
			return pd.NaT

		text = str(text_value).strip()
		if not text:
			return pd.NaT

		normalized = text.replace("\u2019", "'").replace("-", " ")

		# Match "Fall 2024" / "FA 24" style
		direct_match = re.search(
			r"\b(spring|summer|fall|sp|su|fa)\b\s*[\'\s]*(\d{2,4})\b",
			normalized,
			re.IGNORECASE,
		)
		# Match "24 Spring" / "2024 Fall" style
		reverse_match = re.search(
			r"\b(\d{2,4})\b\s*[\'\s]*(spring|summer|fall|sp|su|fa)\b",
			normalized,
			re.IGNORECASE,
		)

		term = None
		year_text = None
		if direct_match:
			term = direct_match.group(1).lower()
			year_text = direct_match.group(2)
		elif reverse_match:
			year_text = reverse_match.group(1)
			term = reverse_match.group(2).lower()

		if term and year_text:
			year = int(year_text)
			if year < 100:
				year += 2000

			if term in {"sp", "spring"}:
				if mode == "scale_start":
					return pd.Timestamp(year=year, month=1, day=15)
				return pd.Timestamp(year=year, month=5, day=15)

			if term in {"su", "summer"}:
				if mode == "scale_start":
					return pd.Timestamp(year=year, month=5, day=15)
				return pd.Timestamp(year=year, month=8, day=15)

			if term in {"fa", "fall"}:
				if mode == "scale_start":
					return pd.Timestamp(year=year, month=8, day=15)
				return pd.Timestamp(year=year, month=12, day=15)

		return pd.NaT

	def _parse_flexible_datetime(self, value, mode):
		"""
		Parse a date value that may be a term string, a Timestamp, or an ISO date string.

		mode is forwarded to _parse_term_year.  Returns pd.NaT on failure.
		"""
		if pd.isna(value) or value == "":
			return pd.NaT

		if isinstance(value, pd.Timestamp):
			return value

		term_date = self._parse_term_year(value, mode=mode)
		if not pd.isna(term_date):
			return term_date

		try:
			parsed = pd.to_datetime(value, errors="coerce")
			return parsed if not pd.isna(parsed) else pd.NaT
		except Exception:
			return pd.NaT

	def _format_date_for_output(self, value):
		"""Format a date value as M/D/YYYY for Excel output, or empty string if missing."""
		if pd.isna(value):
			return ""
		parsed = pd.to_datetime(value, errors="coerce")
		if pd.isna(parsed):
			return ""
		return f"{parsed.month}/{parsed.day}/{parsed.year}"

	def _students_defense_flag(self, value):
		"""
		Return True if the student record indicates defense-industry experience.

		The field may be stored as a numeric code (20 = yes) or a text boolean.
		"""
		if pd.isna(value):
			return False

		numeric = pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0]
		if not pd.isna(numeric):
			return int(round(numeric)) == 20

		text = str(value).strip().lower()
		return text in {"yes", "y", "true", "1", "20", "20.0"}

	def _work_defense_flag(self, value):
		"""
		Return 1 if a work-experience record is defense-related, 0 otherwise.

		Accepts numeric 1/0 or text boolean representations.
		"""
		if pd.isna(value):
			return 0

		numeric = pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0]
		if not pd.isna(numeric):
			return int(round(numeric)) == 1

		text = str(value).strip().lower()
		return text in {"yes", "y", "true", "1"}

	def _is_graduate_degree(self, degree_type_value):
		"""Return True if the degree type is MS, MSc, or PhD-level."""
		if pd.isna(degree_type_value):
			return False

		normalized = re.sub(r"[^a-z]", "", str(degree_type_value).lower())
		if not normalized:
			return False

		if "phd" in normalized or "doctor" in normalized:
			return True
		if normalized in {"ms", "msc"}:
			return True
		if "master" in normalized:
			return True

		return False

	def load_data(self):
		"""Read the three required sheets from the SCALE export workbook."""
		print(f"Loading workbook: {self.database_file_path}")

		workbook_sheets = pd.read_excel(
			self.database_file_path,
			sheet_name=None,
			engine="openpyxl",
		)

		students_sheet = self._resolve_sheet_name(workbook_sheets.keys(), ["students", "student"])
		degrees_sheet = self._resolve_sheet_name(workbook_sheets.keys(), ["degrees", "degree"])
		work_sheet = self._resolve_sheet_name(
			workbook_sheets.keys(),
			["workExperience", "work experience", "work_experience", "work"],
		)

		self.students_df = workbook_sheets[students_sheet].copy()
		self.degrees_df = workbook_sheets[degrees_sheet].copy()
		self.work_experience_df = workbook_sheets[work_sheet].copy()

		print(f"Loaded tab '{students_sheet}' with {len(self.students_df)} rows")
		print(f"Loaded tab '{degrees_sheet}' with {len(self.degrees_df)} rows")
		print(f"Loaded tab '{work_sheet}' with {len(self.work_experience_df)} rows")

	def _build_students_base(self):
		"""
		Extract and normalize core student fields.

		Returns a deduplicated DataFrame keyed by _email_key with columns:
		email, firstName, lastName, studentScaleSemester, hasDefenseExperience, scaleStartDate.
		"""
		email_col = self._resolve_column_name(
			self.students_df,
			["email", "studentEmail", "student email", "emailAddress", "email address"],
			required=True,
		)
		first_name_col = self._resolve_column_name(
			self.students_df,
			["firstName", "first name", "firstname"],
		)
		last_name_col = self._resolve_column_name(
			self.students_df,
			["lastName", "last name", "lastname", "surname"],
		)
		scale_semester_col = self._resolve_column_name(
			self.students_df,
			["studentScaleSemester", "student scale semester", "scale semester"],
			fallback_letter="AD",
			required=True,
		)
		defense_experience_col = self._resolve_column_name(
			self.students_df,
			["hasDefenseExperience", "defenseExperience", "defense experience", "in defense"],
			fallback_letter="CE",
			required=True,
		)

		base_df = pd.DataFrame(
			{
				"email": self.students_df[email_col].astype(str).str.strip(),
				"_email_key": self.students_df[email_col].apply(self._normalize_email),
				"firstName": self.students_df[first_name_col] if first_name_col else "",
				"lastName": self.students_df[last_name_col] if last_name_col else "",
				"studentScaleSemester": self.students_df[scale_semester_col],
				"defenseExperienceRaw": self.students_df[defense_experience_col],
			}
		)

		base_df = base_df[base_df["_email_key"].notna()].copy()
		base_df = base_df.drop_duplicates(subset=["_email_key"], keep="first").reset_index(drop=True)

		base_df["hasDefenseExperience"] = base_df["defenseExperienceRaw"].apply(self._students_defense_flag)
		base_df["scaleStartDate"] = base_df["studentScaleSemester"].apply(
			lambda value: self._parse_flexible_datetime(value, mode="scale_start")
		)

		return base_df

	def _build_degrees_agg(self):
		"""
		Aggregate degree records per student.

		Returns a DataFrame keyed by _email_key with counts and the latest graduation date.
		"""
		email_col = self._resolve_column_name(
			self.degrees_df,
			["studentEmail", "student email", "email", "emailAddress", "email address"],
			fallback_letter="A",
			required=True,
		)
		degree_type_col = self._resolve_column_name(
			self.degrees_df,
			["degreeType", "degree type", "degree"],
			fallback_letter="D",
			required=True,
		)
		graduation_date_col = self._resolve_column_name(
			self.degrees_df,
			["graduationDate", "graduation date", "gradDate", "grad date"],
			required=True,
		)

		working_df = self.degrees_df.copy()
		working_df["_email_key"] = working_df[email_col].apply(self._normalize_email)
		working_df = working_df[working_df["_email_key"].notna()].copy()

		working_df["degreeTypeNormalized"] = (
			working_df[degree_type_col].fillna("").astype(str).str.strip().str.upper()
		)
		working_df["isGraduateDegree"] = working_df[degree_type_col].apply(self._is_graduate_degree)
		working_df["graduationDateParsed"] = working_df[graduation_date_col].apply(
			lambda value: self._parse_flexible_datetime(value, mode="graduation")
		)

		agg_df = (
			working_df.groupby("_email_key", as_index=False)
			.agg(
				totalDegrees=("_email_key", "size"),
				graduateDegreeCount=("isGraduateDegree", "sum"),
				degreeTypes=(
					"degreeTypeNormalized",
					lambda values: ", ".join(
						sorted({value for value in values if isinstance(value, str) and value})
					),
				),
				latestGraduationDate=("graduationDateParsed", "max"),
			)
		)

		return agg_df

	def _build_work_experience_agg(self):
		"""
		Aggregate work experience records per student.

		Returns a DataFrame keyed by _email_key with total and defense-related counts.
		"""
		email_col = self._resolve_column_name(
			self.work_experience_df,
			["studentEmail", "student email", "email", "emailAddress", "email address"],
			fallback_letter="O",
			required=True,
		)
		defense_related_col = self._resolve_column_name(
			self.work_experience_df,
			["defenseRelated", "defense related", "isDefenseRelated"],
			fallback_letter="L",
			required=True,
		)

		working_df = self.work_experience_df.copy()
		working_df["_email_key"] = working_df[email_col].apply(self._normalize_email)
		working_df = working_df[working_df["_email_key"].notna()].copy()

		working_df["isDefenseWorkExperience"] = working_df[defense_related_col].apply(self._work_defense_flag)

		agg_df = (
			working_df.groupby("_email_key", as_index=False)
			.agg(
				totalWorkExperiences=("_email_key", "size"),
				defenseWorkExperiences=("isDefenseWorkExperience", "sum"),
			)
		)

		return agg_df

	def _calculate_years_in_scale(self, row):
		"""Compute years between SCALE start date and latest graduation date."""
		scale_start = row.get("scaleStartDate")
		graduation_date = row.get("latestGraduationDate")

		if pd.isna(scale_start) or pd.isna(graduation_date):
			return np.nan

		delta_days = (graduation_date - scale_start).days
		return round(delta_days / 365.25, 2)

	def evaluate_students(self):
		"""
		Join student, degree, and work-experience data; apply all four criteria;
		and return a fully scored DataFrame.
		"""
		students_base_df = self._build_students_base()
		degrees_agg_df = self._build_degrees_agg()
		work_agg_df = self._build_work_experience_agg()

		# Left-join so every student appears even with no degree / work records
		scored_df = students_base_df.merge(degrees_agg_df, how="left", on="_email_key")
		scored_df = scored_df.merge(work_agg_df, how="left", on="_email_key")

		# Fill missing aggregates with zero
		scored_df["totalDegrees"] = scored_df["totalDegrees"].fillna(0).astype(int)
		scored_df["graduateDegreeCount"] = scored_df["graduateDegreeCount"].fillna(0).astype(int)
		scored_df["degreeTypes"] = scored_df["degreeTypes"].fillna("")
		scored_df["totalWorkExperiences"] = scored_df["totalWorkExperiences"].fillna(0).astype(int)
		scored_df["defenseWorkExperiences"] = scored_df["defenseWorkExperiences"].fillna(0).astype(int)

		# Criterion 2: years in SCALE before graduation
		scored_df["yearsInScaleBeforeGraduation"] = scored_df.apply(
			self._calculate_years_in_scale,
			axis=1,
		)
		scored_df["inScaleTwoPlusYears"] = (
			scored_df["yearsInScaleBeforeGraduation"] >= self.min_years_in_scale
		)

		# Criterion 3: degree criteria (multiple degrees including at least one MS/PhD)
		scored_df["hasMultipleDegrees"] = scored_df["totalDegrees"] >= 2
		scored_df["hasGraduateDegreeMSorPHD"] = scored_df["graduateDegreeCount"] >= 1
		scored_df["meetsDegreeCriteria"] = (
			scored_df["hasMultipleDegrees"] & scored_df["hasGraduateDegreeMSorPHD"]
		)

		# Criterion 4: work experience criteria (total and defense-specific minimums)
		scored_df["meetsStrongInternshipCriteria"] = (
			(scored_df["totalWorkExperiences"] >= self.min_total_work_experiences)
			& (
				scored_df["defenseWorkExperiences"]
				>= self.min_defense_work_experiences
			)
		)

		# Overall: all four criteria must be met
		scored_df["meetsAllCriteria"] = (
			scored_df["hasDefenseExperience"]
			& scored_df["inScaleTwoPlusYears"]
			& scored_df["meetsDegreeCriteria"]
			& scored_df["meetsStrongInternshipCriteria"]
		)

		# Format dates for Excel-friendly string output
		scored_df["scaleStartDate"] = scored_df["scaleStartDate"].apply(self._format_date_for_output)
		scored_df["graduationDateUsed"] = scored_df["latestGraduationDate"].apply(self._format_date_for_output)

		output_columns = [
			"firstName",
			"lastName",
			"email",
			"defenseExperienceRaw",
			"hasDefenseExperience",
			"studentScaleSemester",
			"scaleStartDate",
			"graduationDateUsed",
			"yearsInScaleBeforeGraduation",
			"inScaleTwoPlusYears",
			"totalDegrees",
			"graduateDegreeCount",
			"degreeTypes",
			"hasMultipleDegrees",
			"hasGraduateDegreeMSorPHD",
			"meetsDegreeCriteria",
			"totalWorkExperiences",
			"defenseWorkExperiences",
			"meetsStrongInternshipCriteria",
			"meetsAllCriteria",
		]

		return scored_df[output_columns].copy()

	def save_output(self, scored_df):
		"""
		Write the three-sheet Excel report to output_dir.

		Sheets:
		    Candidates           - students meeting all criteria, cleaned columns, sorted by strength
		    AllStudentsEvaluated - full scoring detail for every evaluated student
		    Summary              - run parameters and result counts

		Returns the Path of the written file, or None on permission error.
		"""
		self.output_dir.mkdir(parents=True, exist_ok=True)

		# Timestamped filename; append a counter if the file already exists
		timestamp = datetime.now().strftime("%m.%d.%Y %I.%M %p").lstrip("0")
		base_name = f"Stories of Success Candidates {timestamp}"
		output_path = self.output_dir / f"{base_name}.xlsx"

		suffix = 1
		while output_path.exists():
			output_path = self.output_dir / f"{base_name} ({suffix}).xlsx"
			suffix += 1

		# Candidates tab: sort strongest candidates first, hide internal scoring columns
		candidates_df = scored_df[scored_df["meetsAllCriteria"]].copy()
		candidates_df = candidates_df.sort_values(
			by=[
				"defenseWorkExperiences",
				"totalWorkExperiences",
				"yearsInScaleBeforeGraduation",
			],
			ascending=[False, False, False],
			na_position="last",
		).reset_index(drop=True)

		# Keep full scoring detail in AllStudentsEvaluated; strip internal flags from Candidates
		candidates_columns_to_remove = [
			"defenseExperienceRaw",
			"inScaleTwoPlusYears",
			"graduateDegreeCount",
			"hasMultipleDegrees",
			"hasGraduateDegreeMSorPHD",
			"meetsDegreeCriteria",
			"meetsStrongInternshipCriteria",
			"meetsAllCriteria",
		]
		candidates_export_df = candidates_df.drop(columns=candidates_columns_to_remove, errors="ignore")

		summary_df = pd.DataFrame(
			[
				{
					"metric": "students_evaluated",
					"value": len(scored_df),
				},
				{
					"metric": "students_meeting_all_criteria",
					"value": len(candidates_df),
				},
				{
					"metric": "minimum_years_in_scale",
					"value": self.min_years_in_scale,
				},
				{
					"metric": "minimum_total_work_experiences",
					"value": self.min_total_work_experiences,
				},
				{
					"metric": "minimum_defense_work_experiences",
					"value": self.min_defense_work_experiences,
				},
			]
		)

		try:
			with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
				candidates_export_df.to_excel(writer, sheet_name="Candidates", index=False)
				scored_df.to_excel(writer, sheet_name="AllStudentsEvaluated", index=False)
				summary_df.to_excel(writer, sheet_name="Summary", index=False)

				# Freeze the header row on every sheet
				for sheet_name in ("Candidates", "AllStudentsEvaluated", "Summary"):
					ws = writer.sheets.get(sheet_name)
					if ws is not None:
						ws.freeze_panes = "A2"
		except PermissionError:
			print("[ERROR] Could not write output file — it may be open in Excel.")
			return None

		print(f"[OK] Output saved to: {output_path}")
		print(f"[OK] Candidates found: {len(candidates_df)}")

		return output_path

	def run(self):
		"""Load data, evaluate candidates, and save the report."""
		print("=" * 60)
		print("SCALE STORIES OF SUCCESS CANDIDATE BUILDER")
		print("=" * 60)
		print(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
		print()

		self.load_data()
		scored_df = self.evaluate_students()
		output_path = self.save_output(scored_df)

		print()
		print("=" * 60)
		print("PROCESS COMPLETE")
		print("=" * 60)

		return output_path


def sanitize_drag_drop_path(raw_path):
	"""
	Clean a file path that may have been drag-and-dropped into the terminal.

	Removes leading '& ', surrounding quotes, and extra whitespace that
	Windows/PowerShell adds when files are dragged into the console.
	"""
	cleaned = raw_path.strip()
	if cleaned.startswith("& "):
		cleaned = cleaned[2:].strip()
	cleaned = cleaned.strip('"').strip("'").strip()
	return cleaned


def main():
	print("Please provide the path to the SCALE student database Excel file.")
	print("(You can drag and drop the file into this window)")
	print()

	input_path = input("Database file path: ").strip()
	input_path = sanitize_drag_drop_path(input_path)

	if not input_path:
		print("[ERROR] No file path provided. Exiting.")
		return

	database_file = Path(input_path)
	if not database_file.exists():
		print(f"[ERROR] File not found: {database_file}")
		return

	if database_file.suffix.lower() not in {".xlsx", ".xls", ".xlsm"}:
		print(f"[ERROR] Unsupported file type: {database_file.suffix}")
		return

	processor = StoriesOfSuccessProcessor(database_file_path=database_file)

	try:
		output_path = processor.run()
	except Exception as exc:
		print(f"[ERROR] Processing failed: {exc}")
		return

	if output_path is not None:
		print(f"[OK] Finished. Candidate report created at:\n  {output_path}")
	else:
		print("[INFO] Process finished but no output file was created.")


if __name__ == "__main__":
	main()
