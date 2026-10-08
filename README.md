# Need Correction & Correction Tool — Deep Analysis Build

## Run
```powershell
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Included
- `app.py` — Streamlit dashboard
- `requirements.txt` — Python dependencies
- `📊 Need Correction & Correction Tool - Logic Document.docx` — detailed business logic document
- `📊 Need Correction & Correction Tool - Logic Document.md` — concise logic reference
- `logic_rules.json` — machine-readable client/common rules

## Supported learned profiles
- Tetris Innovations
- Bambrew
- SunJewels
- Texan Labs

The tool supports both:
1. wide family data → long member data
2. already-long member data → standardized final data

For multiple source sheets, the user chooses which sheet(s) to process.

- Tetris output ordering matches the reference: all Self rows, then Spouse rows, then Child 1/2/3 rows.
- Bambrew output ordering matches the reference mixed Excel sort: numeric Employee IDs first, then text Employee IDs.
- Date-format standardization is included in the Corrections audit when the raw value is textual.


## Raw Data is never modified
The generated workbook preserves every selected source sheet as a separate Raw Data sheet. The tool does not add correction columns or overwrite values in those raw sheets.

For a single selected source sheet, the workbook structure is:
1. `<Original Sheet Name>` — original raw data, unchanged
2. `Corrected Data` — standardized/corrected final data
3. `Corrections` — detailed correction audit
4. `Count` — relationship/member validation
5. `Processing Status`
6. `Valid`
7. `SOP`

For multiple selected source sheets, each source gets its own `Corrected Data - <Source Sheet>`, `Corrections - <Source Sheet>`, and `Count - <Source Sheet>` sheets.


## v2.0 raw-data preservation fix
- Fixed the pandas `DataFrame.itertuples(..., header=False, ...)` error.
- Raw data is copied using `DataFrame.values.tolist()` and is never passed through the correction pipeline.
- Corrected output remains in separate `Corrected Data` sheet(s).
- The output workbook includes a success/status message in the dashboard before download.


## v3.0
- `Error` column removed from Corrected Data.
- `Remark` column added.
- Parent vs Parent-in-law source conflicts are explicitly noted for confirmation.
- GHI Sum Insured is entered by the user at runtime.
- Maximum Children / Employee is entered by the user at runtime.
- Added configurable dependent limits, age limits, age gaps and gender-combination rules.
- Email/contact validation is removed.
- Duplicate pre-check runs immediately after sheet selection and before transformation.


## v5.0
- Duplicate checking has been completely removed from the tool for now: no duplicate pre-check, duplicate validation, duplicate error columns, or duplicate rules are executed.
- The legacy individual validation columns are retained: `Error_Name`, `Error_DoB`, `Error_Age`, `Error_Gender`, `Error_Relationship`, `Error_Ph`, and `Error_Email`.
- The single combined `Error` column is not created.
- Email/phone validation is not performed; those two individual columns remain blank for compatibility with the earlier output structure.
- Date-of-birth normalization is still used internally to calculate Age and produce the standardized Corrected Data date, but DOB changes are no longer logged as automatic Corrections.
- The user-facing `Corrections` column now lists only the fields corrected (for example `Gender; Relationship to Self`). Detailed old/new values remain in the separate `Corrections` audit sheet.


## v6.0 — Clean Corrected Data output
Removed from the user-facing Corrected Data sheet:
- Error_Name
- Error_DoB
- Error_Age
- Error_Gender
- Error_Relationship
- Error_Ph
- Error_Email
- Correction Confidence

The Corrected Data sheet now keeps the core data plus `Corrections`, `Correction Rules`, and `Remark`.


## v7.0 — Age Band Pivot Table
Each processed source sheet now gets an `Age Band Pivot Table` sheet. It summarizes the number of Corrected Data member rows in each Age Band, sorted by starting age, and includes a Grand Total. With multiple selected sheets, each source gets its own pivot tab.


## v8.0 — Proceed button / processing visibility fix
- Added visible progress while each selected sheet is processed.
- Added a spinner while the final Excel workbook is being generated.
- Any processing or workbook-generation exception is now shown explicitly in the UI.
- Removed the obsolete Employment Type assignment.
- Fixed Age Band numeric sorting.
- The main action is now labelled `Proceed & Generate Corrected Data`.


## v9.0 — Naming update
- Tool name: **📊 Need Correction & Correction Tool**
- Original source sheets are preserved as **<Original Sheet Name> (Raw Data)**.
- Corrected Data remains separate and does not overwrite the raw sheet.


## v11.0 — Smart wide-format detection
Wide-to-long transformation is now header-driven rather than based on fixed column positions. The parser dynamically discovers Employee, Spouse, Child N, Father, Mother and Parent fields and associates DOB/Age fields by header context.

For the supplied Tetris workbook, Father's Name/Father's Age and Mother's Name/Mother's Age are therefore detected automatically even though they are later columns. They are converted to Parent members. Reordering or inserting unrelated columns should not break the transformation.


## v12.0 — Final fixes
- Application name is exactly **📊 Data Validation Tool**.
- Phone/mobile/contact numbers are detected by semantic header names and, if the header is absent or unclear, by inspecting actual values across all columns. No fixed phone column is assumed.
- Phone values are carried into `Mobile Number` for the employee/Self member.
- Final `Corrected Data` rows are sorted A→Z by Employee ID after transformation. The Excel header remains the first row and sorting is stable within the same Employee ID.


## Hybrid header intelligence
The tool now combines v12 aliases with conservative fuzzy matching for simple typos. Column position is not used. See `Header Combinations - Accepted Aliases.md`.


## v3 source-format fix

The classifier now recognizes wide files that do not contain an Employee ID
column, including Employee Name + Employee Date of Birth + Spouse/Child/
Father/Mother fields. v12's wide-to-long engine generates the employee
identifier for such rows and preserves the header-driven transformation.
