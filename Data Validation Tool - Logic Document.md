# Need Correction & Correction Tool — Business Logic Document

This document records the rules reverse-engineered from the four supplied client Raw → Final pairs: Tetris Innovations, Bambrew, SunJewels and Texan Labs.

See the DOCX in this package for the complete detailed version.

## Key transformations
- Tetris: wide → long; Self + Spouse + Child; SI 500000; parents excluded.
- Bambrew: wide → long; Self + Spouse + Child + Mother/Father as Parent; in-laws excluded; SI 300000; child limit 4.
- SunJewels: long → long; Mother/Father → Parent; Child 1/2 → Child; preserve SI.
- Texan Labs: long → long; Dependent Child1/2 → Child; preserve SI.

## Important limitation
Only four client pairs were supplied. A fifth pair was not included, so no unsupported client-specific rule was invented.

### Current v5 changes
- Duplicate detection is intentionally disabled/removed.
- Individual Error_* columns are retained; only the combined Error column is removed.
- DOB normalization is not treated as a logged correction. It is an internal standardization used for Corrected Data and age calculation.
- The Corrections field in Corrected Data lists corrected field names only; the separate Corrections sheet retains audit details.


### Output simplification
The individual Error_* columns and Correction Confidence have been removed from the user-facing Corrected Data output.


## v9.0 — Naming update
- Tool name: **📊 Need Correction & Correction Tool**
- Original source sheets are preserved as **<Original Sheet Name> (Raw Data)**.
- Corrected Data remains separate and does not overwrite the raw sheet.


### v11 — Smart wide-format parsing
The wide-format parser does not use fixed column numbers. It detects fields semantically from their headers and dynamically maps spouse, child and parent groups.


## Latest rules

### Header helper
The UI displays a copyable set of recommended semantic headers before upload.
These are recommendations only. Column order is not fixed and the parser
continues to use header meaning.

### Age rule
Age is determined using this priority:
1. If Date of Birth is available and valid, Age is calculated from DOB through the current date.
2. If DOB is unavailable but a valid Age is supplied in the source, the supplied Age is retained.
3. If both DOB and Age are unavailable, the output shows `No age mentioned`.
Age Band is generated only when a numeric Age is available.

### Final Employee ID sorting
Employee ID A-to-Z sorting is the final data transformation in each processed
sheet. It occurs only after wide/long transformation, normalization,
corrections, validation, remarks, and output-column construction are complete.
The Excel header remains the first row.

## Final three rules confirmed

1. **Recommended header helper:** displayed at the top of the UI with a copyable
   tab-separated header row. It is optional; column sequence is not fixed.
2. **Employee ID A-to-Z:** performed as the final data transformation after all
   processing, corrections, validation, remarks and final column construction.
3. **Age:** if DOB is available, calculate current age from DOB; if DOB is absent
   but source Age exists, retain source Age; if both are absent, show
   `No age mentioned`.

## Final type-safety rule
All DOB and age arithmetic is numeric/date-safe; string or missing values cannot reach subtraction operations.

## Final comparison type-safety fix
Age validation now converts Age to numeric before every range comparison.
The explicit `No age mentioned` marker is excluded from numeric validation.
Age Band also converts Age to numeric before comparisons. This prevents both
`str - int` and `int <= str` failures from source Excel values.

## Employee ID final sort
The final output is sorted after all processing using natural alphanumeric
Employee ID order. Numeric IDs therefore sort 1, 2, 3, 10, 11 instead of
lexicographic 1, 10, 11, 2. IDs such as E1, E2, E10 are also ordered naturally.

## Employee ID source-vs-generated display rule
If Employee ID / Employee Code is present in the raw source, the value is
preserved exactly as the source ID and no label is added.
If the tool has to generate the Employee ID because the source contains no
usable Employee ID / Employee Code, the final Corrected Data displays the
generated value as `ID (Given by Tool)`. Natural alphanumeric sorting is
performed on the underlying ID before this display label is added.

## Added validation and normalization rules — generalized

### Spouse gender inference
- If Self gender is Male and Spouse gender is blank, set Spouse gender to Female.
- If Self gender is Female and Spouse gender is blank, set Spouse gender to Male.
- Existing non-blank Spouse gender is never overwritten.
- The inferred value is recorded in the Corrections audit and Remark.

### Standard relationship/member codes
The following source values are treated as generalized aliases and normalized:
- MMBR → Self (Member / Employee)
- Member → Self
- SPSE → Spouse
- UDTR → Child (Daughter)
- Daughter → Child
- SONM → Child (Son)
- Son → Child
- EMOT → Parent (Employee Mother / Mother)
- EFAT → Parent (Employee Father / Father)
- EFLW → Parent (Employee Father-in-law / Father-in-law)
- EMLW → Parent (Employee Mother-in-law / Mother-in-law)

If a relationship value cannot be mapped to the supported hierarchy, the tool
creates a Relationship hierarchy/code validation issue rather than silently
accepting it.

## Updated Parent vs Parent-in-law precedence rule

Any relationship value that clearly contains **law / in-law terminology** is
normalized to **Parent-in-law**, and it must never fall back to Parent.

Examples covered include:
- EFLW → Parent-in-law
- EMLW → Parent-in-law
- Parent-in-law / Parent in law / Parent law
- Father-in-law / Mother-in-law
- Father law / Mother law
- Father-law / Mother-law
- M law / F law
- M-law / F-law
- Employee Father in Law / Employee Mother in Law
- Other equivalent in-law spellings containing `law` or `inlaw`

This rule has precedence over the generic Parent aliases. Normal values such as
Father, Mother, EFAT and EMOT continue to normalize to Parent.

## Sum Insured is now source-driven

The user is no longer asked to enter GHI / Sum Insured on the UI.
The tool reads Sum Insured from the uploaded source workbook.

For wide-format family data, the family-level Sum Insured found on the source
employee row is applied to every generated member of that family (Self,
Spouse, Child, Parent, Parent-in-law, etc.).

For long-format data, if the Sum Insured appears on only one member row, the
value is propagated to the other members with the same Employee ID.

The engine recognizes semantic headers such as `GHI Sum Insured`, `Sum Insured`,
`SI Amount`, `Insured Amount`, `Coverage Amount`, and `Sum Assured`.
Values such as `500000`, `5,00,000`, `₹500000`, and `5 lakh` are normalized to
the numeric amount.

There is no client-specific hardcoded Sum Insured value and no GHI/SI input
field on the user interface.
