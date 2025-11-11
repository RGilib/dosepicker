*** FDA_TABULATION — MACOS APP ***

A small macOS app that creates an Excel workbook stratified by TREATMENT GROUP and summarizes TP53 mutation status.

*** WHAT IT DOES ***

- Lets you pick:
  1) an INPUT DATA file (per-visit measurements),
  2) a PATIENT DOSE DICTIONARY (mapping patients to cohorts),
  3) an OUTPUT FOLDER and an OUTPUT FILE NAME.
- Drops rows with BLANK "Treatment Group" in the dose dictionary before any grouping.
- Builds ONE SHEET PER TREATMENT GROUP with columns:
  SUBJECT, VISIT, BMBLASTS, PBBLASTS, CD123
- Adds a sheet "TP53 mutation" with columns:
  Cohort, N patients, TP53 mutated, Percentage

*** SYSTEM REQUIREMENTS ***

- macOS 12 or newer on Apple Silicon (M-series). (If you received an Intel build, it will run on Apple Silicon via Rosetta.)
- No Python installation required; the app is self-contained.

*** INSTALL / LAUNCH ***

1. Unzip "FDA_tabulation-mac-arm64.zip".
2. (Optional) Move "FDA_tabulation.app" to Applications.
3. First launch might be blocked by Gatekeeper:
   - Right-click the app -> Open -> Open (one-time), OR
   - Terminal: xattr -dr com.apple.quarantine "/path/to/FDA_tabulation.app"

*** USING THE APP (STEP BY STEP) ***

1. Double-click "FDA_tabulation.app".
2. Click "Browse…" next to "Input data file" and select your INPUT file.
3. Click "Browse…" next to "Patient dose dictionary" and select the DOSE DICTIONARY file.
4. Click "Choose…" to select the OUTPUT FOLDER.
5. Enter the OUTPUT FILE NAME (e.g., dose_groups.xlsx).
6. Click "Run". When finished, a message shows where the Excel file was saved.

*** SUPPORTED FILE FORMATS ***

- CSV (.csv), TSV (.tsv, .tab), EXCEL (.xlsx, .xls), JSON (.json)

*** REQUIRED COLUMNS (HEADER MATCHING RULES) ***

Headers are matched CASE-INSENSITIVELY and with LEADING/TRAILING SPACES REMOVED. Inner whitespace is also normalized.

PATIENT DOSE DICTIONARY (COHORTS):
- "Patient Number"
- "Treatment Group"

INPUT DATA (PER SUBJECT / PER VISIT):
- "SUBJECT"
- "VISIT"
- "BMBLASTS" (or "BM BLASTS")
- "PBBLASTS" (or "PB BLASTS")
- "CD123"
- "TP53FINALRESULT" with values such as "positive" or "negative"

If a required column is missing, the app will show an error listing the columns it found.

*** OUTPUT DETAILS ***

PER-COHORT SHEETS:
- One worksheet per unique, non-blank TREATMENT GROUP from the dose dictionary.
- Each sheet includes ALL rows from the input that match SUBJECTs in that cohort.
- Columns: SUBJECT, VISIT, BMBLASTS, PBBLASTS, CD123.
- If a subject exists in the dose dictionary but has NO rows in the input, the subject is added with blank values so every cohort member is visible.
- Excel sheet names are sanitized to comply with Excel (max 31 chars; characters : \ / ? * [ ] replaced). Duplicates get "(2)", "(3)", etc.

TP53 MUTATION SHEET:
- "Cohort": the cohort name (Treatment Group).
- "N patients": UNIQUE SUBJECT count for that cohort (from the dose dictionary).
- "TP53 mutated": number of subjects with ANY input row where TP53FINALRESULT == "positive" (case/space tolerant).
- "Percentage": (TP53 mutated / N patients) * 100, rounded to 2 decimals.
- Subjects absent from the input are NOT counted as mutated (no positive evidence).

*** TYPICAL WORKFLOW TIPS ***

- Different spellings like "BM BLASTS" are accepted automatically.
- Cohort names are used AS WRITTEN after trimming. "Arm A" and "arm a" are considered different cohorts. Standardize upstream if you want them merged.
- Ensure SUBJECT IDs in the input match PATIENT NUMBER IDs in the dose dictionary (spacing is trimmed, but different IDs will not match).

*** TROUBLESHOOTING ***

APP WON’T OPEN:
- Use Right-click -> Open, or the xattr command shown above.

MISSING COLUMN ERROR:
- Check headers in BOTH files. Make sure headers are truly in the first row and spelled as listed (case/space doesn’t matter, but the words do).

EMPTY OUTPUT OR MISSING SUBJECTS:
- Verify that SUBJECT (input) and Patient Number (dictionary) actually refer to the same identifiers.

LARGE FILES:
- Processing is in memory (pandas). For very large datasets, close other applications to free RAM.

*** DATA PRIVACY ***

- All processing is local. The app reads/writes files on your machine only; nothing is uploaded.

*** SUPPORT / CONTACT ***

- Include the exact error text, a small de-identified sample of both input files (few rows), and your macOS version.
- Maintainer: Renato Giliberti / rgiliberti@menarini-ricerche.it

*** VERSION ***

- App: FDA_tabulation.app
- Version: 1.0
