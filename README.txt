*** FDA_TABULATION — MACOS APP ***

A small macOS app that builds an Excel workbook stratified by TREATMENT GROUP and summarizes TP53 mutation status.

*** WHAT’S NEW ***

Only subjects present in the INPUT are included anywhere. Subjects found only in the dose dictionary are ignored (also for TP53 counts).

Per-cohort sheets now show: SUBJECT, VISIT, BMBLASTS, CD123, CD123+ BLASTS (%).
PB blasts are no longer reported.

CD123+ BLASTS (%) is computed as BMBLASTS * (CD123 / 100), parsing values like 78.5 or 78.5%.

Leading zeros preserved: purely numeric SUBJECTs are normalized to 8 digits (e.g., 102018 → 00102018). The column is written as text so Excel won’t drop zeros.

Optional password dialog: you can protect all sheets and workbook structure.

Added one TP53 matrix sheet per cohort (“TP53 muts – <cohort>”) that shows unique “DNA | PROTEIN” combinations (de-duplicated across DNA1/PROT1, DNA2/PROT2, and any higher-numbered pairs).

*** WHAT IT DOES ***

Lets you pick:

an INPUT DATA file (per-visit measurements),

a PATIENT DOSE DICTIONARY (maps subjects to cohorts / treatment groups),

an OUTPUT FOLDER and an OUTPUT FILE NAME.

Drops rows with blank “Treatment Group” in the dose dictionary before any grouping.

Builds one sheet per Treatment Group with columns:
SUBJECT, VISIT, BMBLASTS, CD123, CD123+ BLASTS (%).

Adds a “TP53 mutation” summary sheet with: Cohort, N patients, TP53 mutated, Percentage.

Adds “TP53 muts – <cohort>” sheets: rows = patients in that cohort (from INPUT only); columns = unique “DNA | PROTEIN” labels; cells = “X” if the combination is present.

*** SUPPORTED FILE FORMATS ***

CSV (.csv), TSV (.tsv, .tab), EXCEL (.xlsx, .xls), JSON (.json)

*** REQUIRED COLUMNS (HEADER MATCHING RULES) ***
Headers are matched case-insensitively, trimming spaces (including non-breaking spaces) and normalizing internal whitespace.

PATIENT DOSE DICTIONARY (COHORTS)

Patient Number

Treatment Group

INPUT DATA (PER SUBJECT / PER VISIT)

SUBJECT

VISIT

BMBLASTS (or BM BLASTS)

CD123

TP53FINALRESULT (values like positive / negative)

If a required column is missing, the app shows an error with the columns it found.

*** IMPORTANT LOGIC ***

Only INPUT subjects are considered. The dose dictionary serves to map cohorts, but any subject not appearing in INPUT is discarded everywhere (cohort sheets, TP53 summary, TP53 matrices).

Cohort names are taken after trimming blanks. (Case differences produce distinct cohorts.)

TP53 summary counts unique subjects per cohort and how many have any INPUT row with TP53FINALRESULT == "positive".

SUBJECT is written as text and numeric IDs are padded to 8 digits.

*** OUTPUT DETAILS ***
PER-COHORT SHEETS

One sheet per unique, non-blank Treatment Group.

Columns: SUBJECT, VISIT, BMBLASTS, CD123, CD123+ BLASTS (%).

Sorted by SUBJECT, then VISIT.

TP53 MUTATION (SUMMARY)

Cohort: Treatment Group name

N patients: unique subjects in that cohort (INPUT-only)

TP53 mutated: count with any positive in TP53FINALRESULT

Percentage: (TP53 mutated / N patients) * 100, rounded to 2 decimals

TP53 MUTATION MATRICES (PER COHORT)

Sheet name: “TP53 muts – <cohort>”

Rows: INPUT subjects in that cohort

Columns: unique labels “<DNA> | <PROTEIN>” gathered across TP53DNA1/PROTEIN1, TP53DNA2/PROTEIN2, … (any numbered pairs)

Cells: “X” if present (duplicates collapsed)

*** PASSWORD PROTECTION ***

Optional dialog to protect all sheets and lock workbook structure with your chosen password.

*** INSTALL / LAUNCH ***

Unzip FDA_tabulation-mac-arm64.zip.

(Optional) Move FDA_tabulation.app to Applications.

First launch may be blocked by Gatekeeper:

Right-click → Open → Open (one-time), or

Terminal: xattr -dr com.apple.quarantine "/path/to/FDA_tabulation.app"

*** USING THE APP (STEP BY STEP) ***

Double-click FDA_tabulation.app.

Browse… for the INPUT file.

Browse… for the DOSE DICTIONARY file.

Choose… the OUTPUT FOLDER.

Enter the OUTPUT FILE NAME (e.g., dose_groups.xlsx).

Click Run. Optionally set a password. A message will confirm where the Excel was saved.

*** TROUBLESHOOTING ***
APP WON’T OPEN

Use Right-click → Open, or the xattr command above.

MISSING COLUMN ERROR

Check headers in both files. Ensure the first row contains headers and that names match the list above (case/space-insensitive).

NO PATIENTS IN A COHORT

Expected if subjects are in the dictionary but missing in INPUT. Only INPUT subjects are included.

*** DATA PRIVACY ***

All processing is local. The app reads/writes files on your machine only; nothing is uploaded.

*** SUPPORT / CONTACT ***

Include the exact error text, a small de-identified sample of both input files (few rows), and your macOS version.

Maintainer: Renato Giliberti / rgiliberti@menarini-ricerche.it

*** VERSION ***

App: FDA_tabulation.app
