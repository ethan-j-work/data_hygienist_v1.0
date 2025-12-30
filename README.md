Overview

This workbook is a general-purpose data hygiene tool for cleaning, standardizing, and reviewing Excel datasets. It helps identify duplicate records, normalize text, fill empty cells, and surface basic data quality metrics without making destructive changes to the underlying data. I built this after working on multiple data migration and reporting projects where progress stalled because source files contained thousands of small issues. Extra spaces, inconsistent casing, duplicate rows, and incomplete records often took longer to fix than the actual analysis or import work. This tool is meant to reduce that friction and make basic cleanup faster and more repeatable.

How to Use It

Paste your raw dataset into the Data Input sheet. Headers should be in row 1. The data can be any size or shape, and no pre-formatting is required. On the Control sheet, you can optionally define primary identifier columns by entering comma-separated column letters, such as a,b,c. These columns are used to determine row uniqueness when identifying duplicates. If no columns are specified, the tool treats the entire row as the identifier. When checking for duplicates, only rows found after the first occurrence are flagged. The comparison is case-insensitive, so values like ETHAN, ethan, and Ethan are treated as the same. This avoids Excel’s default behavior of highlighting every repeated value and instead focuses on identifying true duplicate records that typically need review or removal. Next, select your formatting options. You can standardize text case, fill empty cells with N/D, and optionally remove duplicates. When duplicates are removed, they are written to a separate Duplicates sheet rather than deleted outright. Click the big green Process button to run the cleaning cycle. The cleaned dataset is written to the Control sheet starting in column E. During processing, the tool calculates summary metrics such as total rows and columns processed, processing time, duplicate and unique row counts, rows with blank identifiers, and the number of whitespace issues corrected.

Reviewing and Exporting Results
After processing, review the Data Cleanliness section on the Control sheet to understand what changed and how clean the dataset is. To export the cleaned data, click Export Clean Data. This creates a new workbook containing only the cleaned dataset and opens Excel’s standard Save As dialog. You can save the file in any format Excel supports. The Control panel, metrics, and original input data are not included in the export.

Clearing Data
Use Clear Output to remove cleaned results and reset metrics while leaving the original Data Input unchanged.Use Clear All to reset everything, including options and the Data Input sheet. Clear All requires confirmation before running.

Notes
This tool intentionally avoids automatic type conversion, date parsing, schema enforcement, or heuristic guessing. It does not attempt to interpret the meaning of your data. All changes are explicit, measurable, and reversible. The goal is to prepare data for analysis or import, not to validate business rules or apply domain-specific logic.
The workbook is optimized for large datasets and has been tested on files exceeding 150,000 rows. In testing, a dataset of that size processed in roughly 15 seconds, though performance will vary depending on hardware. All logic is written in pure VBA for compatibility with both Windows and Mac versions of Excel.

Intended Workflow
Paste data, configure options, run the cleaning cycle, review the results, and export a clean dataset for downstream use. The focus is on predictable behavior, transparency, and minimizing risk during early-stage data preparation.
<img width="3522" height="888" alt="image" src="https://github.com/user-attachments/assets/7448bedf-0741-49b7-af1b-e0738db16250" />
