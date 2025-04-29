# Excel Sheet Management Macros

This repository contains two useful Excel VBA macros for automating worksheet management:

---

## 📄 1. DuplicateSheets_ByRating.bas

### Description:
This macro duplicates a template sheet (`Temp`) for every row in the `Data` tab and adjusts formulas (in cells `B2:B8`) to reference the correct row from the `Data` sheet. Each new sheet is then renamed based on the bond **rating** (value in cell `B5`).

### Key Features:
- Retains all formulas in the duplicated sheets.
- Adjusts formula row references like `=Data!$B2` → `=Data!$B3`, etc.
- Automatically renames each sheet using the rating (e.g., `BBB`, `CCC`, `A+`, etc.).
- Prevents overwriting existing sheets by appending a suffix if needed.

---

##  2. DeleteSheets_ByB11Data.bas

### Description:
This macro scans all worksheets in the workbook and deletes any sheets that have **identical data** in the range starting from `B11:C` downward.

### Key Features:
- Ignores sheet names, company names, and metadata (e.g., `B2:B10`).
- Compares only the core data (price and yield history) in `B11:C`.
- Keeps the first unique sheet and deletes all others with identical data.
- Helps clean up redundant or cloned tabs after bulk processing.

---

##  How to Use

1. Open your Excel workbook.
2. Press `Alt + F11` to open the VBA editor.
3. Go to `File → Import File...` and select the `.bas` file(s).
4. Run the macro from `Alt + F8` (choose the macro name and click "Run").

---

## Warning

These macros **permanently delete sheets**. It’s highly recommended to work on a backup file before running them.

---

## 👩‍💻 Author

Developed by Rishika Abrol for managing structured Excel reports with dynamically generated tabs.

---
