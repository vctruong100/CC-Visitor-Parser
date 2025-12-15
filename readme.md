
# Planned Start + Patient Parser

A desktop utility that processes an Excel file containing **Planned Start dates** and **Patient names**, optionally deduplicates rows by date and patient, and outputs a cleaned Excel file along with summary statistics.

## ✅ Key Features

- **Single input**:
  - Excel file (`.xlsx`) with columns:
    - `Planned Start` (or similar)
    - `Patient` (or similar)
- **Automatic column detection**:
  - Finds columns by normalized names (e.g., `plannedstart`, `patientname`).
- **Validation**:
  - Removes rows with missing Planned Start or Patient values.
- **Optional deduplication**:
  - Deduplicates by **Planned Start (date-only)** and **Patient**.
- **Summary statistics**:
  - Total rows before and after deduplication
  - Unique pair count
  - Monthly counts of Planned Start dates
- **Outputs**:
  - `planned_result.xlsx` containing:
    - All rows or deduplicated rows
  - Displays summary in GUI

## 🖥 GUI Features

- Browse and select an Excel file.
- Checkbox to enable/disable deduplication.
- Displays:
  - Output file path
  - Unique pair count
  - Rows before deduplication
  - Total rows after processing
  - Monthly counts of Planned Start dates
- Runs in a separate thread for responsiveness.

## 📂 Requirements

- Windows with Python launcher `py`
- Python packages: `pandas`, `openpyxl`

Installed automatically by the provided batch file.

## 🚀 Installation & Launch

1. Place these files in the same folder:
   - `planned_parser.py`
   - `planned_gui.py`
   - `run_planned.bat`
2. Double-click `run_planned.bat`:
   ```bat
   @echo off
   setlocal
   py -m pip install --upgrade pip
   py -m pip install pandas openpyxl
   py planned_gui.py
   endlocal
   ```

## ▶ Using the App

1. **Select Excel file**: Browse and choose your `.xlsx` file.
2. **Deduplication**:
   - Check the box to deduplicate by Planned Start + Patient.
3. Click **Run**.
4. Outputs:
   - `planned_result.xlsx` in the same folder as the input file.
5. GUI displays:
   - Output file path
   - Unique pair count (if deduplication enabled)
   - Rows before deduplication
   - Total rows after processing
   - Monthly counts of Planned Start dates

## 📝 Input File Notes

- Must include columns for Planned Start and Patient.
- Column names are matched case-insensitively and normalized (spaces removed).
- Planned Start must be a valid date; Patient must be non-empty.

## 📊 Output Details

- **Excel file**:
  - Sheet name: `Unique` if deduplicated, else `All`.
- **Summary in GUI**:
  - Unique pairs
  - Rows before deduplication
  - Total rows after processing
  - Monthly counts (e.g., `1/2025: 12`)

## ⚠ Common Issues & Tips

- **Missing columns**: The program will raise an error if Planned Start or Patient columns are not found.
- **Invalid dates**: Rows with invalid Planned Start dates are removed.
- **Large files**: Processing may take a few seconds for large datasets.

## 📂 Folder Structure

```
YourFolder/
├── planned_parser.py
├── planned_gui.py
├── run_planned.bat
└── planned_result.xlsx   # generated after Run
```
