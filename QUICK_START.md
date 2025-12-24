# Quick Start Guide - Elector Name Comparison

## Installation (One-time setup)

1. **Install Python dependencies:**
   ```bash
   pip install -r requirements_elector_comparison.txt
   ```

   Or install manually:
   ```bash
   pip install pandas openpyxl rapidfuzz
   ```

## Running the Tool

1. **Open terminal/command prompt** in the project directory

2. **Run the script:**
   ```bash
   python elector_name_comparison.py
   ```

3. **When prompted, provide your Excel file path:**
   - You can type the full path: `C:\Users\YourName\Documents\electors.xlsx`
   - Or drag and drop the file into the terminal (Windows PowerShell/CMD will auto-paste the path)

4. **Set similarity threshold (optional):**
   - Press Enter for default (85%)
   - Or enter a value between 0-100

5. **Wait for processing** - The script will:
   - Load your Excel file
   - Compare all names
   - Show a summary
   - Create an output Excel file

6. **Check the results:**
   - Output file: `<your_filename>_duplicates_<timestamp>.xlsx`
   - Log file: `elector_comparison.log`

## Example Session

```
C:\Users\imatta\Documents\GitHub\NOC-Runbook-Automation> python elector_name_comparison.py

================================================================================
ELECTOR NAME DUPLICATE FINDER
================================================================================

This tool compares elector names between two Excel sheets:
  - 2025_LIST
  - 2002_LIST

Required columns in each sheet:
  - Elector's Name
  - Elector's Name(Vernacular)

--------------------------------------------------------------------------------

Please enter the path to your Excel file (or drag and drop the file here): 
C:\Users\imatta\Documents\electors.xlsx

--------------------------------------------------------------------------------
Enter similarity threshold (0-100, default 85): 85

Comparing names... This may take a few moments...

================================================================================
ELECTOR NAME COMPARISON SUMMARY
================================================================================
Total records in 2025_LIST: 1000
Total records in 2002_LIST: 950

Duplicates found: 850
  - Exact matches (100%): 800
  - Fuzzy matches (≥85%): 50
  - No matches: 150

Similarity threshold: 85%
================================================================================

Results have been exported to: electors_duplicates_20250115_103045.xlsx
```

## Troubleshooting

### "Module not found" error
```bash
pip install pandas openpyxl rapidfuzz
```

### "Sheet not found" error
- Ensure sheets are named exactly: `2025_LIST` and `2002_LIST`
- Check spelling and capitalization

### "Column not found" error
- Ensure columns are named exactly:
  - `Elector's Name`
  - `Elector's Name(Vernacular)`

### File path issues (Windows)
- Use forward slashes: `C:/Users/Name/file.xlsx`
- Or use raw string: `r"C:\Users\Name\file.xlsx"`
- Or drag and drop the file into the terminal

## Need More Help?

See `README_ELECTOR_COMPARISON.md` for detailed documentation.


