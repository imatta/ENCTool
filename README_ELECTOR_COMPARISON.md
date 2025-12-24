# Elector Name Duplicate Finder

A Python tool to compare elector names between two Excel sheets (`2025_LIST` and `2002_LIST`) and identify duplicates using fuzzy matching. This tool handles the challenge of matching names when direct string comparison fails due to transliteration differences between English and Telugu (Vernacular) names.

## Problem Statement

When comparing elector names between two lists:
- **Direct string matching fails** because:
  - Vernacular (Telugu) names translated to English don't match exactly
  - English names translated to Telugu don't match exactly
  - There may be spelling variations, spacing differences, or transliteration inconsistencies

## Solution

This tool uses **fuzzy matching algorithms** to:
- Compare names across both English and Vernacular columns
- Handle transliteration differences
- Find matches even when exact strings don't match
- Provide similarity scores for each match
- Export detailed results to Excel

## Features

- ✅ **Interactive file upload prompt** - Prompts for Excel file each time you run the script
- ✅ **Fuzzy matching** - Uses token-based similarity matching to handle variations
- ✅ **Cross-column comparison** - Compares English-to-English, Vernacular-to-Vernacular, and cross-language matches
- ✅ **Configurable similarity threshold** - Adjustable matching sensitivity (default: 85%)
- ✅ **Comprehensive reporting** - Detailed Excel output with summary and duplicate lists
- ✅ **Logging** - Detailed logs for troubleshooting
- ✅ **Handles missing data** - Gracefully handles empty or missing name fields

## Requirements

### Python Version
- Python 3.7 or higher

### Dependencies
Install the required packages:

```bash
pip install -r requirements_elector_comparison.txt
```

Or install individually:
```bash
pip install pandas openpyxl rapidfuzz
```

**Note:** `rapidfuzz` is recommended for better performance. If unavailable, the script will attempt to use `fuzzywuzzy` as a fallback.

## Excel File Format

Your Excel file must contain two sheets with the following structure:

### Sheet 1: `2025_LIST`
| Elector's Name | Elector's Name(Vernacular) | ... (other columns) |
|----------------|---------------------------|---------------------|
| John Doe       | జాన్ డో                | ...                 |
| ...            | ...                       | ...                 |

### Sheet 2: `2002_LIST`
| Elector's Name | Elector's Name(Vernacular) | ... (other columns) |
|----------------|---------------------------|---------------------|
| John Doe       | జాన్ డో                | ...                 |
| ...            | ...                       | ...                 |

**Required Columns:**
- `Elector's Name` - English name
- `Elector's Name(Vernacular)` - Telugu name

## Usage

### Basic Usage

1. Run the script:
   ```bash
   python elector_name_comparison.py
   ```

2. When prompted, provide the path to your Excel file:
   ```
   Please enter the path to your Excel file (or drag and drop the file here): 
   C:\path\to\your\file.xlsx
   ```

3. Optionally set the similarity threshold (default: 85):
   ```
   Enter similarity threshold (0-100, default 85): 85
   ```

4. The script will:
   - Load both sheets
   - Compare all names
   - Display a summary
   - Export results to a new Excel file

### Similarity Threshold

The similarity threshold determines how similar two names must be to be considered duplicates:
- **100%** - Only exact matches
- **90-99%** - Very similar names (recommended for strict matching)
- **85-89%** - Similar names with minor variations (default)
- **70-84%** - More lenient matching (may include false positives)
- **<70%** - Very lenient (likely to include many false positives)

**Recommendation:** Start with 85% and adjust based on your results.

## Output

The script generates two output files:

### 1. Excel Report (`<filename>_duplicates_<timestamp>.xlsx`)

Contains two sheets:

#### Sheet 1: `Summary`
| Metric | Value |
|--------|-------|
| Total records in 2025_LIST | 1000 |
| Total records in 2002_LIST | 950 |
| Total duplicates found | 850 |
| Exact matches (100% similarity) | 800 |
| Fuzzy matches (85-99% similarity) | 50 |
| No matches found | 150 |
| Similarity threshold used | 85% |
| Analysis date | 2025-01-15 10:30:00 |

#### Sheet 2: `Duplicates`
| similarity_score | match_type | is_exact_match | 2025_english | 2025_vernacular | 2025_index | 2002_english | 2002_vernacular | 2002_index |
|------------------|------------|----------------|--------------|-----------------|------------|--------------|-----------------|------------|
| 100.0 | English-English | True | John Doe | జాన్ డో | 5 | John Doe | జాన్ డో | 12 |
| 92.5 | Vernacular-Vernacular | False | John Doe | జాన్ డో | 8 | John Doe | జాన్ డో | 25 |
| ... | ... | ... | ... | ... | ... | ... | ... | ... |

**Columns:**
- `similarity_score` - Match confidence (0-100)
- `match_type` - How the match was found (English-English, Vernacular-Vernacular, etc.)
- `is_exact_match` - True if 100% match
- `2025_*` - Data from 2025_LIST sheet
- `2002_*` - Data from 2002_LIST sheet
- `*_index` - Original row index in respective sheet

### 2. Log File (`elector_comparison.log`)

Detailed execution log with:
- File loading status
- Data validation results
- Comparison progress
- Errors and warnings

## How It Works

### Matching Algorithm

1. **Normalization**: Names are normalized (lowercase, trimmed, extra spaces removed)

2. **Multi-strategy Comparison**:
   - English name from 2025_LIST vs English names in 2002_LIST
   - English name from 2025_LIST vs Vernacular names in 2002_LIST
   - Vernacular name from 2025_LIST vs Vernacular names in 2002_LIST
   - Vernacular name from 2025_LIST vs English names in 2002_LIST

3. **Fuzzy Matching**: Uses token-based similarity (token sort ratio) which:
   - Splits names into tokens (words)
   - Sorts tokens alphabetically
   - Compares sorted token sequences
   - This handles word order differences and spacing variations

4. **Best Match Selection**: For each name in 2025_LIST, finds the best matching name in 2002_LIST

5. **Threshold Filtering**: Only matches above the similarity threshold are considered duplicates

### Example Matching Scenarios

| Scenario | Example | Why It Works |
|----------|---------|--------------|
| Exact match | "John Doe" = "John Doe" | 100% similarity |
| Spacing difference | "JohnDoe" ≈ "John Doe" | Token matching handles spacing |
| Word order | "Doe John" ≈ "John Doe" | Token sort handles order |
| Transliteration | "John" ≈ "జాన్" | Cross-language matching |
| Minor spelling | "Jon Doe" ≈ "John Doe" | Fuzzy matching handles typos |

## Troubleshooting

### Common Issues

#### 1. "Sheet '2025_LIST' not found"
- **Solution**: Ensure your Excel file has sheets named exactly `2025_LIST` and `2002_LIST` (case-sensitive)

#### 2. "Column 'Elector's Name' not found"
- **Solution**: Check that column names match exactly:
  - `Elector's Name`
  - `Elector's Name(Vernacular)`
- Note: Column names are case-sensitive

#### 3. "No matches found" or "Too many matches"
- **Solution**: Adjust the similarity threshold:
  - Too strict (few matches): Lower threshold (e.g., 80%)
  - Too lenient (many false positives): Raise threshold (e.g., 90%)

#### 4. Performance issues with large files
- **Solution**: The script processes all combinations. For very large files (>10,000 rows each), consider:
  - Filtering data before comparison
  - Using a higher threshold to reduce processing
  - Running on a machine with more RAM

#### 5. Unicode/Encoding errors
- **Solution**: Ensure your Excel file is saved with UTF-8 encoding. If issues persist:
  - Re-save the Excel file
  - Ensure Telugu fonts are properly installed

## Performance

- **Small files** (<1,000 rows each): < 1 minute
- **Medium files** (1,000-5,000 rows each): 1-5 minutes
- **Large files** (>5,000 rows each): 5-15 minutes

Performance depends on:
- Number of records
- Similarity threshold (lower = more comparisons)
- System resources

## Limitations

1. **One-to-one matching**: Each name in 2025_LIST is matched to the best match in 2002_LIST. If multiple names in 2002_LIST match equally well, only the first best match is reported.

2. **No reverse matching**: The script matches 2025_LIST → 2002_LIST. Names in 2002_LIST that don't match any in 2025_LIST won't appear in results.

3. **Transliteration accuracy**: Fuzzy matching helps but may not catch all transliteration variations. Manual review is recommended for critical matches.

4. **Memory usage**: Very large files may require significant RAM.

## Future Enhancements

Potential improvements:
- [ ] Bidirectional matching (2025→2002 and 2002→2025)
- [ ] Multi-match reporting (show all matches above threshold)
- [ ] Phonetic matching for better transliteration handling
- [ ] GUI interface for easier file selection
- [ ] Batch processing for multiple files
- [ ] Integration with translation APIs for better cross-language matching

## Support

For issues or questions:
1. Check the log file: `elector_comparison.log`
2. Review this documentation
3. Verify Excel file format matches requirements
4. Try adjusting the similarity threshold

## License

This tool is provided as-is for internal use.

## Changelog

### Version 1.0 (2025-01-15)
- Initial release
- Fuzzy matching implementation
- Excel export functionality
- Interactive file upload prompt
- Comprehensive documentation


