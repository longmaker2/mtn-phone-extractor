# MTN Phone Number Extractor

This Python script extracts unique phone numbers from CRBT MIS report Excel files with advanced deduplication and verification features.

## Features

- Processes single Excel files or entire directories
- Extracts unique phone numbers from MSISDN column
- Handles both .xlsx and .xls files
- **Advanced deduplication**: Removes duplicates within files and across multiple files
- **Verification system**: Confirms no duplicates in final output
- **Detailed statistics**: Shows deduplication rates and processing summary
- Outputs clean CSV file with single column "Phone Number"
- CLI interface for easy automation
- Progress tracking and comprehensive error handling
- Robust data cleaning (removes .0 suffixes, handles NaN values)

## Requirements

- Python 3.7+
- pandas
- openpyxl
- xlrd

## Installation

1. Install required packages:

```bash
pip install pandas openpyxl xlrd
```

## Usage

### Process all Excel files in CRBT_MIS_REPORT directory:

```bash
python extract_phones.py CRBT_MIS_REPORT --output all_unique_phones.csv
```

### Process a single Excel file:

```bash
python extract_phones.py path/to/single_file.xlsx --output output.csv
```

### Use default output filename:

```bash
python extract_phones.py CRBT_MIS_REPORT
# Creates unique_phone_numbers.csv
```

## Input Format

The script expects Excel files with:

- A sheet named "Sheet0"
- A column named "MSISDN" containing phone numbers

## Output Format

The output CSV file contains:

- Single column named "Phone Number"
- Unique phone numbers (duplicates removed across all files)
- Sorted in ascending order
- No duplicate entries guaranteed

## Example Output

```csv
Phone Number
21192*******
21192*******
21192*******
...
```

## Sample Run Output

```
🔍 Found 100 Excel files to process
📊 Extracted 4556 phone numbers from file1.xlsx
📊 Extracted 2795 phone numbers from file2.xlsx
...
✅ Verification: No duplicates in final output
📊 DEDUPLICATION SUMMARY:
   Files processed: 100
   Total raw entries: 487,234
   Duplicates removed: 284,968
   Deduplication rate: 58.5%
✅ Final unique phone numbers: 203,266
✅ Saved to: all_unique_phones.csv
```

## Deduplication Features

The script implements multi-level deduplication:

1. **Within-file deduplication**: Removes duplicates within each Excel file
2. **Cross-file deduplication**: Removes duplicates across multiple files
3. **Data cleaning**: Handles numeric formatting issues (removes .0 suffixes)
4. **Verification**: Double-checks final output for any remaining duplicates
5. **Statistics**: Provides detailed deduplication metrics

## Error Handling

The script handles:

- Missing or corrupted Excel files
- Missing MSISDN columns
- Different Excel formats (.xlsx, .xls)
- Files with different sheet structures
- Numeric vs string phone number formats
- Empty or invalid data entries

## Performance

- Efficiently processes large datasets (tested with 100+ files)
- Memory-efficient using Python sets for deduplication
- Progress tracking for long-running operations
- Graceful error handling without stopping the entire process
