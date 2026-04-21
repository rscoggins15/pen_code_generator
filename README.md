# PEN Table Generator

A Python application that generates a PEN TABLE file from Excel Index of Sheets (IOS) data.

## Download
The program for Windows is in the dist folder. Download the .exe file and run it. 

## Overview

This program reads project information from an Excel file and generates formatted PEN table (.tbl) with dynamic variable substitution. It provides a user-friendly GUI for entering project details and automatically processes sheet data with support for description ranges (THRU patterns) and abbreviations.

## Usage

You will need the Index Of Sheets excel file in addition to the executable. In this file, there are three named columns contained in row A (which is hidden). These enable the program to read the Index Of Sheets. Please use the included Index Of Sheets as a starting template while you create your index of sheets. Otherwise, you will need to add this row into an existing excel sheet. 

## Features

- **File Selection**: Browse and select your Index of Sheets Excel file
- **GUI Input Form**: Easy-to-use interface for entering project information
- **THRU Logic**: Automatically generates intermediate sheet instances for range patterns (e.g., BC(1)-21 THRU BC(12)-21)
- **Description Abbreviations**: Converts long descriptions to short abbreviations (e.g., "SUMMARY OF SMALL SIGNS" → "SOSS")
- **Error Handling**: User-friendly error messages for missing files or invalid data
- **Centered Window**: Clean, professional interface displayed in the center of your screen
- **Keyboard Shortcuts**: 
  - Press **Enter** to submit
  - Press **Escape** to cancel

## Requirements

- Python 3.6+
- pandas
- openpyxl (for Excel support)
- tkinter (usually included with Python)

## Usage

1. Run the program:
```bash
python ios_to_pen.py
```

2. A file browser will appear - select your Excel Index of Sheets file (INDX sheet with columns: sht_num, category, description)

3. A GUI form will appear asking for project information:
   - Control Number
   - Section Number
   - Job Number
   - Highway Name
   - District
   - County
   - Year
   - Funding

4. Enter the required information and click **Submit** (or press Enter)

5. The program will generate a file named `{Highway_Name}_PEN_TABLE.tbl` in the current directory

## Excel File Format

Your Excel file must contain a sheet named "INDX" with the following columns:
- **sht_num**: Number of sheet instances
- **category**: Sheet category/section
- **description**: Sheet description (can include THRU patterns)

### Example Data

| sht_num | category | description |
|---------|----------|-------------|
| 1 | INDEX | SUPPLEMENTAL INDEX OF SHEETS |
| 12 | BARRIERS | BC(1)-21 THRU BC(12)-21 |
| 3 | DETAILS | MISCELLANEOUS DETAILS |

## Features Explained

### THRU Logic
When a description contains " THRU ", the program automatically generates all intermediate instances:
- **Input**: `BC(1)-21 THRU BC(12)-21` with sht_num=12
- **Output**: BC(1)-21 through BC(12)-21 with sequential numbering

### Description Abbreviations
Descriptions are automatically abbreviated:
- SUMMARY OF SMALL SIGNS → SOSS
- CONSTRUCTION SEQUENCE OF WORK → SOW
- SUPPLEMENTAL INDEX OF SHEETS → IOS
- QUANTITY SUMMARY SHEETS → QS
- And many more...

## Output File

The generated .tbl file contains:
- Global settings and configuration
- Project-specific variables (Control, Section, Job, Highway, etc.)
- Categorized sheet listings with dynamic numbering

## Error Messages

- **"File not found"**: The selected file cannot be found
- **"Missing required columns"**: The Excel file lacks sht_num, category, or description columns
- **"Failed to load Excel file"**: General error reading the file (check format)

## Building an Executable

To create a standalone .exe file:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed main.py
```

The executable will be in the `dist` folder.

## Troubleshooting

**Excel file not recognized**
- Ensure the sheet is named exactly "INDX" (case-sensitive)
- Verify columns are named: sht_num, category, description

## Support

For issues or questions, review the error messages which provide specific guidance on what went wrong.

## License

This program is for internal use.
