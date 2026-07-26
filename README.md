# Excel Document Comparator

A desktop application that compares two Excel reports containing document information across different time periods. The application identifies changes in document revisions and attached files, and generates a detailed report.

## Features

- **Compare two Excel reports** with document metadata (Document Name, Revision, File Names)
- **Track changes** including:
  - Revision updates
  - File additions
  - File removals
  - File modifications
- **Generate detailed output** with:
  - All changed documents with before/after values
  - Specific change remarks for each document
  - Count of unique documents changed
  - Summary statistics
- **User-friendly GUI interface** for easy file selection and report generation

## System Requirements

- Windows 7 or later
- .NET Framework 4.7.2+ (or standalone Python 3.9+)

## Usage

1. Run `ExcelDocumentComparator.exe`
2. Select the **first Excel report** (earlier date - e.g., Nov 10, 2025)
3. Select the **second Excel report** (later date - e.g., July 1, 2026)
4. Specify output file location
5. Click "Compare" to generate the delta report
6. View results in the output Excel file

## Input Excel Format

Both Excel files should have the following columns:
- **Column A**: Document Name
- **Column B**: Revision
- **Column C**: File name (comma-separated for multiple files)

Example:
| Document Name | Revision | File name |
|---|---|---|
| DocA | v1 | file1.pdf, file2.docx |
| DocB | v2 | file3.xlsx |

## Output Excel Format

The output file contains multiple sheets:

### Changes Sheet
- **Document Name**: Name of the changed document
- **Old Revision**: Previous revision number
- **New Revision**: Current revision number
- **Old Files**: Previous file list
- **New Files**: Current file list
- **Remarks**: Detailed description of changes
- **Change Type**: REVISION_CHANGED / FILES_CHANGED / BOTH_CHANGED

Example Remarks:
- "Revision changed from v1 to v2"
- "Files added: file4.pdf"
- "Files removed: file2.docx"
- "Revision changed from v1 to v2; Files added: file4.pdf; Files removed: file2.docx"

### Summary Sheet
- Total unique documents changed
- Total changes detected
- Change breakdown by type (Revision only, Files only, Both)

## Technical Stack

- **Language**: Python 3.9+
- **GUI**: PySimpleGUI
- **Excel Processing**: openpyxl, pandas
- **Packaging**: PyInstaller (for .exe generation)

## Project Structure

```
excel-document-comparator/
├── src/
│   ├── main.py                 # Entry point and GUI
│   ├── comparator.py           # Core comparison logic
│   ├── excel_handler.py        # Excel file reading/writing
│   └── report_generator.py     # Output report generation
├── tests/
│   └── test_comparator.py
├── build_exe.bat               # Build script for Windows .exe
├── requirements.txt            # Python dependencies
└── README.md
```

## Building the Executable

### Prerequisites
- Python 3.9 or later installed
- pip (Python package manager)

### Steps

1. Install dependencies:
   ```
   pip install -r requirements.txt
   ```

2. Run the build script:
   ```
   build_exe.bat
   ```

3. The executable will be generated in the `dist/` folder:
   ```
   dist/ExcelDocumentComparator.exe
   ```

## Development

To run the application in development mode:
```
python src/main.py
```

## License

MIT License