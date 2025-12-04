# Excel Data Cleaner - Windows Executable

A stand-alone Windows executable that automatically cleans Excel files based on predefined rules.

## Features

- **Drag-and-drop support**: Simply drag an Excel file onto the .exe
- **File picker**: Double-click the .exe to browse for a file
- **Safe operation**: Never overwrites the original file
- **Efficient processing**: Handles large Excel files with thousands of rows
- **Clear feedback**: Shows success messages and statistics

## Data Cleaning Rules

The tool removes entire rows based on substring matches (case-insensitive) in specific columns:

### 1. ShipmentID (Column BV)
- Removes rows containing: **"FOC"**

### 2. Order (Column H)
- Removes rows containing any of:
  - "test" or "testing"
  - "M88" (including variants like "M880123")
  - "GB Test", "GB Testing", or "GB" alone

### 3. Buyer PO Number (Column I)
- Removes rows containing:
  - "test" or "testing"
  - "FOC"

### 4. Comment (Column BO)
- Removes rows containing:
  - "FOC"
  - "M88"

## Output

- Creates a new file: `<original_filename>_CLEANED.xlsx`
- Saved in the same directory as the input file
- Displays statistics: original rows, rows removed, remaining rows

---

## Building the Executable

### Prerequisites

1. **Python 3.8 or higher** installed on Windows
2. **pip** (Python package manager)

### Step 1: Install Dependencies

Open Command Prompt or PowerShell in the project directory and run:

```bash
pip install -r requirements.txt
```

This installs:
- `pandas` - Excel file processing
- `openpyxl` - Excel file reading/writing
- `pyinstaller` - Executable creation

### Step 2: Build the Executable

#### Option A: Using the spec file (Recommended)

```bash
pyinstaller excel_cleaner.spec
```

#### Option B: Using command line arguments

```bash
pyinstaller --onefile --windowed --name ExcelCleaner excel_cleaner.py
```

### Step 3: Locate the Executable

After building, find the executable at:

```
dist\ExcelCleaner.exe
```

You can now distribute this single .exe file - no installation required!

---

## How to Use the Executable

### Method 1: Drag and Drop
1. Locate your Excel file (.xlsx)
2. Drag it onto `ExcelCleaner.exe`
3. Wait for the success message
4. Find the cleaned file in the same directory

### Method 2: File Picker
1. Double-click `ExcelCleaner.exe`
2. Browse and select your Excel file
3. Wait for the success message
4. Find the cleaned file in the same directory

---

## How Drag-and-Drop Works

When you drag a file onto a Windows executable:

1. Windows launches the .exe with the file path as a command-line argument
2. The Python script accesses this via `sys.argv[1]`
3. If no argument is provided (double-click), the script opens a file picker dialog using `tkinter`

This is implemented in the `main()` function:

```python
if len(sys.argv) > 1:
    input_file = sys.argv[1]  # File dragged onto .exe
else:
    input_file = select_file()  # Open file picker
```

---

## Error Handling

The tool handles common errors gracefully:

- **Missing columns**: Alerts if required columns (H, I, BO, BV) are missing
- **File locked**: Warns if the Excel file is open in another program
- **Permission errors**: Notifies if unable to write the output file
- **Corrupt files**: Catches and reports file reading errors
- **Large files**: Efficiently processes files with tens of thousands of rows

---

## Technical Details

### Architecture

- **Language**: Python 3
- **GUI Framework**: tkinter (built into Python)
- **Excel Processing**: pandas + openpyxl
- **Packaging**: PyInstaller (creates stand-alone .exe)

### Column Mapping

Excel columns are referenced by letter and converted to 0-based indices:

- Column H (Order) → Index 7
- Column I (Buyer PO Number) → Index 8
- Column BO (Comment) → Index 66
- Column BV (ShipmentID) → Index 73

### Pattern Matching

- **Case-insensitive**: "FOC", "foc", "FoC" all match
- **Substring matching**: "M880123" matches "M88"
- **Efficient**: Uses pandas vectorized operations for speed

---

## Troubleshooting

### Build Issues

**Problem**: `pyinstaller: command not found`
- **Solution**: Ensure Python Scripts directory is in PATH, or use: `python -m PyInstaller excel_cleaner.spec`

**Problem**: Missing module errors during build
- **Solution**: Reinstall dependencies: `pip install -r requirements.txt --force-reinstall`

### Runtime Issues

**Problem**: Executable won't run
- **Solution**: Windows may block downloaded .exe files. Right-click → Properties → Unblock

**Problem**: "Permission denied" error
- **Solution**: Close the Excel file if it's open in Excel or another program

**Problem**: Columns not found
- **Solution**: Verify your Excel file has columns H, I, BO, and BV

---

## File Structure

```
POLine Test Deletion/
├── excel_cleaner.py       # Main Python source code
├── excel_cleaner.spec     # PyInstaller configuration
├── requirements.txt       # Python dependencies
├── README.md             # This file
└── dist/
    └── ExcelCleaner.exe  # Built executable (after building)
```

---

## Customization

### Changing Cleaning Rules

Edit `excel_cleaner.py` in the `clean_data()` method:

```python
# Add new patterns to existing rules
order_patterns = ['test', 'testing', 'M88', 'GB', 'YOUR_PATTERN']

# Or add new column rules
new_col_idx = self.column_letter_to_index('Z')  # Column Z
if new_col_idx < len(self.df.columns):
    new_col = self.df.iloc[:, new_col_idx]
    keep_mask &= ~new_col.apply(lambda x: self.contains_pattern(x, ['PATTERN']))
```

### Adding an Icon

1. Create or download a .ico file
2. Edit `excel_cleaner.spec`:
   ```python
   icon='path/to/your/icon.ico'
   ```
3. Rebuild with `pyinstaller excel_cleaner.spec`

---

## License

This tool is provided as-is for data cleaning purposes.

---

## Support

For issues or questions:
1. Check the Troubleshooting section
2. Verify your Excel file format matches requirements
3. Ensure all required columns exist in your file
