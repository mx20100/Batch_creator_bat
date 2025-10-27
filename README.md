# AM-Flow Converter

A self-contained Windows batch script for converting Excel files to validated CSV and packaging STL data for AM-Flow modules.
This also gives me a better understanding of git and cli commands.

## Features
- Supports `.xlsx` and `.xlsm`
- Automatically installs `xlsx2csv`
- Validates required columns
- Fixes `copies=0` → `1`
- Finds STL folders and zips everything
- Cleans up temporary files and logs

## Requirements
- Windows 10/11
- Python 3.10+ installed and added to PATH  
  (the script will guide users if it’s missing)

## Usage
1. Place your Excel file and STL folder in the same directory as `converter.bat`.
2. Double-click the batch file.
3. A `.zip` with the same name as your Excel file will be created.

## Validation
Each Excel file must contain the following columns (in order):
batch, filename, material, part_id, copies, next_step, order_id, technology

Rows with empty required fields will fail validation.

## Notes
- The script deletes all temporary files automatically.
- To keep logs, comment out the `del "%logfile%"` line at the end.
