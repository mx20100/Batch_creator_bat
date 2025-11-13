# AM-Flow Converter

A lightweight Windows application that converts AM-Flow Excel batch sheets into properly structured ZIP packages for uploading to AM-Vision.

This tool is built for internal AM-Flow use. No installation required.

---

## Download (EXE)

You can download the latest version from the Releases page:

→ Go to the **Releases** tab on the right side of this GitHub repository  
and download **AM-Flow_Converter.exe**

---

## What the Converter Does

### 1. Reads your Excel (.xlsx or .xlsm)
- Detects the first non-empty visible sheet  
- Removes empty sheets automatically  
- Converts it to meta.csv (in memory)

### 2. Validates and fixes meta.csv
- Checks all required columns  
- Fixes empty or zero copies  
- Adds missing .stl extensions  
- Sanitizes filenames using only the following characters (A–Z, 0–9, _.)  
- Everything stays in RAM, no temp files created in the batch folder

### 3. Scans for STL files
- Reads STL files in the root folder  
- Reads STL files in material-named subfolders  
- Supports mixed structures (root + subfolders)  
- Ignores unrelated folders automatically

### 4. Creates ZIP archives
- Always outputs ZIPs to the **root folder**  
- Automatically splits when a ZIP exceeds **900 MB**  
- Uses ZIP_STORED for maximum speed  
- Typical performance on an NVMe drive: **600–750 MB/s**  
- Adds meta.csv to every archive
- Splits meta.csv according to which stl files are in which zip folder

### 5. Uses RAM for temporary work
- Excel cleanup happens in OS temp directory  
- No temporary files or folders visible
- Fast, clean, reliable

### 6. Real-time GUI
- Live log of what the converter is doing  
- Status labels (Excel → Validation → Packaging → Cleanup)  
- Auto-scrolling text area  

---

## Supported Folder Structures

### 1. All STLs in the root folder
```
/BatchFolder
20250101_01.xlsx
part1.stl
part2.stl
part3.stl
```
### 2. STLs inside material-named subfolders
```
/BatchFolder
20250101_01.xlsx
/PA12_BLUE
part1.stl
part2.stl
/TPU
fileA.stl
fileB.stl
```
### 3. Mixed root + subfolders (supported)
```
/BatchFolder
20250101_01.xlsx
loosePart1.stl
loosePart2.stl
/PA12_BLUE
  p1.stl
  p2.stl
```
### 4. Multiple material folders
```
/BatchFolder
20250101_01.xlsx
/PA12_BLUE
/PA11_BLACK
/TPU
```

### 5. Unrelated folders are ignored
Only folders containing STL files are considered.

---

## How to Use

1. Place your Excel sheet and STL files together in a folder.
2. Run **AM-Flow Converter.exe** inside that folder.
3. The converter:
   - Reads and validates your Excel
   - Builds meta.csv in RAM
   - Scans all STLs (root + subfolders)
   - Creates the ZIP files
   - Places all ZIPs in the root folder

You don't need to configure anything. Just run it.

---

## Highlighted Features

| Feature | Description |
|--------|-------------|
| Very fast ZIP creation | RAM + ZIP_STORED = high throughput |
| 900 MB safety cap | Prevents upload errors |
| RAM-only temp handling | No leftover files |
| Supports all folder layouts | Seamless operation |
| In-memory meta.csv | No file pollution |
| Real-time UI | See progress clearly |
| One-click workflow | Nothing to configure |

---

## Example Output
```
20250101_01_PA12_BLUE_part1.zip
20250101_01_PA12_BLUE_part2.zip
20250101_01_root_part1.zip
```

---

## Issues or Suggestions
If something doesn't work or you'd like improvements, contact Max or open an issue on this repository.