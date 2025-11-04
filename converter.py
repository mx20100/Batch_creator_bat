import os
import sys
import csv
import zipfile
import shutil
import logging
from datetime import datetime
import os, sys

# Prevent infinite self-relaunch in PyInstaller or multiprocessing contexts
if os.environ.get("_CONVERTER_RUNNING") == "1":
    # If this process was already launched once, don't rerun main()
    sys.exit(0)
else:
    os.environ["_CONVERTER_RUNNING"] = "1"

try:
    from openpyxl import load_workbook
except ImportError:
    print("Error: 'openpyxl' is not installed.")
    print("Install it with:  pip install openpyxl")
    input("Press Enter to exit...")
    sys.exit(1)


REQUIRED_COLUMNS = ['batch', 'filename', 'material', 'part_id',
                    'copies', 'next_step', 'order_id', 'technology']


def get_working_dir() -> str:
    # Use the current working directory (where the user runs the EXE / script)
    return os.getcwd()


def setup_logger(log_path: str) -> logging.Logger:
    logger = logging.getLogger("converter")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    fh = logging.FileHandler(log_path, encoding="utf-8")
    fh.setLevel(logging.INFO)
    formatter = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s')
    fh.setFormatter(formatter)
    logger.addHandler(fh)

    logger.info("=" * 60)
    logger.info("Converter started")
    return logger


def find_excel_file(workdir: str) -> str | None:
    # First look for .xlsx, then .xlsm
    for ext in (".xlsx", ".xlsm"):
        candidates = [f for f in os.listdir(workdir)
                      if f.lower().endswith(ext)]
        if candidates:
            # Take the first in sorted order for determinism
            candidates.sort()
            return os.path.join(workdir, candidates[0])
    return None


def convert_excel_to_csv(excel_path: str, csv_path: str, logger: logging.Logger) -> None:
    import pandas as pd

    logger.info(f"Converting Excel to CSV: {excel_path} -> {csv_path}")

    wb = load_workbook(excel_path, data_only=True)
    removed_sheets = 0

    # Remove empty sheets
    for sheet in wb.sheetnames[:]:
        ws = wb[sheet]
        nonempty = False
        for row in ws.iter_rows(values_only=True):
            if any(cell not in (None, "", " ") for cell in row):
                nonempty = True
                break
        if not nonempty:
            wb.remove(ws)
            removed_sheets += 1
            logger.info(f"Removed empty sheet: {sheet}")

    if removed_sheets:
        logger.info(f"Removed {removed_sheets} empty sheet(s) before conversion.")
        # Save cleaned workbook to a temporary file so openpyxl can read it again safely
        tmp_excel = os.path.join(os.path.dirname(excel_path), "_cleaned_" + os.path.basename(excel_path))
        wb.save(tmp_excel)
        wb.close()
        wb = load_workbook(tmp_excel, data_only=True, read_only=True)

    # Pick the first visible worksheet
    ws = None
    for sheet in wb.worksheets:
        if sheet.sheet_state == "visible":
            ws = sheet
            break

    if ws is None:
        logger.warning("No visible non-empty worksheets left — falling back to pandas.")
        try:
            df = pd.read_excel(excel_path, engine="openpyxl")
            df.to_csv(csv_path, index=False, encoding="utf-8-sig")
            logger.info("Conversion to meta.csv completed using pandas fallback.")
            return
        except Exception as e:
            logger.error(f"All sheets empty or unreadable: {e}")
            raise ValueError(f"No usable worksheets found in '{os.path.basename(excel_path)}'.")

    logger.info(f"Using worksheet: {ws.title}")

    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        for row in ws.iter_rows(values_only=True):
            writer.writerow(["" if v is None else str(v) for v in row])

    logger.info("Conversion to meta.csv completed.")

def validate_and_fix_meta(csv_path: str, logger: logging.Logger) -> bool:
    """
    Validates meta.csv:
    - Header must match REQUIRED_COLUMNS (in order, case-insensitive)
    - Rows with any data must have all required fields filled
    - copies == '' or '0' -> set to '1'
    - filename without '.stl' (case-insensitive) -> append '.stl'
    Returns True if validation passes, False otherwise.
    """
    logger.info("Validating meta.csv...")

    errors: list[str] = []
    fixed_copies = 0
    fixed_filenames = 0

    # Read
    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        rows = list(reader)
        header = reader.fieldnames

    if header is None:
        print("Validation failed: meta.csv has no header row.")
        logger.error("Validation failed: no header row in meta.csv")
        return False

    normalized_header = [h.strip().lower() for h in header[:len(REQUIRED_COLUMNS)]]
    if normalized_header != REQUIRED_COLUMNS:
        print("Validation failed: header mismatch.")
        print("Found header:", header)
        print("Expected:", REQUIRED_COLUMNS)
        logger.error(f"Header mismatch. Found: {header}, Expected: {REQUIRED_COLUMNS}")
        return False

    # Validate rows
    for i, row in enumerate(rows, start=2):  # Row numbers starting at 2 (after header)
        if any((row.get(k) or "").strip() for k in row.keys()):
            # Check required fields
            missing = [k for k in REQUIRED_COLUMNS
                       if not (row.get(k) or "").strip()]
            if missing:
                msg = f"Row {i}: missing {', '.join(missing)}"
                errors.append(msg)
                logger.error(msg)

            # Fix copies
            val = (row.get("copies") or "").strip()
            if val == "" or val == "0":
                row["copies"] = "1"
                fixed_copies += 1

            # Fix filename extension
            fname = (row.get("filename") or "").strip()
            if fname and not fname.lower().endswith(".stl"):
                row["filename"] = fname + ".stl"
                fixed_filenames += 1

    # Rewrite file if we made fixes
    if fixed_copies or fixed_filenames:
        with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=REQUIRED_COLUMNS)
            writer.writeheader()
            writer.writerows(rows)
        if fixed_copies:
            msg = f"Corrected {fixed_copies} row(s) with copies=0 or empty."
            print(msg)
            logger.info(msg)
        if fixed_filenames:
            msg = f"Corrected {fixed_filenames} row(s) missing .stl in filename."
            print(msg)
            logger.info(msg)

    if errors:
        print("Validation failed:")
        for e in errors:
            print(" ", e)
        logger.error("Validation failed with errors.")
        return False

    print("Validation passed.")
    logger.info("meta.csv passed validation.")
    return True


def find_stl_folder(workdir: str, logger: logging.Logger) -> str | None:
    """
    Find first subfolder containing at least one .stl file.
    """
    for name in sorted(os.listdir(workdir)):
        path = os.path.join(workdir, name)
        if os.path.isdir(path):
            for fname in os.listdir(path):
                if fname.lower().endswith(".stl"):
                    logger.info(f"Found STL folder: {path}")
                    return path
    logger.error("No STL folder found.")
    return None


def zip_stl_folder(stl_folder: str, zip_path: str, logger: logging.Logger) -> None:
    """
    Create a zip archive of all files in the STL folder directly at the root level (no subfolder).
    """
    logger.info(f"Creating zip archive {zip_path} from {stl_folder}")
    base_folder = os.path.abspath(stl_folder)

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(base_folder):
            for filename in files:
                full_path = os.path.join(root, filename)
                # Compute path relative to the STL folder itself, not its parent
                rel_path = os.path.relpath(full_path, base_folder)
                zf.write(full_path, rel_path)

    logger.info("Zip creation completed.")



def main() -> int:
    workdir = get_working_dir()
    log_path = os.path.join(workdir, "converter_log.txt")
    logger = setup_logger(log_path)

    logger.info(f"Working directory: {workdir}")
    print("Converter started.")
    print()

    success = False
    meta_path = os.path.join(workdir, "meta.csv")

    try:
        # Step 1: Find Excel file
        excel_path = find_excel_file(workdir)
        if not excel_path:
            msg = "No Excel file (.xlsx or .xlsm) found in this folder."
            print(msg)
            logger.error(msg)
            return 1

        basename = os.path.splitext(os.path.basename(excel_path))[0]
        print(f"Found Excel file: {os.path.basename(excel_path)}")
        logger.info(f"Found Excel file: {excel_path}")

        # Step 2: Convert Excel to meta.csv
        print("Converting Excel to meta.csv...")
        convert_excel_to_csv(excel_path, meta_path, logger)
        if not os.path.exists(meta_path):
            msg = "Conversion failed: meta.csv was not created."
            print(msg)
            logger.error(msg)
            return 1
        print("Conversion successful.")
        print()

        # Step 3: Validate meta.csv
        print("Validating meta.csv...")
        if not validate_and_fix_meta(meta_path, logger):
            print("Validation failed — see converter_log.txt for details.")
            return 1
        print()

        # Step 4: Find STL folder
        stl_folder = find_stl_folder(workdir, logger)
        if not stl_folder:
            print("No folder with STL files found.")
            if os.path.exists(meta_path):
                os.remove(meta_path)
            return 1
        print(f"STL folder found: {os.path.basename(stl_folder)}")
        print()

        # Step 5: Copy meta.csv into STL folder
        target_meta = os.path.join(stl_folder, "meta.csv")
        shutil.copy2(meta_path, target_meta)
        logger.info(f"Copied meta.csv to {target_meta}")

        # Step 6: Create zip archive
        zip_name = f"{basename}.zip"
        zip_path = os.path.join(workdir, zip_name)
        print(f"Creating archive: {zip_name}...")
        zip_stl_folder(stl_folder, zip_path, logger)

        if not os.path.exists(zip_path):
            msg = "Failed to create ZIP archive."
            print(msg)
            logger.error(msg)
            return 1

        print(f"Created archive: {zip_name}")
        logger.info(f"Created archive: {zip_path}")
        print()

        # Step 7: Cleanup
        if os.path.exists(target_meta):
            os.remove(target_meta)
        if os.path.exists(meta_path):
            os.remove(meta_path)
        print("Cleanup complete.")
        logger.info("Cleanup complete.")
        success = True
        return 0

    except Exception as e:
        logger.exception("Unexpected error occurred.")
        print("An unexpected error occurred:")
        print(repr(e))
        return 1
    finally:
        logger.info("Converter finished.")
        if success:
            # Delete log file on success, as you requested
            try:
                os.remove(log_path)
            except OSError:
                pass


if __name__ == "__main__":
    try:
        exit_code = main()
    except Exception:
        logging.exception("Fatal unhandled error in converter.")
        exit_code = 1
    finally:
        logging.shutdown()

    try:
        if sys.stdin is None or not sys.stdin.isatty():
            import time
            time.sleep(2)
        else:
            input("Press Enter to exit...")
    except Exception:
        pass

    os._exit(exit_code)
