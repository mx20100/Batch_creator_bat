import os
import sys
import csv
import zipfile
import shutil
import logging
from openpyxl import load_workbook
from datetime import datetime

# ------------------------------------------------------------
# Cooperative cancellation support
# ------------------------------------------------------------
CANCEL_FLAG = False

def request_cancel():
    """Called by the GUI to request graceful cancellation."""
    global CANCEL_FLAG
    CANCEL_FLAG = True

def check_cancel():
    """Raises if a cancellation was requested."""
    if CANCEL_FLAG:
        raise RuntimeError("Conversion cancelled by user.")


# ------------------------------------------------------------
# Core configuration
# ------------------------------------------------------------
REQUIRED_COLUMNS = [
    "batch", "filename", "material", "part_id",
    "copies", "next_step", "order_id", "technology"
]

ENCODING = "utf-8"
MAX_ZIP_SIZE_MB = 900
MAX_ZIP_SIZE_BYTES = MAX_ZIP_SIZE_MB * 1024 * 1024


# ------------------------------------------------------------
# Helpers
# ------------------------------------------------------------
def get_working_dir() -> str:
    """Use current working directory."""
    return os.getcwd()

def setup_logger(log_path: str) -> logging.Logger:
    logger = logging.getLogger("converter")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    fh = logging.FileHandler(log_path, encoding=ENCODING)
    fh.setLevel(logging.INFO)
    formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
    fh.setFormatter(formatter)
    logger.addHandler(fh)

    logger.info("=" * 60)
    logger.info("Converter started")
    return logger

# ------------------------------------------------------------
# Excel → CSV
# ------------------------------------------------------------
def convert_excel_to_csv(excel_path: str, csv_path: str, logger: logging.Logger) -> None:
    """Convert first non-empty sheet to CSV."""
    import pandas as pd
    logger.info(f"Converting Excel to CSV: {excel_path} -> {csv_path}")

    wb = load_workbook(excel_path, data_only=True)
    removed = 0
    for sheet in wb.sheetnames[:]:
        ws = wb[sheet]
        if not any(any(cell not in (None, "", " ") for cell in row) for row in ws.iter_rows(values_only=True)):
            wb.remove(ws)
            logger.info(f"Removed empty sheet: {sheet}")
            removed += 1

    if removed:
        tmp = os.path.join(os.path.dirname(excel_path), "_cleaned_" + os.path.basename(excel_path))
        wb.save(tmp)
        wb.close()
        wb = load_workbook(tmp, data_only=True, read_only=True)

    ws = next((s for s in wb.worksheets if s.sheet_state == "visible"), None)
    if ws is None:
        logger.warning("No visible non-empty worksheets — falling back to pandas.")
        df = pd.read_excel(excel_path, engine="openpyxl")
        df.to_csv(csv_path, index=False, encoding=ENCODING)
        logger.info("Conversion to meta.csv completed (pandas fallback).")
        return

    with open(csv_path, "w", newline="", encoding=ENCODING) as f:
        writer = csv.writer(f)
        for row in ws.iter_rows(values_only=True):
            writer.writerow(["" if v is None else str(v) for v in row])

    logger.info("Conversion to meta.csv completed.")


# ------------------------------------------------------------
# CSV validation and sanitization
# ------------------------------------------------------------
def validate_and_fix_meta(csv_path: str, logger: logging.Logger) -> bool:
    logger.info("Validating meta.csv...")
    check_cancel()

    with open(csv_path, newline="", encoding=ENCODING) as f:
        reader = csv.DictReader(f)
        rows = list(reader)
        header = reader.fieldnames

    if header is None:
        logger.error("meta.csv has no header row.")
        return False

    if [h.strip().lower() for h in header[: len(REQUIRED_COLUMNS)]] != REQUIRED_COLUMNS:
        logger.error(f"Header mismatch. Found: {header}, Expected: {REQUIRED_COLUMNS}")
        return False

    fixed_copies, fixed_filenames = 0, 0
    for i, row in enumerate(rows, start=2):
        check_cancel()
        if not any((row.get(k) or "").strip() for k in row.keys()):
            continue

        if (row.get("copies") or "").strip() in ("", "0"):
            row["copies"] = "1"
            fixed_copies += 1

        fname = (row.get("filename") or "").strip()
        if fname and not fname.lower().endswith(".stl"):
            row["filename"] = fname + ".stl"
            fixed_filenames += 1

    if fixed_copies or fixed_filenames:
        with open(csv_path, "w", newline="", encoding=ENCODING) as f:
            writer = csv.DictWriter(f, fieldnames=REQUIRED_COLUMNS)
            writer.writeheader()
            writer.writerows(rows)
        if fixed_copies:
            logger.info(f"Corrected {fixed_copies} copies field(s).")
        if fixed_filenames:
            logger.info(f"Added missing .stl to {fixed_filenames} filename(s).")

    logger.info("meta.csv passed validation.")
    return True


# ------------------------------------------------------------
# ZIP creation with 900 MB cap
# ------------------------------------------------------------
def zip_with_limit(file_list, base_dir, zip_base_name, meta_path, workdir, logger):
    """Split zips when exceeding 900 MB."""
    check_cancel()
    current_index = 1
    current_size = 0
    zip_paths = []
    zf = None

    def start_new_zip():
        nonlocal current_index, current_size, zf
        zip_name = f"{zip_base_name}_part{current_index}.zip"
        zip_path = os.path.join(workdir, zip_name)
        zf = zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED)
        zip_paths.append(zip_path)
        current_size = 0
        logger.info(f"Started new archive: {zip_name}")
        return zip_path

    start_new_zip()
    for filename in file_list:
        check_cancel()
        full_path = os.path.join(base_dir, filename)
        if not os.path.isfile(full_path):
            continue
        size = os.path.getsize(full_path)
        if current_size + size > MAX_ZIP_SIZE_BYTES:
            zf.write(meta_path, "meta.csv")
            zf.close()
            current_index += 1
            start_new_zip()
        zf.write(full_path, os.path.basename(filename))
        current_size += size

    zf.write(meta_path, "meta.csv")
    zf.close()
    return zip_paths


# ------------------------------------------------------------
# Main converter logic
# ------------------------------------------------------------
def main() -> int:
    workdir = get_working_dir()
    log_path = os.path.join(workdir, "converter_log.txt")
    logger = setup_logger(log_path)

    logger.info(f"Working directory: {workdir}")
    print("Converter started.")
    print()

    try:
        # Step 1 – Find Excel
        check_cancel()
        excel_path = next(
            (os.path.join(workdir, f) for f in sorted(os.listdir(workdir))
             if f.lower().endswith((".xlsx", ".xlsm"))),
            None,
        )
        if not excel_path:
            logger.error("No Excel file found.")
            return 1
        basename = os.path.splitext(os.path.basename(excel_path))[0]
        logger.info(f"Found Excel file: {excel_path}")

        # Step 2 – Convert Excel → CSV
        check_cancel()
        meta_path = os.path.join(workdir, "meta.csv")
        convert_excel_to_csv(excel_path, meta_path, logger)

        # Step 3 – Validate meta
        check_cancel()
        print("Validating meta.csv...")
        logger.info("=== START VALIDATION ===")
        if not validate_and_fix_meta(meta_path, logger):
            print("Validation failed - see converter_log.txt for details.")
            return 1
        print()

        # Step 4 – Locate STL files
        check_cancel()
        logger.info("=== START SCANNING FOR STL FILES ===")
        logger.info("Scanning for STL files in root and subfolders...")
        stl_folders, root_stls = [], []
        for entry in sorted(os.listdir(workdir)):
            check_cancel()
            p = os.path.join(workdir, entry)
            if os.path.isdir(p):
                if any(f.lower().endswith(".stl") for f in os.listdir(p)):
                    stl_folders.append(p)
            elif entry.lower().endswith(".stl"):
                root_stls.append(entry)

        total = len(root_stls) + sum(
            1 for f in stl_folders for fn in os.listdir(f) if fn.lower().endswith(".stl")
        )
        if total == 0:
            logger.error("No STL files found.")
            return 1

        logger.info(f"Found {total} STL file(s) in total.")
        print(f"Found {total} STL files.")
        print()

        # Step 5 – Temp folder for meta
        temp_dir = os.path.join(workdir, "_temp_converter")
        os.makedirs(temp_dir, exist_ok=True)
        target_meta = os.path.join(temp_dir, "meta.csv")
        shutil.copy2(meta_path, target_meta)

        # Step 6 – Create ZIPs (folders + root)
        logger.info("=== START PACKAGING ===")
        logger.info("Creating ZIP(s) with 900 MB limit...")
        for stl_folder in stl_folders:
            check_cancel()
            folder_name = os.path.basename(stl_folder)
            logger.info(f"Packaging folder: {stl_folder}")
            print(f"Packaging folder '{folder_name}'...")
            files = [f for f in os.listdir(stl_folder) if f.lower().endswith(".stl")]
            zip_paths = zip_with_limit(files, stl_folder, f"{basename}_{folder_name}",
                                       target_meta, workdir, logger)
            for zp in zip_paths:
                logger.info(f"Created archive: {zp} ({os.path.getsize(zp)/1024/1024:.2f} MB)")

        if root_stls:
            check_cancel()
            logger.info(f"Packaging root STL files ({len(root_stls)} parts)...")
            print(f"Packaging root STL files ({len(root_stls)} parts)...")
            zip_paths = zip_with_limit(root_stls, workdir, f"{basename}_root",
                                       target_meta, workdir, logger)
            for zp in zip_paths:
                logger.info(f"Created archive: {zp} ({os.path.getsize(zp)/1024/1024:.2f} MB)")

        # Step 7 – Cleanup
        check_cancel()
        logger.info("=== START CLEANUP ===")
        shutil.rmtree(temp_dir, ignore_errors=True)
        logger.info("Cleanup complete.")
        print("Cleanup complete.")
        return 0

    except RuntimeError as e:
        logger.warning(str(e))
        print(str(e))
        return 1
    except Exception as e:
        logger.exception("Unexpected error.")
        print(f"Unexpected error: {e}")
        return 1
    finally:
        logger.info("Converter finished.")
        # --- Cleanup leftover temp files after run ---
        try:
            safe_delete(os.path.join(workdir, "meta.csv"), logger)
            safe_delete(os.path.join(workdir, "converter_log.txt"), logger)
            for f in os.listdir(workdir):
                if f.lower().startswith("_cleaned_") and f.lower().endswith((".xlsx", ".xlsm")):
                    safe_delete(os.path.join(workdir, f), logger)
        except Exception as e:
            logger.warning(f"Post-cleanup failed: {e}")

        for h in logger.handlers[:]:
            try:
                h.flush(); h.close()
            except Exception:
                pass
        logging.shutdown()

        def safe_delete(path: str, logger: logging.Logger | None = None):
            """Try deleting a file, log but ignore errors."""
            try:
                if os.path.exists(path):
                    os.remove(path)
                    if logger:
                        logger.info(f"Deleted temporary file: {os.path.basename(path)}")
            except Exception as e:
                if logger:
                    logger.warning(f"Could not delete {path}: {e}")



if __name__ == "__main__":
    sys.exit(main())