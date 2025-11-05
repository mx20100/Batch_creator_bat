import os
import sys
import csv
import zipfile
import shutil
import logging
from typing import Optional, List, Dict, Tuple

import customtkinter as ctk
import pandas as pd
from openpyxl import load_workbook

APP_TITLE = "AM-Flow Converter"

REQUIRED_COLUMNS = [
    "batch",
    "filename",
    "material",
    "part_id",
    "copies",
    "next_step",
    "order_id",
    "technology",
]

MAX_STL_PER_ZIP = 100


# ---------------------------
# Core (backend) functions
# ---------------------------

def get_working_dir() -> str:
    """Use the current working directory (where the user runs the EXE / script)."""
    return os.getcwd()


def setup_logger(log_path: str) -> logging.Logger:
    logger = logging.getLogger("converter")
    logger.setLevel(logging.INFO)
    # Avoid duplicate handlers if run multiple times
    logger.handlers.clear()

    fh = logging.FileHandler(log_path, encoding="utf-8")
    fh.setLevel(logging.INFO)
    formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
    fh.setFormatter(formatter)
    logger.addHandler(fh)

    logger.info("=" * 60)
    logger.info("Converter started")
    return logger


def find_excel_file(workdir: str) -> Optional[str]:
    """Find first .xlsx, then .xlsm in the working directory."""
    for ext in (".xlsx", ".xlsm"):
        candidates = [f for f in os.listdir(workdir) if f.lower().endswith(ext)]
        if candidates:
            candidates.sort()
            return os.path.join(workdir, candidates[0])
    return None


def convert_excel_to_csv(
    excel_path: str,
    csv_path: str,
    logger: logging.Logger,
) -> None:
    """
    Convert the first non-empty visible worksheet of the Excel file to CSV.
    Removes completely empty sheets first. Falls back to pandas if needed.
    """
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

    tmp_excel = None
    if removed_sheets:
        logger.info(f"Removed {removed_sheets} empty sheet(s) before conversion.")
        # Save cleaned workbook to a temporary file so openpyxl can read it again safely
        tmp_excel = os.path.join(
            os.path.dirname(excel_path),
            "_cleaned_" + os.path.basename(excel_path),
        )
        wb.save(tmp_excel)
        wb.close()
        # Reopen cleaned workbook as read-only
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
            df.to_csv(csv_path, index=False, encoding="utf-8")
            logger.info("Conversion to meta.csv completed using pandas fallback.")
        except Exception as e:
            logger.error(f"All sheets empty or unreadable: {e}")
            raise ValueError(
                f"No usable worksheets found in '{os.path.basename(excel_path)}'."
            )
    else:
        logger.info(f"Using worksheet: {ws.title}")
        with open(csv_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            for row in ws.iter_rows(values_only=True):
                writer.writerow(["" if v is None else str(v) for v in row])
        logger.info("Conversion to meta.csv completed.")

    # Clean up temporary cleaned workbook if we created one
    if tmp_excel and os.path.exists(tmp_excel):
        try:
            os.remove(tmp_excel)
        except OSError:
            pass


import re

def sanitize_filename(value: str) -> str:
    """Replace unsupported chars with underscores and collapse duplicates."""
    sanitized = re.sub(r"[^A-Za-z0-9_.]", "_", value)
    sanitized = re.sub(r"_+", "_", sanitized)
    sanitized = sanitized.strip("_ ")
    return sanitized


def validate_and_fix_meta(csv_path: str, stl_folder: str, logger: logging.Logger) -> bool:
    """
    Validates and sanitizes meta.csv:
    - Header must match REQUIRED_COLUMNS (in order, case-insensitive)
    - Rows with any data must have all required fields filled
    - copies == '' or '0' -> set to '1'
    - filename without '.stl' (case-insensitive) -> append '.stl'
    - Sanitizes filenames (allowed: A-Z, a-z, 0-9, _, .)
    - Renames corresponding STL files in stl_folder
    - Removes rows for missing STL files
    """
    logger.info("Validating and sanitizing meta.csv...")

    errors: list[str] = []
    fixed_copies = 0
    fixed_filenames = 0
    sanitized_files = 0
    renamed_files = 0
    removed_rows = 0

    # Read meta.csv
    with open(csv_path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        rows = list(reader)
        header = reader.fieldnames

    if header is None:
        logger.error("Validation failed: no header row in meta.csv")
        return False

    normalized_header = [h.strip().lower() for h in header[: len(REQUIRED_COLUMNS)]]
    if normalized_header != REQUIRED_COLUMNS:
        logger.error(f"Header mismatch. Found: {header}, Expected: {REQUIRED_COLUMNS}")
        return False

    # Map STL files on disk (case-insensitive)
    stl_files_on_disk = {f.lower(): f for f in os.listdir(stl_folder) if f.lower().endswith(".stl")}
    cleaned_rows = []

    for i, row in enumerate(rows, start=2):
        if not any((row.get(k) or "").strip() for k in row.keys()):
            continue

        # Fix copies
        val = (row.get("copies") or "").strip()
        if val == "" or val == "0":
            row["copies"] = "1"
            fixed_copies += 1

        # Fix and sanitize filename
        fname = (row.get("filename") or "").strip()
        if not fname:
            errors.append(f"Row {i}: missing filename")
            logger.error(f"Row {i}: missing filename")
            continue

        if not fname.lower().endswith(".stl"):
            fname += ".stl"
            fixed_filenames += 1

        clean_name = sanitize_filename(fname)
        if clean_name != fname:
            sanitized_files += 1
            logger.info(f"Sanitized filename: '{fname}' -> '{clean_name}'")

        # Rename the STL file on disk if necessary
        original_disk_name = stl_files_on_disk.get(fname.lower()) or stl_files_on_disk.get(clean_name.lower())
        if original_disk_name:
            src = os.path.join(stl_folder, original_disk_name)
            dst = os.path.join(stl_folder, clean_name)
            if src != dst:
                try:
                    os.rename(src, dst)
                    renamed_files += 1
                    logger.info(f"Renamed STL file: '{original_disk_name}' -> '{clean_name}'")
                    # Update disk map
                    stl_files_on_disk.pop(original_disk_name.lower(), None)
                    stl_files_on_disk[clean_name.lower()] = clean_name
                except Exception as e:
                    logger.error(f"Could not rename '{original_disk_name}' -> '{clean_name}': {e}")
        else:
            logger.warning(f"STL file missing for row {i}: {fname} — row removed.")
            removed_rows += 1
            continue

        row["filename"] = clean_name
        cleaned_rows.append(row)

    # Write cleaned CSV
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=REQUIRED_COLUMNS)
        writer.writeheader()
        writer.writerows(cleaned_rows)

    # Summary
    if fixed_copies:
        logger.info(f"Corrected {fixed_copies} row(s) with copies=0 or empty.")
    if fixed_filenames:
        logger.info(f"Added missing .stl extension to {fixed_filenames} filename(s).")
    if sanitized_files:
        logger.info(f"Sanitized {sanitized_files} filename(s) with invalid characters.")
    if renamed_files:
        logger.info(f"Renamed {renamed_files} STL file(s) on disk to match sanitized names.")
    if removed_rows:
        logger.warning(f"Removed {removed_rows} row(s) due to missing STL files.")

    if errors:
        logger.error("Validation failed with errors.")
        return False

    logger.info("meta.csv passed validation.")
    return True

# ---------------------------
# GUI application (customtkinter)
# ---------------------------

class ConverterGUI(ctk.CTk):
    def __init__(self):
        super().__init__()

        # General window setup
        ctk.set_appearance_mode("system")  # system-adaptive
        ctk.set_default_color_theme("dark-blue")  # just a color theme, still system-aware

        self.title(APP_TITLE)
        self.geometry("680x420")
        self.minsize(680, 420)

        self.running = False
        self.cancel_requested = False
        self.logger: Optional[logging.Logger] = None
        self.log_path: Optional[str] = None

        # Layout: 3 rows: header, text area, controls
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # Header / status
        self.header_frame = ctk.CTkFrame(self)
        self.header_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 5))
        self.header_frame.grid_columnconfigure(0, weight=1)

        self.title_label = ctk.CTkLabel(
            self.header_frame, text=APP_TITLE, font=("Segoe UI", 20, "bold")
        )
        self.title_label.grid(row=0, column=0, sticky="w")

        self.status_label = ctk.CTkLabel(
            self.header_frame, text="Starting...", font=("Segoe UI", 12)
        )
        self.status_label.grid(row=1, column=0, sticky="w", pady=(4, 0))

        self.progress_label = ctk.CTkLabel(
            self.header_frame, text="Step 0 of 0", font=("Segoe UI", 11)
        )
        self.progress_label.grid(row=2, column=0, sticky="w", pady=(3, 0))

        self.progress_bar = ctk.CTkProgressBar(self.header_frame)
        self.progress_bar.grid(row=3, column=0, sticky="ew", pady=(4, 0))
        self.progress_bar.set(0.0)

        # Log / messages area
        self.textbox = ctk.CTkTextbox(self, wrap="word")
        self.textbox.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
        self.textbox.configure(state="disabled")

        # Control buttons
        self.controls_frame = ctk.CTkFrame(self)
        self.controls_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=(5, 10))
        self.controls_frame.grid_columnconfigure(0, weight=1)
        self.controls_frame.grid_columnconfigure(1, weight=0)

        self.cancel_button = ctk.CTkButton(
            self.controls_frame, text="Cancel", command=self.on_cancel
        )
        self.cancel_button.grid(row=0, column=0, sticky="w")

        self.close_button = ctk.CTkButton(
            self.controls_frame, text="Close", command=self.on_close, state="disabled"
        )
        self.close_button.grid(row=0, column=1, sticky="e")

        # Start conversion automatically after the window is displayed
        self.after(200, self.start_conversion)

        # Handle window close button
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    # --- GUI helpers ---

    def append_text(self, message: str) -> None:
        self.textbox.configure(state="normal")
        self.textbox.insert("end", message + "\n")
        self.textbox.see("end")
        self.textbox.configure(state="disabled")

    def set_status(self, message: str) -> None:
        self.status_label.configure(text=message)

    def set_step(self, current: int, total: int, description: str) -> None:
        if total <= 0:
            self.progress_bar.set(0.0)
            self.progress_label.configure(text=description)
            return

        fraction = current / total
        if fraction < 0.0:
            fraction = 0.0
        if fraction > 1.0:
            fraction = 1.0

        self.progress_bar.set(fraction)
        self.progress_label.configure(
            text=f"Step {current} of {total}: {description}"
        )

    def on_cancel(self) -> None:
        if self.running:
            self.cancel_requested = True
            self.set_status("Cancelling...")
            self.append_text("Cancellation requested by user.")

    def on_close(self) -> None:
        if self.running:
            # Allow immediate close anyway
            self.destroy()
        else:
            self.destroy()

    # --- Conversion orchestration ---

    def start_conversion(self) -> None:
        if self.running:
            return
        self.running = True
        self.cancel_requested = False
        self.close_button.configure(state="disabled")
        self.cancel_button.configure(state="normal")
        self.textbox.configure(state="normal")
        self.textbox.delete("1.0", "end")
        self.textbox.configure(state="disabled")

        # Run conversion in the same thread (work is not huge) but keep UI responsive via .update()
        self.after(50, self.run_conversion)

    def check_cancel(self):
        if self.cancel_requested:
            raise RuntimeError("Conversion cancelled by user.")

    def run_conversion(self) -> None:
        workdir = get_working_dir()
        log_path = os.path.join(workdir, "converter_log.txt")
        self.log_path = log_path
        logger = setup_logger(log_path)
        self.logger = logger

        # Create a temporary working folder for intermediate files
        temp_dir = os.path.join(workdir, "_temp_converter")
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir, ignore_errors=True)
        os.makedirs(temp_dir, exist_ok=True)

        self.append_text("Converter started.")
        self.append_text(f"Working directory: {workdir}")
        logger.info(f"Working directory: {workdir}")

        success = False

        # We'll keep the progress bar coarse: 7 logical steps
        total_steps = 7
        step = 0

        try:
            # Step 1: Find Excel file
            step = 1
            self.set_step(step, total_steps, "Searching for Excel file")
            self.set_status("Searching for Excel file...")
            self.update_idletasks()
            self.check_cancel()

            excel_path = find_excel_file(workdir)
            if not excel_path:
                msg = "No Excel file (.xlsx or .xlsm) found in this folder."
                self.append_text(msg)
                logger.error(msg)
                return

            basename = os.path.splitext(os.path.basename(excel_path))[0]
            self.append_text(f"Found Excel file: {os.path.basename(excel_path)}")
            logger.info(f"Found Excel file: {excel_path}")

            # Step 2: Convert Excel to meta.csv (in temp folder)
            step = 2
            self.set_step(step, total_steps, "Converting Excel to meta.csv")
            self.set_status("Converting Excel to meta.csv...")
            self.update_idletasks()
            self.check_cancel()

            meta_path = os.path.join(temp_dir, "meta.csv")
            self.append_text("Converting Excel to meta.csv...")
            convert_excel_to_csv(excel_path, meta_path, logger)
            if not os.path.exists(meta_path):
                msg = "Conversion failed: meta.csv was not created."
                self.append_text(msg)
                logger.error(msg)
                return
            self.append_text("Conversion successful.")
            self.append_text("")
            self.update_idletasks()

            # Step 3: Validate meta.csv
            step = 3
            self.set_step(step, total_steps, "Validating meta.csv")
            self.set_status("Validating meta.csv...")
            self.update_idletasks()
            self.check_cancel()

            self.append_text("Validating meta.csv...")
            stl_folder = None
            for name in os.listdir(workdir):
                path = os.path.join(workdir, name)
                if os.path.isdir(path):
                    if any(f.lower().endswith(".stl") for f in os.listdir(path)):
                        stl_folder = path
                        break

            if not stl_folder:
                msg = "No folder with STL files found."
                self.append_text(msg)
                logger.error(msg)
                return
            if not validate_and_fix_meta(meta_path, stl_folder, logger):
                self.append_text("Validation failed - see converter_log.txt for details.")
                return
            
            self.append_text("Validation completed.")
            self.append_text("")
            self.update_idletasks()

            # Load rows from validated meta.csv
            with open(meta_path, newline="", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                rows = list(reader)

            if not rows:
                msg = "meta.csv contains no data rows."
                self.append_text(msg)
                logger.error(msg)
                return

            # Step 4: Group by (batch, material) and check STL files
            step = 4
            self.set_step(step, total_steps, "Grouping rows and checking STL files")
            self.set_status("Grouping rows by batch/material and checking STL files...")
            self.update_idletasks()
            self.check_cancel()

            from collections import defaultdict

            group_map: Dict[Tuple[str, str], List[Dict[str, str]]] = defaultdict(list)
            for row in rows:
                batch = (row.get("batch") or "").strip()
                material = (row.get("material") or "").strip()
                if not batch or not material:
                    # Should not happen after validation, but be safe
                    logger.warning(f"Skipping row with missing batch/material: {row}")
                    continue
                group_map[(batch, material)].append(row)

            if not group_map:
                msg = "No valid rows with batch and material found in meta.csv."
                self.append_text(msg)
                logger.error(msg)
                return

            # Map subfolders by name (case-insensitive) -> path
            material_dirs: Dict[str, str] = {}
            for name in os.listdir(workdir):
                path = os.path.join(workdir, name)
                if os.path.isdir(path):
                    material_dirs[name.lower()] = path

            groups: List[Dict[str, object]] = []

            for (batch, material), group_rows in sorted(group_map.items(), key=lambda x: (x[0][0], x[0][1])):
                self.check_cancel()
                self.append_text(f"Processing batch '{batch}' / material '{material}'...")
                logger.info(f"Processing group: batch={batch}, material={material}")

                folder = material_dirs.get(material.lower())
                if not folder:
                    msg = f"Folder '{material}' not found for batch {batch}."
                    self.append_text(msg)
                    logger.error(msg)
                    return

                existing_rows: List[Dict[str, str]] = []
                missing_files: List[str] = []

                for row in group_rows:
                    fname = (row.get("filename") or "").strip()
                    if not fname:
                        logger.warning(f"Skipping row with empty filename in batch {batch}, material {material}")
                        continue
                    stl_path = os.path.join(folder, fname)
                    if not os.path.exists(stl_path):
                        missing_files.append(fname)
                    else:
                        existing_rows.append(row)

                if missing_files:
                    unique_missing = sorted(set(missing_files))
                    warn_msg = (
                        f"Batch {batch}, material {material}: "
                        f"missing STL files removed from meta.csv: {', '.join(unique_missing)}"
                    )
                    self.append_text(warn_msg)
                    logger.warning(warn_msg)

                if not existing_rows:
                    msg = f"No existing STL files found for batch {batch}, material {material}."
                    self.append_text(msg)
                    logger.error(msg)
                    return

                groups.append(
                    {
                        "batch": batch,
                        "material": material,
                        "folder": folder,
                        "rows": existing_rows,
                    }
                )

            if not groups:
                msg = "No usable (batch, material) groups after checking STL files."
                self.append_text(msg)
                logger.error(msg)
                return

            self.append_text("")
            self.update_idletasks()

            # Step 5: Create ZIP archives with splitting (max 100 STLs per zip)
            step = 5
            self.set_step(step, total_steps, "Creating ZIP archives")
            self.set_status("Packaging files into ZIP...")
            self.update_idletasks()
            self.check_cancel()

            for grp in groups:
                self.check_cancel()

                batch = grp["batch"]
                material = grp["material"]
                folder = grp["folder"]
                group_rows = grp["rows"]
                num_parts = len(group_rows)

                safe_material = str(material).replace(" ", "_")
                self.append_text(
                    f"Creating ZIP(s) for batch '{batch}', material '{material}' ({num_parts} parts)..."
                )
                logger.info(
                    f"Creating ZIP(s) for batch={batch}, material={material}, parts={num_parts}"
                )

                # Split into chunks of at most MAX_STL_PER_ZIP
                chunks: List[List[Dict[str, str]]] = [
                    group_rows[i : i + MAX_STL_PER_ZIP]
                    for i in range(0, num_parts, MAX_STL_PER_ZIP)
                ]

                for idx, chunk in enumerate(chunks, start=1):
                    self.check_cancel()

                    if len(chunks) == 1:
                        zip_name = f"{batch}_{safe_material}.zip"
                        chunk_dir_name = f"{batch}_{safe_material}"
                    else:
                        zip_name = f"{batch}_{safe_material}_part{idx}.zip"
                        chunk_dir_name = f"{batch}_{safe_material}_part{idx}"

                    zip_path = os.path.join(workdir, zip_name)
                    chunk_dir = os.path.join(temp_dir, chunk_dir_name)
                    os.makedirs(chunk_dir, exist_ok=True)

                    chunk_meta_path = os.path.join(chunk_dir, "meta.csv")

                    # Write chunk meta.csv
                    with open(chunk_meta_path, "w", newline="", encoding="utf-8") as f:
                        writer = csv.DictWriter(f, fieldnames=REQUIRED_COLUMNS)
                        writer.writeheader()
                        writer.writerows(chunk)

                    # Create ZIP: meta.csv + STL files at root
                    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                        zf.write(chunk_meta_path, "meta.csv")
                        for row in chunk:
                            fname = (row.get("filename") or "").strip()
                            if not fname:
                                continue
                            stl_path = os.path.join(folder, fname)
                            if os.path.exists(stl_path):
                                zf.write(stl_path, os.path.basename(fname))

                    self.append_text(f"  Created {zip_name}")
                    logger.info(f"Created archive: {zip_path}")

                self.append_text("")

            self.update_idletasks()

            # Step 6: Cleanup temporary files
            step = 6
            self.set_step(step, total_steps, "Cleaning up temporary files")
            self.set_status("Cleaning up temporary files...")
            self.update_idletasks()
            self.check_cancel()

            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)
            self.append_text("Cleanup complete.")
            logger.info("Cleanup complete.")

            # Step 7: Done
            step = 7
            self.set_step(step, total_steps, "Done")
            self.set_status("All tasks completed successfully.")
            self.append_text("")
            self.append_text("All tasks completed successfully.")
            success = True

        except RuntimeError as e:
            # Used for user cancellation
            logger.info(f"Cancelled: {e}")
            self.append_text("Conversion cancelled.")
            self.set_status("Cancelled by user.")
        except Exception as e:
            logger.exception("Unexpected error occurred.")
            self.append_text("An unexpected error occurred:")
            self.append_text(repr(e))
            self.set_status("Error occurred.")
        finally:
            logger.info("Converter finished.")
            self.running = False
            self.cancel_requested = False
            self.cancel_button.configure(state="disabled")
            self.close_button.configure(state="normal")
            # We keep converter_log.txt in the root folder (no automatic deletion).


def main():
    app = ConverterGUI()
    app.mainloop()


if __name__ == "__main__":
    main()
