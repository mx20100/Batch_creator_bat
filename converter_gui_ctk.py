import os
import sys
import csv
import zipfile
import shutil
import logging
from datetime import datetime
from typing import Optional, List

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
            df.to_csv(csv_path, index=False, encoding="utf-8-sig")
            logger.info("Conversion to meta.csv completed using pandas fallback.")
        except Exception as e:
            logger.error(f"All sheets empty or unreadable: {e}")
            raise ValueError(
                f"No usable worksheets found in '{os.path.basename(excel_path)}'."
            )
    else:
        logger.info(f"Using worksheet: {ws.title}")
        with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
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

    errors: List[str] = []
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

    normalized_header = [h.strip().lower() for h in header[: len(REQUIRED_COLUMNS)]]
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
            missing = [
                k for k in REQUIRED_COLUMNS if not (row.get(k) or "").strip()
            ]
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


def find_stl_folder(workdir: str, logger: logging.Logger) -> Optional[str]:
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
                rel_path = os.path.relpath(full_path, base_folder)
                zf.write(full_path, rel_path)

    logger.info("Zip creation completed.")


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
        self.success = False
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
        """Handle Close button or window X."""
        try:
            # Only delete log if run finished successfully
            if not self.running and self.success:
                logging.shutdown()
                if self.log_path and os.path.exists(self.log_path):
                    try:
                        os.remove(self.log_path)
                        self.append_text("(Temporary log file cleaned up on exit.)")
                    except PermissionError:
                        self.append_text("(Log file still in use; will remain for review.)")
                    except Exception as e:
                        self.append_text(f"(Could not remove log file: {e})")
        except Exception:
            pass
        finally:
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

        self.append_text("Converter started.")
        self.append_text(f"Working directory: {workdir}")
        logger.info(f"Working directory: {workdir}")

        success = False
        meta_path = os.path.join(workdir, "meta.csv")

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

            # Step 2: Convert Excel to meta.csv
            step = 2
            self.set_step(step, total_steps, "Converting Excel to meta.csv")
            self.set_status("Converting Excel to meta.csv...")
            self.update_idletasks()
            self.check_cancel()

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
            if not validate_and_fix_meta(meta_path, logger):
                self.append_text("Validation failed — see converter_log.txt for details.")
                return
            self.append_text("Validation completed.")
            self.append_text("")
            self.update_idletasks()

            # Step 4: Find STL folder
            step = 4
            self.set_step(step, total_steps, "Searching for STL folder")
            self.set_status("Searching for STL folder...")
            self.update_idletasks()
            self.check_cancel()

            stl_folder = find_stl_folder(workdir, logger)
            if not stl_folder:
                msg = "No folder with STL files found."
                self.append_text(msg)
                if os.path.exists(meta_path):
                    os.remove(meta_path)
                return
            self.append_text(f"STL folder found: {os.path.basename(stl_folder)}")
            self.append_text("")
            self.update_idletasks()

            # Step 5: Copy meta.csv into STL folder
            step = 5
            self.set_step(step, total_steps, "Copying meta.csv into STL folder")
            self.set_status("Copying meta.csv into STL folder...")
            self.update_idletasks()
            self.check_cancel()

            target_meta = os.path.join(stl_folder, "meta.csv")
            shutil.copy2(meta_path, target_meta)
            logger.info(f"Copied meta.csv to {target_meta}")
            self.append_text("Copied meta.csv into STL folder.")
            self.update_idletasks()

            # Step 6: Create zip archive
            step = 6
            self.set_step(step, total_steps, "Creating ZIP archive")
            self.set_status("Creating ZIP archive...")
            self.update_idletasks()
            self.check_cancel()

            zip_name = f"{basename}.zip"
            zip_path = os.path.join(workdir, zip_name)
            self.append_text(f"Creating archive: {zip_name}...")
            zip_stl_folder(stl_folder, zip_path, logger)

            if not os.path.exists(zip_path):
                msg = "Failed to create ZIP archive."
                self.append_text(msg)
                logger.error(msg)
                return

            self.append_text(f"Created archive: {zip_name}")
            self.append_text("")
            self.update_idletasks()

            # Step 7: Cleanup
            step = 7
            self.set_step(step, total_steps, "Cleaning up temporary files")
            self.set_status("Cleaning up temporary files...")
            self.update_idletasks()
            self.check_cancel()

            if os.path.exists(target_meta):
                os.remove(target_meta)
            if os.path.exists(meta_path):
                os.remove(meta_path)
            self.append_text("Cleanup complete.")
            logger.info("Cleanup complete.")

            self.success = True
            success = True
            self.set_status("All tasks completed successfully.")
            self.set_step(total_steps, total_steps, "Done")
            self.append_text("")
            self.append_text("All tasks completed successfully.")

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
            # Do NOT delete log file; you asked to keep it.


def main():
    app = ConverterGUI()
    app.mainloop()


if __name__ == "__main__":
    main()