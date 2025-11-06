import os
import sys
import csv
import zipfile
import shutil
import logging
from typing import Optional, List, Dict, Tuple
import converter

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
ENCODING = "utf-8"
MAX_ZIP_SIZE_MB = 900
MAX_ZIP_SIZE_BYTES = MAX_ZIP_SIZE_MB * 1024 * 1024


class ConverterGUI(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- General window setup ---
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("dark-blue")

        self.title(APP_TITLE)
        self.geometry("680x420")
        self.minsize(680, 420)

        self.running = False
        self.cancel_requested = False
        self.logger: Optional[logging.Logger] = None
        self.log_path: Optional[str] = None

        # --- Layout: 3 rows: header, log area, controls ---
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # --- Header / status ---
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

        # --- Log / message area ---
        self.textbox = ctk.CTkTextbox(self, wrap="word")
        self.textbox.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
        self.textbox.configure(state="disabled")

        # --- Controls ---
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

        # --- Start automatically ---
        self.after(200, self.start_conversion)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    # --- GUI helpers ---

    def append_text(self, message: str) -> None:
        self.textbox.configure(state="normal")
        self.textbox.insert("end", message + "\n")
        self.textbox.see("end")
        self.textbox.configure(state="disabled")

    def set_status(self, message: str) -> None:
        self.status_label.configure(text=message)

    # --- Button events ---

    def on_cancel(self):
        if self.running:
            self.cancel_requested = True
            converter.request_cancel()
            self.set_status("Cancelling...")
            self.append_text("Cancellation requested by user.")
            self.cancel_button.configure(state="disabled")

    def on_close(self):
        if self.running:
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

        self.after(50, self.run_conversion)

    def run_conversion(self) -> None:
        import threading, time, re

        workdir = converter.get_working_dir()
        log_path = os.path.join(workdir, "converter_log.txt")
        self.log_path = log_path
        self.logger = converter.setup_logger(log_path)

        self.append_text("Converter started.")
        self.append_text(f"Working directory: {workdir}")
        self.set_status("Running converter...")
        self.progress_bar.set(0.0)
        self.progress_label.configure(text="Starting...")
        self.update_idletasks()

        # --- Run converter backend in a background thread ---
        def run_backend():
            try:
                exit_code = converter.main()
                if exit_code == 0:
                    self.set_status("All tasks completed successfully.")
                    self.progress_bar.set(1.0)
                    self.progress_label.configure(text="Completed.")
                else:
                    self.set_status("Conversion failed — check log.")
                    self.progress_label.configure(text="Error.")
            except RuntimeError as e:
                self.append_text(str(e))
                self.set_status("Cancelled by user.")
                self.progress_label.configure(text="Cancelled.")
            except Exception as e:
                self.append_text(f"Error: {e}")
                self.set_status("Error occurred.")
                self.progress_label.configure(text="Error.")
            finally:
                self.running = False
                self.cancel_requested = False
                self.cancel_button.configure(state="disabled")
                self.close_button.configure(state="normal")

        backend_thread = threading.Thread(target=run_backend, daemon=True)
        backend_thread.start()

        # --- Tail the log and update GUI dynamically ---
        def tail_log():
            last_pos = 0
            progress = 0.0
            last_stage = ""
            try:
                while backend_thread.is_alive() or os.path.exists(log_path):
                    if os.path.exists(log_path):
                        with open(log_path, "r", encoding="utf-8", errors="ignore") as f:
                            f.seek(last_pos)
                            new_lines = f.read()
                            last_pos = f.tell()
                            if new_lines.strip():
                                clean_lines = re.sub(
                                    r"^\d{4}-\d{2}-\d{2} .*?\] ", "", new_lines, flags=re.MULTILINE
                                )
                                self.append_text(clean_lines.strip())

                                # --- Heuristic progress tracking with immediate stage locking ---
                                text = new_lines.lower()
                                updated = False

                                # Detect new stage markers before any other messages
                                if "=== start validation" in text:
                                    progress, label = 0.45, "Validating meta.csv"
                                    updated = True
                                elif "=== start scanning for stl files" in text:
                                    progress, label = 0.60, "Scanning STL files"
                                    updated = True
                                elif "=== start packaging" in text:
                                    progress, label = 0.75, "Packaging ZIP archives"
                                    updated = True
                                elif "=== start cleanup" in text:
                                    progress, label = 0.95, "Finalizing"
                                    updated = True
                                elif "converter finished" in text or "all zip creation completed" in text:
                                    progress, label = 1.0, "Completed"
                                    updated = True

                                # Fallback: legacy text search (but only if no stage marker triggered)
                                if not updated:
                                    if "found excel" in text:
                                        progress, label = 0.15, "Found Excel file"
                                    elif "converting excel" in text:
                                        progress, label = 0.25, "Converting Excel to CSV"
                                    elif "validating meta" in text:
                                        progress, label = 0.45, "Validating meta.csv"
                                    elif (
                                        "creating zip" in text
                                        or "creating archive" in text
                                        or "started new archive" in text
                                        or "created archive" in text
                                    ):
                                        progress, label = 0.75, "Packaging ZIP archives"
                                    elif "cleanup complete" in text:
                                        progress, label = 0.95, "Finalizing"
                                    else:
                                        label = last_stage

                                # Update GUI only when the label actually changes
                                if label != last_stage:
                                    self.progress_label.configure(text=label)
                                    last_stage = label
                                    self.progress_bar.set(progress)
                                    self.update_idletasks()

                    time.sleep(0.3)
            except Exception as e:
                self.append_text(f"Log reader error: {e}")

        threading.Thread(target=tail_log, daemon=True).start()


def main():
    app = ConverterGUI()
    app.mainloop()


if __name__ == "__main__":
    main()
