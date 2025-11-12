import os
import threading
import logging
import io
import customtkinter as ctk
import converter


APP_TITLE = "AM-Flow Converter"
ENCODING = "utf-8"


class InMemoryLogHandler(logging.Handler):
    """A handler that forwards log lines to the GUI directly."""
    def __init__(self, gui_ref):
        super().__init__()
        self.gui_ref = gui_ref

    def emit(self, record):
        msg = self.format(record)
        self.gui_ref.append_text(msg)


class ConverterGUI(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- Window setup ---
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("dark-blue")

        self.title(APP_TITLE)
        self.geometry("680x420")
        self.minsize(680, 420)

        self.running = False
        self.cancel_requested = False
        self.logger = None

        # --- Layout ---
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # --- Header ---
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
            self.header_frame, text="Idle", font=("Segoe UI", 11)
        )
        self.progress_label.grid(row=2, column=0, sticky="w", pady=(3, 0))

        self.progress_bar = ctk.CTkProgressBar(self.header_frame)
        self.progress_bar.grid(row=3, column=0, sticky="ew", pady=(4, 0))
        self.progress_bar.set(0.0)

        # --- Log area ---
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
    def append_text(self, message: str):
        self.textbox.configure(state="normal")
        self.textbox.insert("end", message.strip() + "\n")
        self.textbox.see("end")
        self.textbox.configure(state="disabled")

    def set_status(self, message: str):
        self.status_label.configure(text=message)

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

    def start_conversion(self):
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

    def run_conversion(self):
        import threading
        import io
        import re
        import time

        # Prepare in-memory log stream
        log_stream = io.StringIO()
        handler = logging.StreamHandler(log_stream)
        handler.setFormatter(logging.Formatter("%(message)s"))

        # Use the same logger as the converter
        converter_logger = logging.getLogger("converter")
        converter_logger.setLevel(logging.INFO)
        converter_logger.handlers.clear()
        converter_logger.addHandler(handler)

        self.append_text("Converter started (RAM mode).")
        self.set_status("Running converter...")
        self.progress_bar.set(0.0)
        self.progress_label.configure(text="Starting...")
        self.update_idletasks()

        # --- Backend Thread ---
        def backend_task():
            try:
                exit_code = converter.main(converter_logger)
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

        backend_thread = threading.Thread(target=backend_task, daemon=True)
        backend_thread.start()

        # --- Live GUI updates from log stream ---
        def tail_log():
            last_pos = 0
            last_stage = ""

            progress_map = {
                "found excel": (0.15, "Found Excel file"),
                "converting excel": (0.25, "Converting Excel to CSV"),
                "validating meta": (0.45, "Validating meta.csv"),
                "scanning for stl": (0.60, "Scanning STL files"),
                "creating zip": (0.75, "Packaging ZIP archives"),
                "cleanup complete": (0.95, "Finalizing"),
                "converter finished": (1.0, "Completed"),
            }

            while backend_thread.is_alive():
                log_text = log_stream.getvalue()
                new_text = log_text[last_pos:]
                last_pos = len(log_text)

                if new_text.strip():
                    # Remove timestamps and clean lines
                    clean_lines = re.sub(r"^\d{4}-\d{2}-\d{2} .*?\] ", "", new_text, flags=re.MULTILINE)
                    self.append_text(clean_lines.strip())

                    # Check for progress hints
                    lower_text = clean_lines.lower()
                    for key, (val, label) in progress_map.items():
                        if key in lower_text and label != last_stage:
                            self.progress_bar.set(val)
                            self.progress_label.configure(text=label)
                            self.update_idletasks()
                            last_stage = label
                            break

                time.sleep(0.2)

            # Final read when thread completes
            remaining = log_stream.getvalue()[last_pos:]
            if remaining.strip():
                clean_lines = re.sub(r"^\d{4}-\d{2}-\d{2} .*?\] ", "", remaining, flags=re.MULTILINE)
                self.append_text(clean_lines.strip())

            self.progress_bar.set(1.0)
            self.set_status("Done.")
            self.progress_label.configure(text="Completed")

        threading.Thread(target=tail_log, daemon=True).start()

def main():
    app = ConverterGUI()
    app.mainloop()


if __name__ == "__main__":
    main()
