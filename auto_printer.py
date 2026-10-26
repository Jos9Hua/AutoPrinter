import os
from pathlib import Path
import shutil
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import win32print
except ImportError:
    win32print = None

try:
    import winreg
except ImportError:
    winreg = None


class PDFPrinterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Batch Printer")
        self.root.geometry("700x500")
        self.root.minsize(600, 400)

        self.folder_path = tk.StringVar()
        self.printer_name = tk.StringVar()
        self.pdf_files = []
        self.file_settings = {}
        self.bulk_duplex_target = "Double-sided"
        self.log_box = None
        self._pending_log_messages = []

        title_label = tk.Label(
            root, text="Batch PDF Printer", font=("Arial", 16, "bold")
        )
        title_label.pack(pady=(10, 5))

        settings_frame = tk.Frame(root)
        settings_frame.pack(pady=10, padx=20, fill="x")

        printer_label = tk.Label(
            settings_frame, text="Printer:", font=("Arial", 10, "bold")
        )
        printer_label.pack(anchor="w")

        self.printer_combo = ttk.Combobox(
            settings_frame,
            textvariable=self.printer_name,
            values=self.get_printers(),
            state="readonly",
            width=40,
        )
        self.printer_combo.pack(fill="x", pady=(2, 8))

        folder_frame = tk.Frame(root)
        folder_frame.pack(padx=20, fill="x")

        self.entry_path = tk.Entry(
            folder_frame,
            textvariable=self.folder_path,
            font=("Arial", 10),
            state="readonly",
        )
        self.entry_path.pack(side="left", expand=True, fill="x", padx=(0, 10))

        btn_browse = tk.Button(
            folder_frame,
            text="Browse...",
            command=self.browse_folder,
            font=("Arial", 10),
        )
        btn_browse.pack(side="right")

        list_header = tk.Frame(root)
        list_header.pack(fill="x", padx=20)
        list_title = tk.Label(
            list_header, text="PDF Files to Print", font=("Arial", 10, "bold")
        )
        list_title.pack(side="left")
        self.btn_toggle_duplex = tk.Button(
            list_header,
            text="Set all double-sided",
            command=self.toggle_all_duplex,
            state="disabled",
        )
        self.btn_toggle_duplex.pack(side="right")

        self.list_frame = tk.Frame(root)
        self.list_frame.pack(padx=20, pady=(5, 10), fill="both", expand=True)

        self.list_canvas = tk.Canvas(self.list_frame, height=220, bg="#f5f5f5")
        self.list_canvas.pack(side="left", fill="both", expand=True)

        self.list_scrollbar = ttk.Scrollbar(
            self.list_frame, orient="vertical", command=self.list_canvas.yview
        )
        self.list_scrollbar.pack(side="right", fill="y")
        self.list_canvas.configure(yscrollcommand=self.list_scrollbar.set)

        self.list_container = tk.Frame(self.list_canvas, bg="#f5f5f5")
        self.list_window = self.list_canvas.create_window(
            (0, 0), window=self.list_container, anchor="nw"
        )
        self.list_container.bind("<Configure>", self._update_canvas_scrollregion)
        self.list_canvas.bind("<Configure>", self._resize_list_container)

        self.btn_print = tk.Button(
            root,
            text="Print Selected PDFs",
            command=self.print_pdfs,
            font=("Arial", 12, "bold"),
            bg="#4CAF50",
            fg="white",
            padx=10,
            pady=5,
        )
        self.btn_print.pack(pady=(0, 10))

        log_title = tk.Label(root, text="System Log", font=("Arial", 10, "bold"))
        log_title.pack(anchor="w", padx=20)

        self.log_box = tk.Text(
            root, height=8, font=("Courier", 9), state="disabled", bg="#f7f7f7"
        )
        self.log_box.pack(padx=20, pady=(5, 10), fill="both", expand=True)

        for message in self._pending_log_messages:
            self.log(message)
        self._pending_log_messages.clear()
        self.log("Application started. Select a folder to begin.")

    def _update_canvas_scrollregion(self, event=None):
        self.list_canvas.configure(scrollregion=self.list_canvas.bbox("all"))

    def _resize_list_container(self, event):
        self.list_canvas.itemconfigure(self.list_window, width=event.width)

    @staticmethod
    def _find_sumatra_pdf():
        executable = shutil.which("SumatraPDF.exe")
        if executable:
            return executable

        program_files = [
            os.environ.get("ProgramFiles"),
            os.environ.get("ProgramFiles(x86)"),
            os.environ.get("LOCALAPPDATA"),
        ]
        for directory in program_files:
            if directory:
                candidate = Path(directory) / "SumatraPDF" / "SumatraPDF.exe"
                if candidate.is_file():
                    return str(candidate)

        if winreg is not None:
            app_path_key = (
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\SumatraPDF.exe"
            )
            registry_hives = [winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE]
            for hive in registry_hives:
                try:
                    with winreg.OpenKey(hive, app_path_key) as key:
                        executable, _ = winreg.QueryValueEx(key, None)
                    if Path(executable).is_file():
                        return executable
                except OSError:
                    continue

        return ""

    def toggle_all_duplex(self):
        """Apply one duplex setting to all loaded PDF files."""
        if not self.file_settings:
            return

        target = self.bulk_duplex_target
        for settings in self.file_settings.values():
            settings["duplex"].set(target)

        self.bulk_duplex_target = (
            "Single-sided" if target == "Double-sided" else "Double-sided"
        )
        self.btn_toggle_duplex.config(
            text=f"Set all {self.bulk_duplex_target.lower()}"
        )
        self.log(f"Set all {len(self.file_settings)} files to {target}.")

    def browse_folder(self):
        """Open a folder selection dialog and list the PDFs inside."""
        folder_selected = filedialog.askdirectory()
        if not folder_selected:
            return

        self.folder_path.set(folder_selected)
        self.log(f"Selected folder: {folder_selected}")
        self.load_pdf_list(folder_selected)

    def load_pdf_list(self, folder_path):
        """Populate the list with each PDF and per-file print settings."""
        for child in self.list_container.winfo_children():
            child.destroy()

        self.file_settings = {}
        self.bulk_duplex_target = "Double-sided"
        self.btn_toggle_duplex.config(
            text="Set all double-sided", state="disabled"
        )
        folder = Path(folder_path)
        self.pdf_files = sorted(
            [f for f in folder.iterdir() if f.is_file() and f.suffix.lower() == ".pdf"],
            key=lambda item: item.name.lower(),
        )

        if not self.pdf_files:
            empty_label = tk.Label(
                self.list_container,
                text="No PDF files were found in this folder.",
                fg="#666666",
                bg="#f5f5f5",
                anchor="w",
                justify="left",
                padx=10,
                pady=15,
            )
            empty_label.pack(fill="x")
            return

        for pdf_file in self.pdf_files:
            row = tk.Frame(self.list_container, bg="#ffffff", bd=1, relief="solid")
            row.pack(fill="x", padx=10, pady=4)

            file_label = tk.Label(
                row,
                text=pdf_file.name,
                anchor="w",
                width=40,
                padx=8,
                pady=8,
                bg="#ffffff",
            )
            file_label.pack(side="left", expand=True, fill="x")

            duplex_var = tk.StringVar(value="Single-sided")
            copies_var = tk.IntVar(value=1)

            duplex_combo = ttk.Combobox(
                row,
                textvariable=duplex_var,
                values=["Single-sided", "Double-sided"],
                state="readonly",
                width=15,
            )
            duplex_combo.pack(side="left", padx=(0, 10), pady=6)

            copies_spin = tk.Spinbox(
                row, from_=1, to=99, textvariable=copies_var, width=6
            )
            copies_spin.pack(side="left", padx=(0, 8), pady=6)

            self.file_settings[str(pdf_file)] = {
                "duplex": duplex_var,
                "copies": copies_var,
            }

        self.btn_toggle_duplex.config(state="normal")
        self.list_container.update_idletasks()
        self._update_canvas_scrollregion()

    def get_printers(self):
        """Return the list of available Windows printers."""
        if win32print is None:
            return ["Default Printer"]

        printers = []
        seen_printers = set()

        def log_printers(label, names):
            self.log(f"{label}:")
            if not names:
                self.log("  (none found)")
                return
            for name in names:
                self.log(f"  {name}")

        try:
            local_flags = (
                win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
            )
            local_entries = win32print.EnumPrinters(local_flags, None, 2)
            local_printers = []
            for printer_info in local_entries:
                printer_name = printer_info["pPrinterName"]
                if printer_name:
                    local_printers.append(printer_name)
                    if printer_name not in seen_printers:
                        printers.append(printer_name)
                        seen_printers.add(printer_name)
            log_printers("LOCAL / CONNECTIONS printers", local_printers)
        except Exception as exc:
            self.log(f"Unable to list LOCAL / CONNECTIONS printers: {exc}")

        try:
            network_entries = win32print.EnumPrinters(
                win32print.PRINTER_ENUM_NETWORK, None, 1
            )
            network_printers = []
            for printer_info in network_entries:
                printer_name = printer_info[2]
                if printer_name:
                    network_printers.append(printer_name)
                    if printer_name not in seen_printers:
                        printers.append(printer_name)
                        seen_printers.add(printer_name)
            log_printers("NETWORK printers", network_printers)
        except Exception as exc:
            self.log(f"Unable to list NETWORK printers: {exc}")

        return printers if printers else ["Default Printer"]

    def log(self, message):
        """Append messages to the log box."""
        if self.log_box is None:
            self._pending_log_messages.append(message)
            return

        self.log_box.config(state="normal")
        self.log_box.insert(tk.END, message + "\n")
        self.log_box.see(tk.END)
        self.log_box.config(state="disabled")

    def print_pdfs(self):
        """Print each selected PDF using SumatraPDF and its per-file settings."""
        self.log("=== Print button clicked ===")
        path_str = self.folder_path.get()
        if not path_str:
            self.log("Print cancelled: no folder is selected.")
            messagebox.showerror("Error", "Please select a folder first!")
            return

        selected_printer = self.printer_name.get() or "Default Printer"
        self.log(f"Printer selected in the app: {selected_printer}")
        default_printer = None
        if win32print is not None:
            try:
                default_printer = win32print.GetDefaultPrinter()
                self.log(f"Windows default printer: {default_printer}")
            except Exception as exc:
                self.log(
                    f"Unable to read the Windows default printer "
                    f"({type(exc).__name__}): {exc}"
                )
        else:
            self.log("Windows printer APIs are unavailable (pywin32 not installed).")

        printer_choice = (
            None if selected_printer == "Default Printer" else selected_printer
        )
        if selected_printer != "Default Printer" and not printer_choice:
            self.log("Print cancelled: no usable printer was resolved.")
            messagebox.showerror("Error", "Please select a printer first!")
            return

        if not self.file_settings:
            self.log(f"Print cancelled: no PDFs are loaded from {path_str}")
            messagebox.showwarning(
                "No PDFs Found", f"No PDF files were found in:\n{path_str}"
            )
            return

        sumatra_exe = self._find_sumatra_pdf()
        if not sumatra_exe:
            self.log("SumatraPDF executable was not found on this system.")
            messagebox.showerror(
                "SumatraPDF Required",
                "SumatraPDF was not found. Install it, then try printing again.",
            )
            return

        if selected_printer == "Default Printer":
            self.log(
                f"Resolved printer for this job: {default_printer or 'Windows default'}"
            )
            job_printer_label = default_printer or "Windows default"
        else:
            self.log(f"Resolved printer for this job: {printer_choice}")
            job_printer_label = printer_choice
        self.log(f"PDF print engine: {sumatra_exe}")
        self.log(f"Folder: {path_str}")
        self.log(f"--- Starting print job ({len(self.file_settings)} files) ---")

        success_count = 0
        failure_count = 0
        for pdf_path, settings in self.file_settings.items():
            pdf_file = Path(pdf_path)
            try:
                duplex = settings["duplex"].get()
                copies_value = settings["copies"].get()
                self.log(f"File path: {pdf_file}")
                self.log(
                    f"Requested settings: duplex={duplex}, copies={copies_value}"
                )
                copies = int(copies_value)
                if not 1 <= copies <= 99:
                    raise ValueError("Copies must be between 1 and 99.")
                if duplex not in ("Single-sided", "Double-sided"):
                    raise ValueError(f"Unsupported duplex option: {duplex}")
                self.root.update_idletasks()

                duplex_setting = (
                    "simplex" if duplex == "Single-sided" else "duplexlong"
                )
                print_settings = (
                    f"{copies}x,{duplex_setting},ignore-pdf-print-settings"
                )
                command = [sumatra_exe]
                if selected_printer == "Default Printer":
                    command.append("-print-to-default")
                else:
                    command.extend(["-print-to", printer_choice])
                command.extend(
                    ["-print-settings", print_settings, str(pdf_file)]
                )
                self.log(f"Submitting SumatraPDF command: {command!r}")
                result = subprocess.run(
                    command,
                    capture_output=True,
                    text=True,
                )
                if result.stdout.strip():
                    self.log(f"SumatraPDF output: {result.stdout.strip()}")
                if result.stderr.strip():
                    self.log(f"SumatraPDF messages: {result.stderr.strip()}")
                if result.returncode != 0:
                    raise subprocess.CalledProcessError(
                        result.returncode,
                        command,
                        output=result.stdout,
                        stderr=result.stderr,
                    )
                success_count += 1
                self.log(
                    f"SumatraPDF finished printing {pdf_file.name} "
                    f"(exit code {result.returncode})."
                )
            except Exception as exc:
                failure_count += 1
                self.log(
                    f"Failed to print {pdf_file.name} "
                    f"({type(exc).__name__}): {exc}"
                )

        self.log(
            f"--- Completed: {success_count} succeeded, {failure_count} failed; "
            f"intended printer was "
            f"'{job_printer_label}' ---"
        )
        summary = (
            f"Completed {success_count} PDF print job(s); "
            f"{failure_count} failed. Check the System Log for details."
        )
        if failure_count:
            messagebox.showwarning("Printing completed with errors", summary)
        else:
            messagebox.showinfo("Printing complete", summary)


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFPrinterApp(root)
    root.mainloop()
