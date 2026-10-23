import os
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import win32con
    import win32print
except ImportError:
    win32con = None
    win32print = None


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

        list_title = tk.Label(
            root, text="PDF Files to Print", font=("Arial", 10, "bold")
        )
        list_title.pack(anchor="w", padx=20)

        self.list_canvas = tk.Canvas(root, height=220, bg="#f5f5f5")
        self.list_canvas.pack(
            side="left", fill="both", expand=True, padx=(20, 0), pady=(5, 10)
        )

        self.list_scrollbar = ttk.Scrollbar(
            root, orient="vertical", command=self.list_canvas.yview
        )
        self.list_scrollbar.pack(side="right", fill="y", padx=(0, 20), pady=(5, 10))
        self.list_canvas.configure(yscrollcommand=self.list_scrollbar.set)

        self.list_container = tk.Frame(self.list_canvas, bg="#f5f5f5")
        self.list_canvas.create_window((0, 0), window=self.list_container, anchor="nw")
        self.list_container.bind("<Configure>", self._update_canvas_scrollregion)

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

        self.list_container.update_idletasks()
        self._update_canvas_scrollregion()

    def get_printers(self):
        """Return the list of available Windows printers."""
        if win32print is None:
            return ["Default Printer"]

        try:
            flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS

            levels = win32print.EnumPrinters(flags, None, 2)
            printers = []
            for printer_info in levels:
                printer_name = printer_info["pPrinterName"]
                if printer_name:
                    printers.append(printer_name)

            return printers if printers else ["Default Printer"]
        except Exception as exc:
            self.log(f"Unable to list printers: {exc}")
            return ["Default Printer"]

    def log(self, message):
        """Append messages to the log box."""
        if self.log_box is None:
            self._pending_log_messages.append(message)
            return

        self.log_box.config(state="normal")
        self.log_box.insert(tk.END, message + "\n")
        self.log_box.see(tk.END)
        self.log_box.config(state="disabled")

    def _update_printer_settings(self, duplex, copies, printer_name):
        """Try to apply duplex and copies to the selected Windows printer when pywin32 is available."""
        if win32print is None or win32con is None:
            return

        try:
            handle = win32print.OpenPrinter(printer_name)
            printer_info = win32print.GetPrinter(handle, 2)
            devmode = printer_info["pDevMode"]

            if duplex == "Single-sided":
                devmode.Duplex = win32con.DMDUP_SIMPLEX
            else:
                devmode.Duplex = win32con.DMDUP_VERTICAL

            devmode.Copies = int(copies)
            printer_info["pDevMode"] = devmode
            win32print.SetPrinter(handle, 2, printer_info, 0)
            win32print.ClosePrinter(handle)
        except Exception as exc:
            self.log(f"Printer setting warning for '{printer_name}': {exc}")

    def print_pdfs(self):
        """Print each selected PDF using its own duplex and copy settings."""
        path_str = self.folder_path.get()
        if not path_str:
            messagebox.showerror("Error", "Please select a folder first!")
            return

        printer_choice = self.printer_name.get() or "Default Printer"
        if printer_choice == "Default Printer":
            printer_choice = win32print.GetDefaultPrinter() if win32print else None

        if not printer_choice:
            messagebox.showerror("Error", "Please select a printer first!")
            return

        if not self.file_settings:
            messagebox.showwarning(
                "No PDFs Found", f"No PDF files were found in:\n{path_str}"
            )
            return

        self.log(f"Selected printer: {printer_choice}")
        self.log(f"--- Starting print job ({len(self.file_settings)} files) ---")

        success_count = 0
        for pdf_path, settings in self.file_settings.items():
            pdf_file = Path(pdf_path)
            duplex = settings["duplex"].get()
            copies = max(1, int(settings["copies"].get()))

            try:
                self.log(f"Printing: {pdf_file.name} | {duplex} | Copies: {copies}")
                self.root.update_idletasks()

                self._update_printer_settings(duplex, copies, printer_choice)
                os.startfile(str(pdf_file), "print")
                success_count += 1
            except Exception as exc:
                self.log(f"Failed to print {pdf_file.name}: {exc}")

        self.log(
            f"--- Completed: Sent {success_count} file(s) to printer '{printer_choice}' ---"
        )
        messagebox.showinfo(
            "Success",
            f"Successfully sent {success_count} PDF(s) to printer '{printer_choice}'.",
        )


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFPrinterApp(root)
    root.mainloop()
