import os
import json
import fitz
import win32print
import tkinter as tk
from tkinter import PhotoImage, ttk, messagebox
from io import BytesIO

class PrinterApp:
    def __init__(self, root):
        self.root = root
        icon = PhotoImage(file="icons/icon0.png")  # Chargement de l'icône
        self.root.iconphoto(False, icon)  # Définir l'icône
        self.root.title("artisanPrint")
        self.root.resizable(False, False)  # Empêcher le redimensionnement

        # Variables
        self.file_path = tk.StringVar()
        self.selected_printer = tk.StringVar()
        self.page_range = tk.StringVar(value="all")
        self.password = tk.StringVar()

        # Load configuration
        self.load_config()

        # UI Components
        self.create_widgets()

    def create_widgets(self):
        # Printer selection
        ttk.Label(self.root, text="Printer:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.printer_combo = ttk.Combobox(self.root, textvariable=self.selected_printer, width=40)
        self.printer_combo['values'] = self.get_available_printers()
        self.printer_combo.grid(row=0, column=1, padx=5, pady=5, sticky="we")

        # Page range
        ttk.Label(self.root, text="Page Range (ex: 1-3,5 or 'all'):").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(self.root, textvariable=self.page_range, width=40).grid(row=1, column=1, padx=5, pady=5, sticky="we")

        # Print button with icon
        print_icon = PhotoImage(file="icons/icon6.png")
        self.print_button = ttk.Button(self.root, text=" Print PDF", image=print_icon, compound="left", command=self.print_pdf, style="Dark.TButton")
        self.print_button.image = print_icon  # Keep a reference
        self.print_button.grid(row=2, column=1, padx=5, pady=10, sticky="e")

        # Style for dark button
        style = ttk.Style()
        style.configure("Dark.TButton", background="#333", foreground="white", font=("Arial", 10, "bold"))

    def get_available_printers(self):
        printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 1)
        return [printer[2] for printer in printers]

    def load_config(self):
        config_file = "printer.json"
        if os.path.exists(config_file):
            try:
                with open(config_file, "r") as f:
                    config = json.load(f)
                    self.file_path.set(config.get("file_path", ""))
                    self.password.set(config.get("password", ""))
                open(config_file, "w").close()
            except json.JSONDecodeError:
                messagebox.showerror("Error", "Invalid JSON configuration file")

    def parse_page_range(self, total_pages):
        page_range = self.page_range.get().lower().strip()
        if page_range in ["", "all"]:
            return list(range(1, total_pages + 1))

        pages = set()
        parts = page_range.split(',')
        
        for part in parts:
            part = part.strip()
            if '-' in part:
                start, end = part.split('-', 1)
                try:
                    start = int(start)
                    end = int(end)
                    if 1 <= start <= end <= total_pages:
                        pages.update(range(start, end + 1))
                except ValueError:
                    pass
            else:
                try:
                    page = int(part)
                    if 1 <= page <= total_pages:
                        pages.add(page)
                except ValueError:
                    pass

        return sorted(pages) if pages else list(range(1, total_pages + 1))

    def print_pdf(self):
        file_path = self.file_path.get()
        printer_name = self.selected_printer.get()
        password = self.password.get()

        # Validation
        if not all([file_path, os.path.exists(file_path), file_path.lower().endswith('.pdf')]):
            messagebox.showerror("Error", "Invalid or missing PDF file")
            return

        try:
            # Open PDF with Fitz
            doc = fitz.open(file_path)
            
            # Handle encryption
            if doc.is_encrypted:
                if not doc.authenticate(password):
                    raise ValueError("Incorrect password or password required")

            total_pages = doc.page_count
            selected_pages = self.parse_page_range(total_pages)

            # Create in-memory PDF
            buffer = BytesIO()
            new_doc = fitz.open()
            
            # Copy selected pages
            for pg in selected_pages:
                new_doc.insert_pdf(doc, from_page=pg-1, to_page=pg-1)
            
            # Save to buffer
            new_doc.save(buffer)
            pdf_data = buffer.getvalue()
            new_doc.close()
            doc.close()

            # Print raw data
            hprinter = win32print.OpenPrinter(printer_name)
            try:
                win32print.StartDocPrinter(hprinter, 1, ("PDF Print", None, "RAW"))
                win32print.StartPagePrinter(hprinter)
                win32print.WritePrinter(hprinter, pdf_data)
                win32print.EndPagePrinter(hprinter)
                win32print.EndDocPrinter(hprinter)
            finally:
                win32print.ClosePrinter(hprinter)

            messagebox.showinfo("Success", "PDF printed successfully")

        except Exception as e:
            messagebox.showerror("Print Error", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = PrinterApp(root)
    root.mainloop()
