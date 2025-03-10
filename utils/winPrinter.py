import os
import fitz  # PyMuPDF
import win32print
import win32api
import subprocess
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from datetime import datetime
from io import BytesIO  # Importation manquante pour BytesIO

class Printer:
    def __init__(self, printer_name):
        self.printer_name = printer_name

    def parse_page_range(self, total_pages, page_range):
        """Parse le range des pages sélectionnées."""
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

    def print_pdf(self, pdf_path, page_range="all", password=""):
        """Méthode pour imprimer un PDF sur l'imprimante spécifiée"""

        # Validation du fichier PDF
        if not all([pdf_path, os.path.exists(pdf_path), pdf_path.lower().endswith('.pdf')]):
            raise ValueError("Invalid or missing PDF file")
        
        try:
            # Ouvrir le fichier PDF avec Fitz
            doc = fitz.open(pdf_path)

            # Gérer le mot de passe si le PDF est protégé
            if doc.is_encrypted:
                if not doc.authenticate(password):
                    raise ValueError("Incorrect password or password required")

            # Total des pages et pages sélectionnées
            total_pages = doc.page_count
            selected_pages = self.parse_page_range(total_pages, page_range)

            # Créer un PDF en mémoire avec les pages sélectionnées
            buffer = BytesIO()
            new_doc = fitz.open()

            for pg in selected_pages:
                new_doc.insert_pdf(doc, from_page=pg-1, to_page=pg-1)

            # Sauvegarder le PDF dans le buffer
            new_doc.save(buffer)
            pdf_data = buffer.getvalue()
            new_doc.close()
            doc.close()

            # Utiliser win32print pour envoyer le fichier PDF à l'imprimante
            hprinter = win32print.OpenPrinter(self.printer_name)
            try:
                win32print.StartDocPrinter(hprinter, 1, ("PDF Print", None, "RAW"))
                win32print.StartPagePrinter(hprinter)
                win32print.WritePrinter(hprinter, pdf_data)
                win32print.EndPagePrinter(hprinter)
                win32print.EndDocPrinter(hprinter)
            finally:
                win32print.ClosePrinter(hprinter)

            print("PDF printed successfully")

        except Exception as e:
            raise RuntimeError(f"Print Error: {str(e)}")


class PrinterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("artisanPrint")
        self.root.resizable(False, False)

        # Variables
        self.folder_path = tk.StringVar()
        self.selected_printer = tk.StringVar()
        self.page_range = tk.StringVar(value="all")
        self.num_copies = tk.IntVar(value=1)
        self.duplex = tk.BooleanVar(value=False)
        self.sort_option = tk.StringVar(value="None")
        self.pdf_files = []

        # UI Components
        self.create_widgets()
        self.set_default_printer()

    def create_widgets(self):
        ttk.Label(self.root, text="Printer:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.printer_combo = ttk.Combobox(self.root, textvariable=self.selected_printer, width=40)
        self.printer_combo['values'] = self.get_available_printers()
        self.printer_combo.grid(row=0, column=1, padx=5, pady=5, sticky="we")

        ttk.Button(self.root, text="Select Folder", command=self.select_folder).grid(row=1, column=0, padx=5, pady=5, sticky="w")
        ttk.Label(self.root, textvariable=self.folder_path).grid(row=1, column=1, padx=5, pady=5, sticky="we")

        ttk.Label(self.root, text="Sort by:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.sort_combo = ttk.Combobox(self.root, textvariable=self.sort_option, values=["Date Created", "File Size", "Page Count", "None"])
        self.sort_combo.grid(row=2, column=1, padx=5, pady=5, sticky="we")
        self.sort_combo.bind("<<ComboboxSelected>>", lambda e: self.list_pdf_files())

        self.file_listbox = ttk.Treeview(self.root, columns=("Name", "Size", "Pages", "Created At"), show="headings", selectmode="extended")
        self.file_listbox.heading("Name", text="File Name")
        self.file_listbox.heading("Size", text="Size (MB)")
        self.file_listbox.heading("Pages", text="Pages")
        self.file_listbox.heading("Created At", text="Created At")
        self.file_listbox.grid(row=3, column=0, columnspan=2, padx=5, pady=5)

        ttk.Label(self.root, text="Page Range (ex: 1-3,5 or 'all'):").grid(row=4, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(self.root, textvariable=self.page_range, width=40).grid(row=4, column=1, padx=5, pady=5, sticky="we")

        ttk.Label(self.root, text="Number of Copies:").grid(row=5, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(self.root, textvariable=self.num_copies, width=10).grid(row=5, column=1, padx=5, pady=5, sticky="w")

        ttk.Checkbutton(self.root, text="Duplex (Double-sided)", variable=self.duplex).grid(row=6, column=1, padx=5, pady=5, sticky="w")

        ttk.Button(self.root, text="Print Selected PDFs", command=self.print_selected_pdfs).grid(row=7, column=1, padx=5, pady=10, sticky="e")

    def get_available_printers(self):
        printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 1)
        return [printer[2] for printer in printers]

    def set_default_printer(self):
        printers = self.get_available_printers()
        if printers:
            self.selected_printer.set(printers[0])

    def select_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path.set(folder_selected)
            self.list_pdf_files()

    def list_pdf_files(self):
        self.file_listbox.delete(*self.file_listbox.get_children())
        self.pdf_files = []
        folder = self.folder_path.get()
        if not folder:
            return

        for file in os.listdir(folder):
            if file.lower().endswith(".pdf"):
                file_path = os.path.join(folder, file)
                stats = os.stat(file_path)
                size_mb = stats.st_size / (1024 * 1024)
                created = datetime.fromtimestamp(stats.st_ctime).strftime("%d/%m/%Y %H:%M:%S")
                page_count = fitz.open(file_path).page_count
                self.pdf_files.append((file, file_path, created, size_mb, page_count))

        sort_by = self.sort_option.get()
        if sort_by == "Date Created":
            self.pdf_files.sort(key=lambda x: x[2])
        elif sort_by == "File Size":
            self.pdf_files.sort(key=lambda x: x[3], reverse=True)
        elif sort_by == "Page Count":
            self.pdf_files.sort(key=lambda x: x[4], reverse=True)

        for file, _, created, size, pages in self.pdf_files:
            self.file_listbox.insert("", "end", values=(file, f"{size:.2f} MB", pages, created))

    def print_selected_pdfs(self):
        # Récupère les fichiers sélectionnés
        selected_files = [self.file_listbox.item(file)["values"][0] for file in self.file_listbox.selection()]

        if not selected_files:
            messagebox.showerror("Error", "No files selected")
            return

        if not self.selected_printer.get():
            messagebox.showerror("Error", "No printer selected")
            return

        # Message indiquant le début de l'impression
        messagebox.showinfo("Printing", f"Printing {len(selected_files)} PDFs...")

        printer = Printer(self.selected_printer.get())  # Instance de la classe Printer

        # Envoie chaque fichier PDF sélectionné à l'imprimante
        for file in selected_files:
            file_path = os.path.join(self.folder_path.get(), file)
            try:
                # Utilisation de la classe Printer pour imprimer
                printer.print_pdf(file_path, page_range=self.page_range.get())
            except RuntimeError as e:
                messagebox.showerror("Error", f"Failed to print {file}: {str(e)}")
                continue

        # Message de fin
        messagebox.showinfo("Success", f"Printing of {len(selected_files)} PDFs completed.")

# # Lancer l'application Tkinter
# if __name__ == "__main__":
#     root = tk.Tk
