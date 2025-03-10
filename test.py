
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

# # Exemple d'utilisation
# if __name__ == "__main__":
#     printer = Printer(printer_name="Your Printer Name")
#     try:
#         printer.print_pdf("path_to_your_pdf_file.pdf", page_range="1-3,5", password="your_password")
#     except Exception as e:
#         print(f"An error occurred: {str(e)}")
