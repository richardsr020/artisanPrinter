import os
import win32print
import win32api
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

class ImprimanteApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Impression de fichier")

        # Variables
        self.nom_fichier = tk.StringVar()
        self.printer_selected = tk.StringVar()
        self.page_debut = tk.StringVar(value="1")
        self.page_fin = tk.StringVar(value="1")
        self.nb_copies = tk.IntVar(value=1)

        # Sélection du fichier
        tk.Label(root, text="Fichier à imprimer :").grid(row=0, column=0, sticky="w")
        tk.Entry(root, textvariable=self.nom_fichier, width=40).grid(row=0, column=1)
        tk.Button(root, text="Parcourir", command=self.selectionner_fichier).grid(row=0, column=2)

        # Sélection de l'imprimante
        tk.Label(root, text="Sélectionner l'imprimante :").grid(row=1, column=0, sticky="w")
        self.combo_imprimante = ttk.Combobox(root, textvariable=self.printer_selected)
        self.combo_imprimante['values'] = self.get_liste_imprimantes()
        self.combo_imprimante.grid(row=1, column=1, columnspan=2, sticky="we")

        # Plage de pages
        tk.Label(root, text="Pages à imprimer (ex: 1-5) :").grid(row=2, column=0, sticky="w")
        tk.Entry(root, textvariable=self.page_debut, width=5).grid(row=2, column=1, sticky="w")
        tk.Label(root, text="à").grid(row=2, column=1)
        tk.Entry(root, textvariable=self.page_fin, width=5).grid(row=2, column=1, sticky="e")

        # Nombre de copies
        tk.Label(root, text="Nombre de copies :").grid(row=3, column=0, sticky="w")
        tk.Spinbox(root, from_=1, to=100, textvariable=self.nb_copies).grid(row=3, column=1, sticky="w")

        # Bouton d'impression
        tk.Button(root, text="Imprimer", command=self.imprimer_fichier).grid(row=4, column=1, pady=10)

    def selectionner_fichier(self):
        fichier = filedialog.askopenfilename(title="Sélectionner un fichier à imprimer")
        if fichier:
            self.nom_fichier.set(fichier)

    def get_liste_imprimantes(self):
        imprimantes = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 1)
        return [imprimante[2] for imprimante in imprimantes]  # Extraire le nom de l'imprimante

    def imprimer_fichier(self):
        fichier = self.nom_fichier.get()
        imprimante = self.printer_selected.get()
        copies = self.nb_copies.get()

        # Vérifications
        if not fichier or not os.path.exists(fichier):
            messagebox.showerror("Erreur", "Veuillez sélectionner un fichier valide.")
            return
        if not imprimante:
            messagebox.showerror("Erreur", "Veuillez sélectionner une imprimante.")
            return

        try:
            # Envoyer le fichier à l'imprimante
            for _ in range(copies):
                self.envoyer_a_imprimante(fichier, imprimante)
            messagebox.showinfo("Succès", "Le fichier a été envoyé à l'imprimante.")
        except Exception as e:
            messagebox.showerror("Erreur d'impression", str(e))

    def envoyer_a_imprimante(self, fichier, imprimante):
        try:
            # Ouvrir le fichier
            with open(fichier, 'rb') as f:
                data = f.read()

            # Ouvrir l'imprimante
            hprinter = win32print.OpenPrinter(imprimante)
            try:
                # Commencer un travail d'impression
                job = win32print.StartDocPrinter(hprinter, 1, (fichier, None, "RAW"))
                win32print.StartPagePrinter(hprinter)
                win32print.WritePrinter(hprinter, data)
                win32print.EndPagePrinter(hprinter)
                win32print.EndDocPrinter(hprinter)
            finally:
                win32print.ClosePrinter(hprinter)
        except Exception as e:
            raise e

# Création de la fenêtre principale
root = tk.Tk()
app = ImprimanteApp(root)
root.mainloop()