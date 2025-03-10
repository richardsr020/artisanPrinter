import tkinter as tk

# Importez la classe PrinterApp
from utils.winPrinter import *  # Assurez-vous que le fichier contenant la classe s'appelle "printer_app.py" ou ajustez le nom

# Fonction principale pour démarrer l'application
def main():
    # Créez la fenêtre principale Tkinter
    root = tk.Tk()

    # Créez une instance de l'application d'impression
    app = PrinterApp(root)

    # Lancer l'interface graphique
    root.mainloop()

# Exécutez la fonction principale lorsque ce fichier est lancé
if __name__ == "__main__":
    main()
