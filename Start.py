"""Point d'entrée léger de l'application.

Ce script ne fait que lancer l'interface graphique et déléguer
les traitements lourds à des modules spécialisés.
"""

import tkinter as tk
from tkinter import ttk


def open_context_eco() -> None:
    """Ouvre le module d'analyse du contexte écologique."""
    from modules.context_eco import analyse_zonages
    # Les chemins sont fournis ici à titre d'exemple.
    analyse_zonages("AE.shp", "ZE.shp")


def open_remonter() -> None:
    """Lance le module « Remonter le temps »."""
    from modules.remonter_temps import run_workflow
    run_workflow()


def open_plantnet() -> None:
    """Lance l'identification de plantes via l'API Pl@ntNet."""
    from modules.plantnet import identify_plant
    identify_plant("image.jpg", "leaf")


def main() -> None:
    """Crée l'interface principale et les onglets."""
    root = tk.Tk()
    root.title("Start")
    notebook = ttk.Notebook(root)
    notebook.pack(expand=True, fill="both")

    tab1 = ttk.Frame(notebook)
    ttk.Button(tab1, text="Lancer", command=open_context_eco).pack(padx=20, pady=20)
    notebook.add(tab1, text="Contexte éco")

    tab2 = ttk.Frame(notebook)
    ttk.Button(tab2, text="Lancer", command=open_remonter).pack(padx=20, pady=20)
    notebook.add(tab2, text="Remonter le temps")

    tab3 = ttk.Frame(notebook)
    ttk.Button(tab3, text="Lancer", command=open_plantnet).pack(padx=20, pady=20)
    notebook.add(tab3, text="Pl@ntNet")

    root.mainloop()


if __name__ == "__main__":
    main()
