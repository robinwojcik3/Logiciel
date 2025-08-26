"""Point d'entrée léger de l'application."""

import tkinter as tk
from app import MainApp


def main() -> None:
    """Lancer l'interface utilisateur."""
    root = tk.Tk()
    MainApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
