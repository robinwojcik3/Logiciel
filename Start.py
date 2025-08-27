"""Point d'entrée léger de l'application.

Ce module se contente d'ouvrir l'interface graphique et délègue les
traitements aux modules spécialisés.
"""


def main() -> None:
    """Lance l'interface principale."""
    import time
    t0 = time.time()
    from modules import main_app
    main_app.launch(start_time=t0)


if __name__ == "__main__":
    main()
