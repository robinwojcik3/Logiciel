# -*- coding: utf-8 -*-
"""Onglet unique regroupant l'export des cartes et l'ID Contexte éco."""
import os
import sys
import datetime
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter import font as tkfont
from typing import List

from LOGICIEL import (
    StyleHelper, DEFAULT_SHAPE_DIR, OUT_IMG, QGIS_ROOT, QGIS_APP, PY_VER,
    LAYER_AE_NAME, LAYER_ZE_NAME, OVERWRITE_DEFAULT, DPI_DEFAULT,
    N_WORKERS_DEFAULT, MARGIN_FAC_DEFAULT, log_with_time, discover_projects,
    chunk_even, save_prefs, worker_run)


class TextRedirector:
    """Réexpédition du stdout vers un widget Text."""
    def __init__(self, widget):
        self.widget = widget

    def write(self, s):
        self.widget.config(state="normal")
        self.widget.insert(tk.END, s)
        self.widget.see(tk.END)
        self.widget.config(state="disabled")
        self.widget.update_idletasks()

    def flush(self):
        pass


class ContexteEcoTab(ttk.Frame):
    """Onglet combinant export PNG et identification des zonages."""

    def __init__(self, parent, style_helper: StyleHelper, prefs: dict):
        super().__init__(parent, padding=12)
        self.prefs = prefs
        self.style_helper = style_helper

        self.font_mono = tkfont.Font(family="Consolas", size=9)

        # Variables communes
        self.ze_shp_var = tk.StringVar(value=self.prefs.get("ZE_SHP", ""))
        self.ae_shp_var = tk.StringVar(value=self.prefs.get("AE_SHP", ""))
        self.cadrage_var = tk.StringVar(value=self.prefs.get("CADRAGE_MODE", "BOTH"))
        self.overwrite_var = tk.BooleanVar(value=self.prefs.get("OVERWRITE", OVERWRITE_DEFAULT))
        self.dpi_var = tk.IntVar(value=int(self.prefs.get("DPI", DPI_DEFAULT)))
        self.workers_var = tk.IntVar(value=int(self.prefs.get("N_WORKERS", N_WORKERS_DEFAULT)))
        self.margin_var = tk.DoubleVar(value=float(self.prefs.get("MARGIN_FAC", MARGIN_FAC_DEFAULT)))
        self.buffer_var = tk.DoubleVar(value=float(self.prefs.get("ID_TAMPON_KM", 5.0)))

        self.project_vars: dict[str, tk.IntVar] = {}
        self.all_projects: List[str] = []
        self.filtered_projects: List[str] = []
        self.total_expected = 0
        self.progress_done = 0
        self.busy = False

        self._build_ui()
        self._populate_projects()
        self._update_counts()

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        # Sélecteurs shapefile
        shp = ttk.Frame(self, style="Card.TFrame", padding=12)
        shp.pack(fill=tk.X)
        self._file_row(shp, 0, "📁 Zone d'étude…", self.ze_shp_var,
                       lambda: self._select_shapefile('ZE'))
        self._file_row(shp, 1, "📁 Aire d'étude élargie…", self.ae_shp_var,
                       lambda: self._select_shapefile('AE'))
        shp.columnconfigure(1, weight=1)

        # Export cartes (options + projets)
        grid = ttk.Frame(self)
        grid.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        left = ttk.Frame(grid)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        right = ttk.Frame(grid)
        right.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        grid.columnconfigure(0, weight=1)
        grid.columnconfigure(1, weight=1)
        grid.rowconfigure(0, weight=1)

        opt = ttk.Frame(left, style="Card.TFrame", padding=12)
        opt.pack(fill=tk.X)
        ttk.Label(opt, text="Cadrage et options", style="Card.TLabel")\
            .grid(row=0, column=0, columnspan=6, sticky="w")
        ttk.Radiobutton(opt, text="AE + ZE", variable=self.cadrage_var, value="BOTH",
                        style="Card.TRadiobutton").grid(row=1, column=0, sticky="w", pady=(6, 2))
        ttk.Radiobutton(opt, text="ZE uniquement", variable=self.cadrage_var, value="ZE",
                        style="Card.TRadiobutton").grid(row=1, column=1, sticky="w", padx=(12, 0))
        ttk.Radiobutton(opt, text="AE uniquement", variable=self.cadrage_var, value="AE",
                        style="Card.TRadiobutton").grid(row=1, column=2, sticky="w", padx=(12, 0))
        ttk.Checkbutton(opt, text="Écraser si le PNG existe", variable=self.overwrite_var,
                        style="Card.TCheckbutton").grid(row=1, column=3, sticky="w", padx=(24, 0))
        ttk.Label(opt, text="DPI", style="Card.TLabel").grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Spinbox(opt, from_=72, to=1200, textvariable=self.dpi_var, width=6,
                    justify="right").grid(row=2, column=1, sticky="w", pady=(8, 0))
        ttk.Label(opt, text="Workers", style="Card.TLabel").grid(row=2, column=2, sticky="w",
                                                                  padx=(12, 0), pady=(8, 0))
        ttk.Spinbox(opt, from_=1, to=max(1, (os.cpu_count() or 2)),
                    textvariable=self.workers_var, width=6, justify="right").grid(row=2, column=3,
                                                                                   sticky="w", pady=(8, 0))
        ttk.Label(opt, text="Marge", style="Card.TLabel").grid(row=2, column=4, sticky="w",
                                                                 padx=(12, 0), pady=(8, 0))
        ttk.Spinbox(opt, from_=1.00, to=2.00, increment=0.05, textvariable=self.margin_var,
                    width=6, justify="right").grid(row=2, column=5, sticky="w", pady=(8, 0))

        act = ttk.Frame(left, style="Card.TFrame", padding=12)
        act.pack(fill=tk.X, pady=(10, 0))
        self.export_button = ttk.Button(act, text="▶ Lancer l’export cartes",
                                        style="Accent.TButton",
                                        command=self.start_export_thread)
        self.export_button.grid(row=0, column=0, sticky="w")

        proj = ttk.Frame(right, style="Card.TFrame", padding=12)
        proj.pack(fill=tk.BOTH, expand=True)
        ttk.Label(proj, text="Projets QGIS", style="Card.TLabel")\
            .grid(row=0, column=0, columnspan=4, sticky="w")
        ttk.Label(proj, text="Filtrer", style="Card.TLabel").grid(row=1, column=0, sticky="w",
                                                                    pady=(6, 6))
        self.filter_var = tk.StringVar()
        fe = ttk.Entry(proj, textvariable=self.filter_var, width=32)
        fe.grid(row=1, column=1, sticky="w", pady=(6, 6))
        fe.bind("<KeyRelease>", lambda _e: self._apply_filter())
        ttk.Button(proj, text="Tout", width=6,
                   command=lambda: self._select_all(True)).grid(row=1, column=2, padx=(8, 0))
        ttk.Button(proj, text="Aucun", width=6,
                   command=lambda: self._select_all(False)).grid(row=1, column=3, padx=(6, 0))

        canvas = tk.Canvas(proj, highlightthickness=0, borderwidth=0)
        scrollbar = ttk.Scrollbar(proj, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)
        self.scrollable_frame.bind("<Configure>",
                                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.grid(row=2, column=0, columnspan=4, sticky="nsew", pady=(6, 6))
        scrollbar.grid(row=2, column=4, sticky="ns", padx=(6, 0))
        proj.rowconfigure(2, weight=1)
        proj.columnconfigure(1, weight=1)

        # ID Contexte éco
        id_card = ttk.Frame(self, style="Card.TFrame", padding=12)
        id_card.pack(fill=tk.X, pady=(10, 0))
        ttk.Label(id_card, text="Tampon zone d'étude (km)", style="Card.TLabel")\
            .grid(row=0, column=0, sticky="w")
        ttk.Spinbox(id_card, from_=0.0, to=50.0, increment=0.5,
                    textvariable=self.buffer_var, width=6, justify="right").grid(row=0, column=1,
                                                                                  sticky="w", padx=(8, 0))
        self.id_button = ttk.Button(id_card, text="▶ Lancer l’ID Contexte éco",
                                    style="Accent.TButton",
                                    command=self.start_id_thread)
        self.id_button.grid(row=0, column=2, padx=(12, 0))
        id_card.columnconfigure(1, weight=1)

        # Console / progression
        bottom = ttk.Frame(self, style="Card.TFrame", padding=12)
        bottom.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        self.status_label = ttk.Label(bottom, text="Prêt.", style="Status.TLabel")
        self.status_label.grid(row=0, column=0, sticky="w")
        self.progress = ttk.Progressbar(bottom, orient="horizontal", mode="determinate", length=220)
        self.progress.grid(row=0, column=1, sticky="e")
        bottom.columnconfigure(0, weight=1)

        log_frame = ttk.Frame(bottom, style="Card.TFrame")
        log_frame.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(8, 0))
        bottom.rowconfigure(1, weight=1)
        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD, state='disabled',
                                bg=self.style_helper.style.lookup("Card.TFrame", "background"),
                                fg=self.style_helper.style.lookup("TLabel", "foreground"))
        self.log_text.configure(font=self.font_mono, relief="flat")
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text['yscrollcommand'] = log_scroll.set
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.stdout_redirect = TextRedirector(self.log_text)

    # ------------------------------------------------------------------ utilitaires UI
    def _file_row(self, parent, row: int, label: str, var: tk.StringVar, cmd):
        btn = ttk.Button(parent, text=label, command=cmd)
        btn.grid(row=row, column=0, sticky="w", pady=(8 if row == 0 else 4, 2))
        ent = ttk.Entry(parent, textvariable=var, width=10)
        ent.grid(row=row, column=1, sticky="ew", padx=8)
        ent.configure(state="readonly")
        copy_btn = ttk.Button(parent, text="Copier", width=6,
                               command=lambda: self._copy_to_clipboard(var.get()))
        copy_btn.grid(row=row, column=2, sticky="e")
        clear_btn = ttk.Button(parent, text="✖", width=3, command=lambda: var.set(""))
        clear_btn.grid(row=row, column=3, sticky="e")
        parent.columnconfigure(1, weight=1)

    def _copy_to_clipboard(self, text: str):
        try:
            self.winfo_toplevel().clipboard_clear()
            self.winfo_toplevel().clipboard_append(text)
        except Exception:
            pass

    def _select_shapefile(self, shp_type):
        label_text = "Zone d'étude" if shp_type == 'ZE' else "Aire d'étude élargie"
        title = f"Sélectionner le shapefile pour '{label_text}'"
        base_dir = DEFAULT_SHAPE_DIR if os.path.isdir(DEFAULT_SHAPE_DIR) else os.path.expanduser("~")
        path = filedialog.askopenfilename(title=title, initialdir=base_dir,
                                          filetypes=[("Shapefile ESRI", "*.shp")])
        if path:
            if shp_type == 'ZE':
                self.ze_shp_var.set(path)
            else:
                self.ae_shp_var.set(path)

    # ------------------------------------------------------------------ projets QGIS
    def _populate_projects(self):
        for w in list(self.scrollable_frame.children.values()):
            w.destroy()
        self.project_vars.clear()
        self.all_projects = discover_projects()
        self.filtered_projects = list(self.all_projects)
        if not self.all_projects:
            ttk.Label(self.scrollable_frame,
                      text="Aucun projet trouvé ou dossier inaccessible.",
                      foreground="red").pack(anchor="w")
            return
        for proj_path in self.filtered_projects:
            var = tk.IntVar(value=1)
            self.project_vars[proj_path] = var
            ttk.Checkbutton(self.scrollable_frame, text=os.path.basename(proj_path),
                            variable=var, style="Card.TCheckbutton")\
                .pack(anchor='w', padx=4, pady=1)

    def _apply_filter(self):
        term = self.filter_var.get().lower()
        for w in list(self.scrollable_frame.children.values()):
            w.destroy()
        self.filtered_projects = [p for p in self.all_projects
                                  if term in os.path.basename(p).lower()]
        if not self.filtered_projects:
            ttk.Label(self.scrollable_frame,
                      text="Aucun projet ne correspond au filtre.",
                      foreground="red").pack(anchor="w")
            self.project_vars = {}
            self._update_counts()
            return
        for proj_path in self.filtered_projects:
            current = self.project_vars.get(proj_path, tk.IntVar(value=1))
            self.project_vars[proj_path] = current
            ttk.Checkbutton(self.scrollable_frame, text=os.path.basename(proj_path),
                            variable=current, style="Card.TCheckbutton")\
                .pack(anchor='w', padx=4, pady=1)
        self._update_counts()

    def _select_all(self, state: bool):
        for var in self.project_vars.values():
            var.set(1 if state else 0)
        self._update_counts()

    def _selected_projects(self) -> List[str]:
        return [p for p, v in self.project_vars.items()
                if v.get() == 1 and p in self.filtered_projects]

    def _update_counts(self):
        selected = len(self._selected_projects())
        total = len(self.filtered_projects)
        self.status_label.config(text=f"Projets sélectionnés : {selected} / {total}")

    # ------------------------------------------------------------------ export cartes
    def start_export_thread(self):
        if self.busy:
            print("Une action est déjà en cours.", file=self.stdout_redirect)
            return
        ze = self.ze_shp_var.get().strip()
        ae = self.ae_shp_var.get().strip()
        if not ze or not ae:
            messagebox.showerror("Erreur", "Sélectionnez les deux shapefiles.")
            return
        if not os.path.isfile(ze) or not os.path.isfile(ae):
            messagebox.showerror("Erreur", "Un shapefile est introuvable.")
            return
        projets = self._selected_projects()
        if not projets:
            messagebox.showerror("Erreur", "Sélectionnez au moins un projet.")
            return

        self.busy = True
        self.export_button.config(state="disabled")
        self.id_button.config(state="disabled")
        self.progress_done = 0
        mode = self.cadrage_var.get()
        per_project = 2 if mode == "BOTH" else 1
        self.total_expected = per_project * len(projets)
        self.progress.config(mode="determinate",
                             maximum=max(1, self.total_expected), value=0)
        self.status_label.config(text="Export en cours…")

        self.prefs.update({
            "ZE_SHP": ze,
            "AE_SHP": ae,
            "CADRAGE_MODE": self.cadrage_var.get(),
            "OVERWRITE": bool(self.overwrite_var.get()),
            "DPI": int(self.dpi_var.get()),
            "N_WORKERS": int(self.workers_var.get()),
            "MARGIN_FAC": float(self.margin_var.get()),
        })
        save_prefs(self.prefs)

        t = threading.Thread(target=self._run_export_logic, args=(projets,), daemon=True)
        t.start()

    def _run_export_logic(self, projets: List[str]):
        old_stdout = sys.stdout
        sys.stdout = self.stdout_redirect
        try:
            start = datetime.datetime.now()
            os.makedirs(OUT_IMG, exist_ok=True)

            log_with_time(f"{len(projets)} projets (attendu = calcul en cours)")
            log_with_time(
                f"Workers={self.workers_var.get()}, DPI={self.dpi_var.get()}, "
                f"marge={self.margin_var.get():.2f}, overwrite={self.overwrite_var.get()}"
            )

            chunks = chunk_even(projets, self.workers_var.get())
            cfg = {
                "QGIS_ROOT": QGIS_ROOT, "QGIS_APP": QGIS_APP, "PY_VER": PY_VER,
                "OUT_IMG": OUT_IMG, "DPI": int(self.dpi_var.get()),
                "MARGIN_FAC": float(self.margin_var.get()),
                "LAYER_AE_NAME": LAYER_AE_NAME, "LAYER_ZE_NAME": LAYER_ZE_NAME,
                "AE_SHP": self.ae_shp_var.get(), "ZE_SHP": self.ze_shp_var.get(),
                "CADRAGE_MODE": self.cadrage_var.get(),
                "OVERWRITE": bool(self.overwrite_var.get()),
            }

            ok_total = 0
            ko_total = 0

            def ui_update_progress(done_inc):
                self.progress_done += done_inc
                self.progress['value'] = min(self.progress_done, self.total_expected)
                self.status_label.config(
                    text=f"Progression : {self.progress_done}/{self.total_expected}")

            from concurrent.futures import ProcessPoolExecutor, as_completed

            with ProcessPoolExecutor(max_workers=int(self.workers_var.get())) as ex:
                futures = [ex.submit(worker_run, (chunk, cfg))
                           for chunk in chunks if chunk]
                for fut in as_completed(futures):
                    try:
                        ok, ko = fut.result()
                        ok_total += ok
                        ko_total += ko
                        self.after(0, ui_update_progress, ok + ko)
                        log_with_time(f"Lot terminé: {ok} OK, {ko} KO")
                    except Exception as e:
                        log_with_time(f"Erreur worker: {e}")

            elapsed = datetime.datetime.now() - start
            log_with_time(
                f"FIN — OK={ok_total} | KO={ko_total} | Attendu={self.total_expected} | Durée={elapsed}"
            )
            self.after(0, lambda: self.status_label.config(
                text=f"Terminé — OK={ok_total} / KO={ko_total}"))
        except Exception as e:
            log_with_time(f"Erreur critique: {e}")
            self.after(0, lambda: messagebox.showerror("Erreur", str(e)))
        finally:
            sys.stdout = old_stdout
            self.after(0, self._export_done)

    def _export_done(self):
        self.export_button.config(state="normal")
        self.id_button.config(state="normal")
        self.busy = False

    # ------------------------------------------------------------------ ID contexte
    def start_id_thread(self):
        if self.busy:
            print("Une action est déjà en cours.", file=self.stdout_redirect)
            return
        ze = self.ze_shp_var.get().strip()
        ae = self.ae_shp_var.get().strip()
        if not ze or not ae:
            messagebox.showerror("Erreur", "Sélectionnez les deux shapefiles.")
            return
        if not os.path.isfile(ze) or not os.path.isfile(ae):
            messagebox.showerror("Erreur", "Un shapefile est introuvable.")
            return
        buffer_km = float(self.buffer_var.get())

        self.busy = True
        self.export_button.config(state="disabled")
        self.id_button.config(state="disabled")
        self.progress.config(mode="indeterminate", value=0)
        self.progress.start()
        self.status_label.config(text="Analyse en cours…")

        self.prefs.update({"ZE_SHP": ze, "AE_SHP": ae, "ID_TAMPON_KM": buffer_km})
        save_prefs(self.prefs)

        t = threading.Thread(target=self._run_id_logic, args=(buffer_km,), daemon=True)
        t.start()

    def _run_id_logic(self, buffer_km: float):
        old_stdout = sys.stdout
        sys.stdout = self.stdout_redirect
        try:
            from id_contexte_eco import run_analysis as run_id_context
            run_id_context(self.ae_shp_var.get(), self.ze_shp_var.get(), buffer_km=buffer_km)
            print("Analyse terminée.")
        except Exception as e:
            print(f"Erreur: {e}")
            self.after(0, lambda: messagebox.showerror("Erreur", str(e)))
        finally:
            sys.stdout = old_stdout
            self.after(0, self._id_done)

    def _id_done(self):
        self.progress.stop()
        self.progress.config(mode="determinate", value=0)
        self.status_label.config(text="Terminé.")
        self.export_button.config(state="normal")
        self.id_button.config(state="normal")
        self.busy = False

