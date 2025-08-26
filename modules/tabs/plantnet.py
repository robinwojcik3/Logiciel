import os
import sys
import threading
import shutil
from io import BytesIO

import requests
import pillow_heif
from PIL import Image
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter import font as tkfont

from ..utils import TextRedirector, ToolTip, save_prefs, OUT_IMG

pillow_heif.register_heif_opener()

API_KEY = "2b10vfT6MvFC2lcAzqG1ZMKO"
PROJECT = "all"
API_URL = f"https://my-api.plantnet.org/v2/identify/{PROJECT}?api-key={API_KEY}"


def resize_image(image_path, max_size=(800, 800), quality=70):
    try:
        with Image.open(image_path) as img:
            img.thumbnail(max_size)
            buffer = BytesIO()
            img.save(buffer, format='JPEG', quality=quality)
            buffer.seek(0)
            return buffer
    except Exception as e:
        print(f"Erreur lors du redimensionnement de l'image : {e}")
        return None


def identify_plant(image_path, organ):
    print(f"Envoi de l'image à l'API : {image_path}")
    try:
        resized_image = resize_image(image_path)
        if not resized_image:
            print(f"Échec du redimensionnement de l'image : {image_path}")
            return None

        files = {
            'images': (os.path.basename(image_path), resized_image, 'image/jpeg')
        }
        data = {'organs': organ}

        response = requests.post(API_URL, files=files, data=data)
        print(f"Réponse de l'API : {response.status_code}")
        if response.status_code == 200:
            json_result = response.json()
            try:
                species = json_result['results'][0]['species']['scientificNameWithoutAuthor']
                print(f"Plante identifiée : {species}")
                return species
            except (KeyError, IndexError):
                print(f"Aucun résultat trouvé pour l'image : {image_path}")
                return None
        else:
            print(f"Erreur API : {response.status_code} - {response.text}")
            return None
    except Exception as e:
        print(f"Exception lors de l'identification de la plante : {e}")
        return None


def copy_and_rename_file(file_path, dest_folder, new_name, count):
    ext = os.path.splitext(file_path)[1]
    if count == 1:
        new_file_name = f"{new_name} @plantnet{ext}"
    else:
        new_file_name = f"{new_name} @plantnet({count}){ext}"
    new_path = os.path.join(dest_folder, new_file_name)
    try:
        shutil.copy(file_path, new_path)
        print(f"Fichier copié et renommé : {file_path} -> {new_path}")
    except Exception as e:
        print(f"Erreur lors de la copie du fichier : {e}")


class PlantNetTab(ttk.Frame):
    def __init__(self, parent, style_helper, prefs: dict):
        super().__init__(parent, padding=12)
        self.style_helper = style_helper
        self.prefs = prefs

        self.font_title = tkfont.Font(family="Segoe UI", size=15, weight="bold")
        self.font_sub = tkfont.Font(family="Segoe UI", size=10)
        self.font_mono = tkfont.Font(family="Consolas", size=9)

        self.folder_var = tk.StringVar(value=self.prefs.get("PLANTNET_FOLDER", ""))
        self._build_ui()

    def _build_ui(self):
        header = ttk.Frame(self, style="Header.TFrame", padding=(14, 12))
        header.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(header, text="Identification Pl@ntNet", style="Card.TLabel", font=self.font_title)\
            .grid(row=0, column=0, sticky="w")
        ttk.Label(header, text="Analyse un dossier d'images via l'API Pl@ntNet.", style="Subtle.TLabel", font=self.font_sub)\
            .grid(row=1, column=0, sticky="w", pady=(4,0))
        header.columnconfigure(0, weight=1)

        card = ttk.Frame(self, style="Card.TFrame", padding=12)
        card.pack(fill=tk.X)
        ttk.Label(card, text="Dossier d'images", style="Card.TLabel").grid(row=0, column=0, sticky="w")
        row = ttk.Frame(card, style="Card.TFrame")
        row.grid(row=0, column=1, sticky="ew", padx=0)
        row.columnconfigure(0, weight=1)
        ttk.Entry(row, textvariable=self.folder_var).grid(row=0, column=0, sticky="ew")
        ttk.Button(row, text="Parcourir…", command=self._pick_folder).grid(row=0, column=1, padx=(6,0))
        card.columnconfigure(1, weight=1)

        act = ttk.Frame(self, style="Card.TFrame", padding=12)
        act.pack(fill=tk.X, pady=(10,0))
        self.run_btn = ttk.Button(act, text="▶ Lancer l'analyse", style="Accent.TButton", command=self._start_thread)
        self.run_btn.grid(row=0, column=0, sticky="w")
        obtn = ttk.Button(act, text="📂 Ouvrir le dossier de sortie", command=self._open_out_dir)
        obtn.grid(row=0, column=1, padx=(10,0))
        ToolTip(obtn, "Ouvrir le dossier cible")

        bottom = ttk.Frame(self, style="Card.TFrame", padding=12)
        bottom.pack(fill=tk.BOTH, expand=True, pady=(10,0))
        self.log_text = tk.Text(bottom, height=12, wrap=tk.WORD, state='disabled',
                                bg=self.style_helper.style.lookup("Card.TFrame", "background"),
                                fg=self.style_helper.style.lookup("TLabel", "foreground"))
        self.log_text.configure(font=self.font_mono, relief="flat")
        log_scroll = ttk.Scrollbar(bottom, orient="vertical", command=self.log_text.yview)
        self.log_text['yscrollcommand'] = log_scroll.set
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.stdout_redirect = TextRedirector(self.log_text)

    def _pick_folder(self):
        base = self.folder_var.get() or os.path.expanduser("~")
        d = filedialog.askdirectory(title="Choisir le dossier d'images",
                                    initialdir=base if os.path.isdir(base) else os.path.expanduser("~"))
        if d:
            self.folder_var.set(d)

    def _open_out_dir(self):
        try:
            os.makedirs(OUT_IMG, exist_ok=True)
            os.startfile(OUT_IMG)
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d’ouvrir le dossier : {e}")

    def _start_thread(self):
        self.run_btn.config(state="disabled")
        t = threading.Thread(target=self._run_process)
        t.daemon = True
        t.start()

    def _run_process(self):
        folder = self.folder_var.get().strip()
        if not folder:
            print("Veuillez sélectionner un dossier.", file=self.stdout_redirect)
            self.after(0, lambda: self.run_btn.config(state="normal"))
            return
        if not os.path.exists(folder):
            print(f"Le dossier à traiter n'existe pas : {folder}", file=self.stdout_redirect)
            self.after(0, lambda: self.run_btn.config(state="normal"))
            return

        self.prefs["PLANTNET_FOLDER"] = folder
        save_prefs(self.prefs)

        image_extensions = ['.jpg', '.jpeg', '.png', '.heic', '.heif']
        image_files = []
        for root, dirs, files in os.walk(folder):
            for f in files:
                if os.path.splitext(f)[1].lower() in image_extensions:
                    if '@plantnet' not in f:
                        image_files.append(os.path.join(root, f))

        if not image_files:
            print("Aucune image à traiter dans le dossier.", file=self.stdout_redirect)
            self.after(0, lambda: self.run_btn.config(state="normal"))
            return

        plant_name_counts = {}
        old_stdout = sys.stdout
        sys.stdout = self.stdout_redirect
        try:
            for image_path in image_files:
                organ = 'flower'
                plant_name = identify_plant(image_path, organ)
                if plant_name:
                    count = plant_name_counts.get(plant_name, 0) + 1
                    plant_name_counts[plant_name] = count
                    copy_and_rename_file(image_path, folder, plant_name, count)
                else:
                    print(f"Aucune identification possible pour l'image : {image_path}")
            print("Analyse terminée.")
        finally:
            sys.stdout = old_stdout
            self.after(0, lambda: self.run_btn.config(state="normal"))
