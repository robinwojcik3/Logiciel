"""Fonctions liées à l'identification des plantes via l'API Pl@ntNet."""

import os
import shutil
from io import BytesIO

import requests
from PIL import Image

from .utils import log_with_time

API_KEY = "2b10vfT6MvFC2lcAzqG1ZMKO"  # clé d'exemple
PROJECT = "all"
API_URL = f"https://my-api.plantnet.org/v2/identify/{PROJECT}?api-key={API_KEY}"


def resize_image(image_path: str, max_size=(800, 800), quality: int = 70) -> BytesIO | None:
    """Redimensionne et compresse une image avant l'envoi."""
    try:
        with Image.open(image_path) as img:
            img.thumbnail(max_size)
            buffer = BytesIO()
            img.save(buffer, format="JPEG", quality=quality)
            buffer.seek(0)
            return buffer
    except Exception as exc:  # pragma: no cover - dépend des fichiers
        log_with_time(f"Erreur lors du redimensionnement : {exc}")
        return None


def identify_plant(image_path: str, organ: str) -> str | None:
    """Interroge l'API Pl@ntNet et renvoie le nom scientifique."""
    log_with_time(f"Envoi de l'image à l'API : {image_path}")
    resized = resize_image(image_path)
    if not resized:
        return None

    files = {"images": (os.path.basename(image_path), resized, "image/jpeg")}
    data = {"organs": organ}
    response = requests.post(API_URL, files=files, data=data)
    if response.status_code != 200:
        log_with_time(f"Erreur API : {response.status_code}")
        return None
    try:
        return response.json()["results"][0]["species"]["scientificNameWithoutAuthor"]
    except Exception:  # pragma: no cover - dépend des réponses API
        log_with_time("Aucun résultat trouvé")
        return None


def copy_and_rename_file(file_path: str, dest_folder: str, new_name: str, count: int) -> None:
    """Copie un fichier identifié dans le dossier de destination."""
    ext = os.path.splitext(file_path)[1]
    suffix = f"({count})" if count > 1 else ""
    new_file = f"{new_name} @plantnet{suffix}{ext}"
    try:
        shutil.copy(file_path, os.path.join(dest_folder, new_file))
        log_with_time(f"Fichier copié sous {new_file}")
    except Exception as exc:  # pragma: no cover - dépend du système de fichiers
        log_with_time(f"Erreur lors de la copie : {exc}")
