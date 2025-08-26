import os
import shutil
from io import BytesIO

import requests
from PIL import Image
import pillow_heif

API_KEY = "2b10vfT6MvFC2lcAzqG1ZMKO"
PROJECT = "all"
API_URL = f"https://my-api.plantnet.org/v2/identify/{PROJECT}?api-key={API_KEY}"

pillow_heif.register_heif_opener()

def resize_image(image_path, max_size=(800, 800), quality=70):
    """Redimensionne et compresse une image."""
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
    """Envoie une image à l'API Pl@ntNet pour identification."""
    print(f"Envoi de l'image à l'API : {image_path}")
    try:
        resized_image = resize_image(image_path)
        if not resized_image:
            print(f"Échec du redimensionnement de l'image : {image_path}")
            return None

        files = {'images': (os.path.basename(image_path), resized_image, 'image/jpeg')}
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
    """Copie et renomme un fichier dans le dossier de destination."""
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
