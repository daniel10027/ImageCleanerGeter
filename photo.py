import pandas as pd
import requests
from io import BytesIO
from PIL import Image
import os
from tqdm import tqdm

# Charger le fichier Excel
excel_file = "photos.xlsx"
df = pd.read_excel(excel_file)

# Dossier de sortie
output_folder = os.path.expanduser("~/Téléchargements/photos")
os.makedirs(output_folder, exist_ok=True)

# Fonction pour télécharger, redimensionner et renommer les images
def download_resize_and_rename_image(row):
    matricule = str(row['MATRICULE'])
    nom = row['NOM']
    prenom = row['PRENOMS']
    filiere = row['FILIERE']
    niveau = row['NIVEAU']
    cycle = row['CYCLE']
    contact1 = row['CONTACT 1']
    contact2 = row['CONTACT 2']
    url = row['URL']

    # Télécharger l'image
    response = requests.get(url, stream=True)
    total_size = int(response.headers.get('content-length', 0))

    # Progress Bar avec tqdm
    with tqdm(total=total_size, unit='B', unit_scale=True, desc=f"Downloading {matricule}_{contact1}") as pbar:
        img_data = BytesIO()
        for chunk in response.iter_content(chunk_size=1024):
            img_data.write(chunk)
            pbar.update(len(chunk))

    # Redimensionner l'image
    img = Image.open(img_data)
    new_size = (225, 250)
    img = img.resize(new_size)

    # Extraire l'extension du fichier d'origine
    original_extension = os.path.splitext(url)[-1]

    # Construire le nouveau chemin d'accès depuis la racine
    new_filepath = os.path.join(output_folder, f"{matricule}_{contact1}{original_extension}")

    # Sauvegarder l'image redimensionnée dans le dossier de sortie
    try:
        img.save(new_filepath)
    except:
        pass

    # Mettre à jour le chemin d'accès dans le DataFrame
    df.at[row.name, 'URL'] = new_filepath

# Appliquer la fonction à chaque ligne du DataFrame
df.apply(download_resize_and_rename_image, axis=1)

# Sauvegarder le DataFrame mis à jour dans le fichier Excel
df.to_excel(excel_file, index=False)

print("Téléchargement, redimensionnement et renommage terminés.")
