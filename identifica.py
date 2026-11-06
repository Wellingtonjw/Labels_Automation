"""
Agrupador de imagens por similaridade visual.

O script:
1. Lê todas as imagens de uma pasta.
2. Calcula um hash perceptual (pHash) para cada imagem.
3. Compara as imagens e agrupa as que possuem layout semelhante.
4. Cria subpastas GRUPO_001, GRUPO_002, etc.
5. Move ou copia as imagens para essas pastas.

Ideal para:
- Etiquetas
- PDFs convertidos em imagem
- Logos
- Layouts semelhantes

Instalação:
pip install pillow imagehash

Se quiser apenas copiar (sem remover da pasta original),
altere MOVE_FILES = False.
"""

import os
import shutil
from PIL import Image
import imagehash

# ==========================================
# CONFIGURAÇÕES
# ==========================================
INPUT_FOLDER = r"C:\Users\jhonatan.pereira\OneDrive - Bomi Group\Documentos\Materiais Solventum\Etiquetas parte 3\Etiquetas parte 3"
OUTPUT_FOLDER = r"C:\Users\jhonatan.pereira\OneDrive - Bomi Group\Documentos\Materiais Solventum\Etiquetas parte 3\Etiquetas parte 3\Nova pasta"

# Distância máxima entre hashes.
# Quanto menor, mais rigoroso.
# 0 = idêntico
# 5 = muito parecido
# 10 = parecido
HASH_THRESHOLD = 8

# True = move arquivos
# False = copia arquivos
MOVE_FILES = False

SUPPORTED_EXTENSIONS = (".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff")

# ==========================================
# FUNÇÕES
# ==========================================

def calculate_hash(image_path):
    """
    Calcula o perceptual hash da imagem.
    """
    try:
        with Image.open(image_path) as img:
            # converte para RGB para padronizar
            img = img.convert("RGB")
            return imagehash.phash(img)
    except Exception as e:
        print(f"Erro ao processar {image_path}: {e}")
        return None


def is_similar(hash1, hash2, threshold):
    """
    Compara dois hashes.
    """
    return (hash1 - hash2) <= threshold


def get_images(folder):
    """
    Lista imagens da pasta.
    """
    files = []
    for file in os.listdir(folder):
        if file.lower().endswith(SUPPORTED_EXTENSIONS):
            files.append(os.path.join(folder, file))
    return files


def group_images(image_files):
    """
    Agrupa imagens semelhantes.
    """
    hashes = {}
    for path in image_files:
        h = calculate_hash(path)
        if h is not None:
            hashes[path] = h

    groups = []

    for path, h in hashes.items():
        added = False

        for group in groups:
            representative = group[0]
            rep_hash = hashes[representative]

            if is_similar(h, rep_hash, HASH_THRESHOLD):
                group.append(path)
                added = True
                break

        if not added:
            groups.append([path])

    return groups


def save_groups(groups):
    """
    Move ou copia os arquivos para pastas separadas.
    """
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    for i, group in enumerate(groups, start=1):
        group_folder = os.path.join(OUTPUT_FOLDER, f"GRUPO_{i:03d}")
        os.makedirs(group_folder, exist_ok=True)

        for file_path in group:
            destination = os.path.join(
                group_folder,
                os.path.basename(file_path)
            )

            if MOVE_FILES:
                shutil.move(file_path, destination)
            else:
                shutil.copy2(file_path, destination)


# ==========================================
# EXECUÇÃO
# ==========================================

def main():
    image_files = get_images(INPUT_FOLDER)

    if not image_files:
        print("Nenhuma imagem encontrada.")
        return

    print(f"Encontradas {len(image_files)} imagens.")
    print("Calculando similaridade...")

    groups = group_images(image_files)

    print(f"Foram identificados {len(groups)} grupos.")

    for idx, group in enumerate(groups, start=1):
        print(f"Grupo {idx:03d}: {len(group)} arquivo(s)")

    print("Salvando grupos...")
    save_groups(groups)

    print("Processo concluído com sucesso.")


if __name__ == "__main__":
    main()