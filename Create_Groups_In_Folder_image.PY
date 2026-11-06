import os
import re
import shutil
from PIL import Image
import imagehash

# ==========================================================
# CONFIGURAÇÕES
# ==========================================================
INPUT_FOLDER = r"C:\Users\jean.silva\GRUPOS DE LAYOUTS\Lote 9\Modelos em  PNG"
OUTPUT_FOLDER = r"C:\Users\jean.silva\GRUPOS DE LAYOUTS\Lote 9\Grupos"

# Distância máxima entre hashes:
# 0  = idêntico Eu já instalei esse daí daquela vez, né? Era esse aqui? Eu não sei Você deixa aqui, já, no caso, como executar dentro já do terminal, reexecutar, como faz isso?
# 5  = muito parecido
# 8  = parecido (bom equilíbrio)
# 10 = mais flexível
HASH_THRESHOLD = 8

# True  = move os arquivos
# False = copia os arquivos
MOVE_FILES = False

SUPPORTED_EXTENSIONS = (
    ".png",
    ".jpg",
    ".jpeg",
    ".bmp",
    ".tif",
    ".tiff",
    ".gif",
    ".pdf"
)

# ==========================================================
# FUNÇÕES
# ==========================================================

def calculate_hash(image_path):
    """
    Calcula o hash perceptual da imagem.
    """
    try:
        with Image.open(image_path) as img:
            img = img.convert("RGB")
            return imagehash.phash(img)
    except Exception as e:
        print(f"Erro ao processar {image_path}: {e}")
        return None


def is_similar(hash1, hash2, threshold):
    """
    Verifica se dois hashes são semelhantes.
    """
    return (hash1 - hash2) <= threshold


def get_images(folder):
    """
    Retorna todas as imagens da pasta.
    """
    files = []

    for file_name in os.listdir(folder):
        full_path = os.path.join(folder, file_name)

        if (
            os.path.isfile(full_path)
            and file_name.lower().endswith(SUPPORTED_EXTENSIONS)
        ):
            files.append(full_path)

    return files


def group_images(image_files):
    """
    Agrupa imagens semelhantes.
    """
    print("Calculando hashes das imagens...")

    hashes = {}

    for path in image_files:
        print(f" - {os.path.basename(path)}")
        h = calculate_hash(path)

        if h is not None:
            hashes[path] = h

    groups = []

    print("\nAgrupando imagens...")

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


def extract_dimensions(filename):
    """
    Extrai as dimensões do nome do arquivo e normaliza para:
    100X100MM

    Exemplos aceitos:
    - 100x100mm
    - 100x100 mm
    - 100 mm x 100 mm
    - 100x 100 mm
    """
    filename = filename.lower()

    patterns = [
        r'(\d+)\s*mm\s*x\s*(\d+)\s*mm',
        r'(\d+)\s*x\s*(\d+)\s*mm',
        r'(\d+)\s*mm\s*x\s*(\d+)',
        r'(\d+)\s*x\s*(\d+)'
    ]

    for pattern in patterns:
        match = re.search(pattern, filename)
        if match:
            largura = match.group(1)
            altura = match.group(2)
            return f"{largura}X{altura}MM"

    return "SEM_DIMENSAO"


def save_groups(groups):
    """
    Move ou copia os arquivos para pastas separadas
    e gera um arquivo etiquetas.txt em cada grupo.

    Exemplo:
    GRUPO_001 - 100X100MM (10 DC)
    GRUPO_002 - 80X100MM (3 DC)
    """
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    for i, group in enumerate(groups, start=1):
        qtd_docs = len(group)

        # Pega dimensão do primeiro arquivo do grupo
        primeiro_arquivo = os.path.basename(group[0])
        dimensao = extract_dimensions(primeiro_arquivo)

        # Nome da pasta
        group_folder = os.path.join(
            OUTPUT_FOLDER,
            f"GRUPO_{i:03d} - {dimensao} ({qtd_docs} DC)"
        )

        os.makedirs(group_folder, exist_ok=True)

        nomes_arquivos = []

        for file_path in group:
            file_name = os.path.basename(file_path)
            destination = os.path.join(group_folder, file_name)

            nomes_arquivos.append(file_name)

            if MOVE_FILES:
                shutil.move(file_path, destination)
            else:
                shutil.copy2(file_path, destination)

        # Cria arquivo etiquetas.txt
        txt_path = os.path.join(group_folder, "etiquetas.txt")

        with open(txt_path, "w", encoding="utf-8") as f:
            for nome in sorted(nomes_arquivos):
                f.write(nome + "\n")


def print_summary(groups):
    """
    Exibe um resumo dos grupos encontrados.
    """
    print("\n================ RESUMO ================\n")

    for idx, group in enumerate(groups, start=1):
        primeiro_arquivo = os.path.basename(group[0])
        dimensao = extract_dimensions(primeiro_arquivo)

        print(
            f"GRUPO_{idx:03d} - {dimensao} ({len(group)} DC)"
        )

    print("\n=======================================\n")


# ==========================================================
# EXECUÇÃO
# ==========================================================

def main():
    if not os.path.exists(INPUT_FOLDER):
        print("Pasta de entrada não encontrada.")
        return

    image_files = get_images(INPUT_FOLDER)

    if not image_files:
        print("Nenhuma imagem encontrada na pasta.")
        return

    print(f"Encontradas {len(image_files)} imagens.\n")

    groups = group_images(image_files)

    print(f"\nForam identificados {len(groups)} grupos.")
    print_summary(groups)

    print("Salvando arquivos...")
    save_groups(groups)

    print("Processo concluído com sucesso!")
    print(f"Resultado salvo em:\n{OUTPUT_FOLDER}")


if __name__ == "__main__":
    main()