import os
import shutil
import tempfile
from PIL import Image
import imagehash
import fitz  # PyMuPDF


INPUT_FOLDER = r"C:\Users\jhonatan.pereira\Downloads\etiquetas"
OUTPUT_FOLDER = r"C:\Users\jhonatan.pereira\Downloads\etiquetas\Nova pasta"

HASH_THRESHOLD = 8
MOVE_FILES = False

SUPPORTED_EXTENSIONS = (
    ".png", ".jpg", ".jpeg", ".bmp",
    ".tif", ".tiff", ".gif", ".pdf"
)


def pdf_to_temp_image(pdf_path):
    try:
        doc = fitz.open(pdf_path)
        page = doc[0]

        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)

        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        temp_file.close()

        pix.save(temp_file.name)
        doc.close()

        return temp_file.name

    except Exception as e:
        print(f"Erro ao converter PDF para análise {pdf_path}: {e}")
        return None


def convert_pdf_preview(pdf_path, output_folder):
    try:
        doc = fitz.open(pdf_path)
        page = doc[0]

        pix = page.get_pixmap(matrix=fitz.Matrix(3, 3), alpha=False)

        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        output_image = os.path.join(output_folder, f"{base_name}_preview.png")

        pix.save(output_image)
        doc.close()

        print(f"   Preview gerado: {os.path.basename(output_image)}")

    except Exception as e:
        print(f"Erro ao converter preview do PDF {pdf_path}: {e}")


def calculate_hash(file_path):
    temp_image = None

    try:
        if file_path.lower().endswith(".pdf"):
            temp_image = pdf_to_temp_image(file_path)

            if temp_image is None:
                return None

            image_path = temp_image
        else:
            image_path = file_path

        with Image.open(image_path) as img:
            img = img.convert("RGB")
            return imagehash.phash(img)

    except Exception as e:
        print(f"Erro ao processar {file_path}: {e}")
        return None

    finally:
        if temp_image and os.path.exists(temp_image):
            os.remove(temp_image)


def is_similar(hash1, hash2, threshold):
    return (hash1 - hash2) <= threshold


def get_files(folder):
    files = []

    for file_name in os.listdir(folder):
        full_path = os.path.join(folder, file_name)

        if (
            os.path.isfile(full_path)
            and file_name.lower().endswith(SUPPORTED_EXTENSIONS)
        ):
            files.append(full_path)

    return files


def group_files(files):
    print("Calculando hashes dos arquivos...")

    hashes = {}

    for path in files:
        print(f" - {os.path.basename(path)}")
        h = calculate_hash(path)

        if h is not None:
            hashes[path] = h
        else:
            print(f"   NÃO FOI POSSÍVEL PROCESSAR: {os.path.basename(path)}")

    groups = []

    print("\nAgrupando arquivos...")

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
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    for i, group in enumerate(groups, start=1):
        group_folder = os.path.join(
            OUTPUT_FOLDER,
            f"GRUPO_{i:03d} ({len(group)} DC)"
        )

        os.makedirs(group_folder, exist_ok=True)

        nomes_arquivos = []

        print(f"\nSalvando {os.path.basename(group_folder)}...")

        for file_path in group:
            file_name = os.path.basename(file_path)
            destination = os.path.join(group_folder, file_name)

            nomes_arquivos.append(file_name)

            if MOVE_FILES:
                shutil.move(file_path, destination)
            else:
                shutil.copy2(file_path, destination)

            if destination.lower().endswith(".pdf"):
                convert_pdf_preview(destination, group_folder)

        txt_path = os.path.join(group_folder, "etiquetas.txt")

        with open(txt_path, "w", encoding="utf-8") as f:
            for nome in sorted(nomes_arquivos):
                f.write(nome + "\n")


def print_summary(groups):
    print("\n================ RESUMO ================\n")

    for idx, group in enumerate(groups, start=1):
        print(f"GRUPO_{idx:03d}: {len(group)} arquivo(s)")

        for file_path in group:
            print(f"  - {os.path.basename(file_path)}")

    print("\n=======================================\n")


def main():
    if not os.path.exists(INPUT_FOLDER):
        print("Pasta de entrada não encontrada.")
        return

    files = get_files(INPUT_FOLDER)

    if not files:
        print("Nenhum arquivo encontrado na pasta.")
        return

    print(f"Encontrados {len(files)} arquivos.\n")

    groups = group_files(files)

    print(f"\nForam identificados {len(groups)} grupos.")
    print_summary(groups)

    print("Salvando arquivos...")
    save_groups(groups)

    print("\nProcesso concluído com sucesso!")
    print(f"Resultado salvo em:\n{OUTPUT_FOLDER}")


if __name__ == "__main__":
    main()