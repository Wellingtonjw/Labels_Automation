import os
import pymupdf


# ==========================================================
# CONFIGURAÇÕES
# ==========================================================

# Pasta onde estão os PDFs
INPUT_FOLDER = r"C:\Users\jean.silva\GRUPOS DE LAYOUTS\LOTE 12"

# Pasta onde serão salvas as imagens convertidas
OUTPUT_FOLDER = r"C:\Users\jean.silva\GRUPOS DE LAYOUTS\LOTE 12"

# Qualidade da conversão
# 150 = mais rápido / arquivo menor
# 200 = recomendado
# 300 = maior qualidade / arquivos maiores
DPI = 200


# ==========================================================
# CONVERSÃO
# ==========================================================

def convert_pdf_to_png(pdf_path, output_folder):

    try:

        # Abre o PDF
        pdf = pymupdf.open(pdf_path)

        # Verifica se possui páginas
        if len(pdf) == 0:
            print(f"[ERRO] PDF sem páginas: {os.path.basename(pdf_path)}")
            pdf.close()
            return False

        # ==================================================
        # PEGA A PRIMEIRA PÁGINA
        # ==================================================

        page = pdf[0]

        # Renderiza a página
        pix = page.get_pixmap(
            dpi=DPI,
            alpha=False
        )

        # Nome original do arquivo
        original_name = os.path.splitext(
            os.path.basename(pdf_path)
        )[0]

        # Nome da imagem
        output_path = os.path.join(
            output_folder,
            original_name + ".png"
        )

        # Salva PNG
        pix.save(output_path)

        pdf.close()

        print(
            f"[OK] {os.path.basename(pdf_path)}"
            f" -> {original_name}.png"
        )

        return True

    except Exception as e:

        print(
            f"[ERRO] {os.path.basename(pdf_path)}: {e}"
        )

        return False


# ==========================================================
# MAIN
# ==========================================================

def main():

    print("\n==============================================")
    print(" CONVERSOR DE PDF PARA PNG")
    print("==============================================\n")

    # Verifica pasta de entrada
    if not os.path.exists(INPUT_FOLDER):

        print("ERRO: Pasta de entrada não encontrada.")
        print(INPUT_FOLDER)

        return

    # Cria pasta de saída
    os.makedirs(
        OUTPUT_FOLDER,
        exist_ok=True
    )

    # Busca PDFs
    pdf_files = []

    for file_name in os.listdir(INPUT_FOLDER):

        full_path = os.path.join(
            INPUT_FOLDER,
            file_name
        )

        if (
            os.path.isfile(full_path)
            and file_name.lower().endswith(".pdf")
        ):

            pdf_files.append(full_path)

    # Nenhum PDF
    if not pdf_files:

        print("Nenhum PDF encontrado.")

        return

    print(
        f"Encontrados {len(pdf_files)} PDFs.\n"
    )

    # ======================================================
    # CONVERTE
    # ======================================================

    sucesso = 0
    erros = 0

    for pdf_path in pdf_files:

        resultado = convert_pdf_to_png(
            pdf_path,
            OUTPUT_FOLDER
        )

        if resultado:

            sucesso += 1

        else:

            erros += 1

    # ======================================================
    # RESUMO
    # ======================================================

    print("\n==============================================")
    print(" PROCESSO FINALIZADO")
    print("==============================================")

    print(f"PDFs encontrados : {len(pdf_files)}")
    print(f"Convertidos      : {sucesso}")
    print(f"Com erro         : {erros}")

    print(
        f"\nImagens salvas em:\n{OUTPUT_FOLDER}"
    )

    print("\n==============================================\n")


# ==========================================================
# EXECUÇÃO
# ==========================================================

if __name__ == "__main__":
    main()