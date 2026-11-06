import os
import re
import shutil
from pathlib import Path
 
import openpyxl
from PIL import Image, ImageOps
 
 
# ============================================================
# CONFIGURAÇÕES
# ============================================================
 
# Pasta onde estão as LINHA_XXXX
PASTA_ORIGEM = Path(
    r"C:\Users\jean.silva\OutPDFLabelsAutomator\LabelsAtomation\LOT 11_05NOV\GRAPHICS_TODAS_LINHAS"
)
 
# Pasta única onde serão colocados os PDFs
PASTA_DESTINO = Path(
    r"C:\Users\jean.silva\OutPDFLabelsAutomator\LabelsAtomation\PDF_LOT 11_05NOV"
)
  
# Planilha
ARQUIVO_EXCEL = Path(
    r"C:\Users\jean.silva\Sheets for PDF Generator\LOTE 11\Lote 11_rebranding_2_Banco de dados_05NOV.xlsx"
)

# Aba da planilha
ABA = "Planilha1"
 
# Nome da coluna que contém o nome final
COLUNA_DESCRICAO = "Descrição da Etiqueta"

 
# Extensões consideradas imagens
EXTENSOES_IMAGEM = {
    ".png",
    ".jpg",
    ".jpeg",
    ".bmp",
    ".tif",
    ".tiff",
    ".webp",
}
 
# PDFs
EXTENSAO_PDF = ".pdf"
 
# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================
 
def limpar_nome_arquivo(nome):
    """
    Remove caracteres que não podem existir em nomes de arquivos
    no Windows.
    """
 
    nome = str(nome).strip()
 
    # Caracteres proibidos no Windows:
    # < > : " / \ | ? *
    nome = re.sub(r'[<>:"/\\|?*]', '_', nome)
 
    # Remove pontos/espaços no final
    nome = nome.rstrip(" .")
 
    return nome
 
 
def encontrar_coluna(ws, nome_coluna):
    """
    Localiza a coluna pelo nome do cabeçalho.
    """
 
    for coluna in range(1, ws.max_column + 1):
        valor = ws.cell(1, coluna).value
 
        if valor is not None and str(valor).strip() == nome_coluna:
            return coluna
 
    raise ValueError(
        f'Coluna "{nome_coluna}" não encontrada na planilha.'
    )
 
 
def obter_numero_linha(nome_pasta):
    """
    LINHA_0002 -> 2
    LINHA_0015 -> 15
    """
 
    match = re.fullmatch(r"LINHA_(\d+)", nome_pasta.upper())
 
    if not match:
        return None
 
    return int(match.group(1))
 
 
def encontrar_arquivo_mais_recente(pasta):
    """
    Retorna o arquivo mais recentemente modificado dentro da pasta.
    Considera apenas imagens e PDFs.
    """
 
    arquivos = []
 
    for arquivo in pasta.iterdir():
 
        if not arquivo.is_file():
            continue
 
        extensao = arquivo.suffix.lower()
 
        if extensao in EXTENSOES_IMAGEM or extensao == EXTENSAO_PDF:
            arquivos.append(arquivo)
 
    if not arquivos:
        return None
 
    return max(
        arquivos,
        key=lambda arquivo: arquivo.stat().st_mtime
    )
 
 
def converter_imagem_para_pdf(imagem_origem, pdf_destino):
    """
    Converte uma imagem para PDF.
    """
 
    with Image.open(imagem_origem) as img:
 
        # Corrige orientação de imagens que possuem EXIF
        img = ImageOps.exif_transpose(img)
 
        # PDF precisa estar em RGB
        if img.mode in ("RGBA", "LA"):
            fundo = Image.new("RGB", img.size, "white")
 
            if img.mode == "RGBA":
                fundo.paste(img, mask=img.getchannel("A"))
            else:
                fundo.paste(img, mask=img.getchannel("A"))
 
            img = fundo
 
        elif img.mode != "RGB":
            img = img.convert("RGB")
 
        img.save(
            pdf_destino,
            "PDF",
            resolution=100.0
        )
 
 
def mover_pdf(pdf_origem, pdf_destino):
    """
    Move o PDF para a pasta final.
    """
 
    shutil.move(
        str(pdf_origem),
        str(pdf_destino)
    )
 
 
# ============================================================
# PROCESSAMENTO PRINCIPAL
# ============================================================
 
def main():
    
 
    print("=" * 70)
    print(" ORGANIZADOR DE GRAPHICS -> PDF")
    print("=" * 70)
    print()
 
    # --------------------------------------------------------
    # Verificações
    # --------------------------------------------------------
 
    if not PASTA_ORIGEM.exists():
        print(f"[ERRO] Pasta de origem não encontrada:")
        print(PASTA_ORIGEM)
        return
 
    if not ARQUIVO_EXCEL.exists():
        print(f"[ERRO] Planilha não encontrada:")
        print(ARQUIVO_EXCEL)
        return
 
    # Cria pasta de destino
    PASTA_DESTINO.mkdir(
        parents=True,
        exist_ok=True
    )
 
    # --------------------------------------------------------
    # Abre Excel
    # --------------------------------------------------------
 
    print("[INFO] Abrindo planilha...")
 
    wb = openpyxl.load_workbook(
        ARQUIVO_EXCEL,
        read_only=True,
        data_only=True
    )
 
    if ABA not in wb.sheetnames:
        print(f'[ERRO] Aba "{ABA}" não encontrada.')
        print("Abas disponíveis:", wb.sheetnames)
        return
 
    ws = wb[ABA]
 
    coluna_descricao = encontrar_coluna(
        ws,
        COLUNA_DESCRICAO
    )
 
    print(f"[OK] Coluna encontrada: {COLUNA_DESCRICAO}")
    print()
 
    # --------------------------------------------------------
    # Relatórios
    # --------------------------------------------------------
 
    linhas_nao_encontradas = []
    pastas_sem_arquivo = []
    erros = []
    processados = []
    ignorados = []
 
    # --------------------------------------------------------
    # Localiza as pastas LINHA_XXXX
    # --------------------------------------------------------
 
    pastas = sorted(
        [
            pasta
            for pasta in PASTA_ORIGEM.iterdir()
            if pasta.is_dir()
            and obter_numero_linha(pasta.name) is not None
        ],
        key=lambda p: obter_numero_linha(p.name)
    )
 
    print(f"[INFO] Foram encontradas {len(pastas)} pastas.")
    print()
 
    # --------------------------------------------------------
    # Processa cada pasta
    # --------------------------------------------------------
 
    for pasta in pastas:
 
        nome_pasta = pasta.name
        numero_linha = obter_numero_linha(nome_pasta)
 
        print("-" * 70)
        print(f"Processando: {nome_pasta}")
 
        # ----------------------------------------------------
        # Busca linha correspondente na planilha
        #
        # LINHA_0002 -> linha 2
        # LINHA_0003 -> linha 3
        # ----------------------------------------------------
 
        if numero_linha > ws.max_row:
 
            print(
                f"[FALTANTE] Não existe linha {numero_linha} "
                f"na planilha."
            )
 
            linhas_nao_encontradas.append(
                nome_pasta
            )
 
            continue
 
        descricao = ws.cell(
            numero_linha,
            coluna_descricao
        ).value
 
        if descricao is None or str(descricao).strip() == "":
 
            print(
                f"[FALTANTE] A linha {numero_linha} "
                f"não possui Descrição da Etiqueta."
            )
 
            linhas_nao_encontradas.append(
                f"{nome_pasta} - Descrição vazia"
            )
 
            continue
 
        descricao = limpar_nome_arquivo(
            descricao
        )
 
        print(f"[INFO] Descrição: {descricao}")
 
        # ----------------------------------------------------
        # Procura o arquivo mais recente
        # ----------------------------------------------------
 
        arquivo = encontrar_arquivo_mais_recente(
            pasta
        )
 
        if arquivo is None:
 
            print(
                "[SEM ARQUIVO] Nenhuma imagem ou PDF encontrada."
            )
 
            pastas_sem_arquivo.append(
                nome_pasta
            )
 
            continue
 
        print(f"[INFO] Arquivo mais recente: {arquivo.name}")
 
        # ----------------------------------------------------
        # Define destino
        # ----------------------------------------------------
 
        destino = PASTA_DESTINO / f"{descricao}.pdf"
 
        # ----------------------------------------------------
        # Evita sobrescrever
        # ----------------------------------------------------
 
        if destino.exists():
 
            print(
                f"[IGNORADO] Arquivo já existe:"
            )
            print(f"           {destino}")
 
            ignorados.append(
                f"{nome_pasta} -> {destino.name}"
            )
 
            continue
 
        # ----------------------------------------------------
        # PDF
        # ----------------------------------------------------
 
        if arquivo.suffix.lower() == ".pdf":
 
            try:
 
                mover_pdf(
                    arquivo,
                    destino
                )
 
                print("[OK] PDF movido.")
                print(f"     {destino.name}")
 
                processados.append(
                    f"{nome_pasta} -> {destino.name} [PDF]"
                )
 
            except Exception as e:
 
                print(f"[ERRO] Falha ao mover PDF: {e}")
 
                erros.append(
                    f"{nome_pasta}: {e}"
                )
 
        # ----------------------------------------------------
        # IMAGEM
        # ----------------------------------------------------
 
        elif arquivo.suffix.lower() in EXTENSOES_IMAGEM:
 
            try:
 
                converter_imagem_para_pdf(
                    arquivo,
                    destino
                )
 
                print("[OK] Imagem convertida para PDF.")
                print(f"     {destino.name}")
 
                # Só remove a imagem original depois que
                # o PDF foi criado com sucesso.
                arquivo.unlink()
 
                print("[OK] Imagem original removida.")
 
                processados.append(
                    f"{nome_pasta} -> {destino.name} [IMAGEM]"
                )
 
            except Exception as e:
 
                print(
                    f"[ERRO] Falha ao converter imagem: {e}"
                )
 
                # Se o PDF chegou a ser criado parcialmente,
                # remove para evitar deixar arquivo corrompido.
                if destino.exists():
                    try:
                        destino.unlink()
                    except:
                        pass
 
                erros.append(
                    f"{nome_pasta}: {e}"
                )
 
    # --------------------------------------------------------
    # GERA RELATÓRIOS
    # --------------------------------------------------------
 
    print()
    print("=" * 70)
    print(" FINALIZADO")
    print("=" * 70)
 
    # --------------------------------------------------------
    # Linhas não encontradas
    # --------------------------------------------------------
 
    arquivo_faltantes = (
        PASTA_DESTINO /
        "LINHAS_NAO_ENCONTRADAS.txt"
    )
 
    with open(
        arquivo_faltantes,
        "w",
        encoding="utf-8"
    ) as f:
 
        f.write(
            "LINHAS / PASTAS SEM CORRESPONDÊNCIA\n"
        )
        f.write("=" * 50 + "\n\n")
 
        if linhas_nao_encontradas:
 
            for item in linhas_nao_encontradas:
                f.write(f"{item}\n")
 
        else:
            f.write(
                "Nenhuma linha sem correspondência.\n"
            )
 
    # --------------------------------------------------------
    # Sem arquivo
    # --------------------------------------------------------
 
    arquivo_sem_arquivo = (
        PASTA_DESTINO /
        "PASTAS_SEM_ARQUIVO.txt"
    )
 
    with open(
        arquivo_sem_arquivo,
        "w",
        encoding="utf-8"
    ) as f:
 
        f.write(
            "PASTAS SEM IMAGEM OU PDF\n"
        )
        f.write("=" * 50 + "\n\n")
 
        if pastas_sem_arquivo:
 
            for item in pastas_sem_arquivo:
                f.write(f"{item}\n")
 
        else:
            f.write(
                "Nenhuma pasta sem arquivo.\n"
            )
 
    # --------------------------------------------------------
    # Erros
    # --------------------------------------------------------
 
    arquivo_erros = (
        PASTA_DESTINO /
        "ERROS.txt"
    )
 
    with open(
        arquivo_erros,
        "w",
        encoding="utf-8"
    ) as f:
 
        f.write(
            "ERROS DURANTE O PROCESSAMENTO\n"
        )
        f.write("=" * 50 + "\n\n")
 
        if erros:
 
            for item in erros:
                f.write(f"{item}\n")
 
        else:
            f.write(
                "Nenhum erro encontrado.\n"
            )
 
    # --------------------------------------------------------
    # Resumo
    # --------------------------------------------------------
 
    print()
    print(f"Processados:              {len(processados)}")
    print(f"Sem correspondência:      {len(linhas_nao_encontradas)}")
    print(f"Sem arquivo:              {len(pastas_sem_arquivo)}")
    print(f"Ignorados:                {len(ignorados)}")
    print(f"Erros:                    {len(erros)}")
    print()
    print("Relatórios:")
    print(f"  {arquivo_faltantes}")
    print(f"  {arquivo_sem_arquivo}")
    print(f"  {arquivo_erros}")
    print()
    print("Processamento concluído.")
 
 
# ============================================================
# EXECUÇÃO
# ============================================================
 
if __name__ == "__main__":
    main()