import os
import re
import time
import unicodedata
import pyautogui
import pyperclip
from openpyxl import load_workbook
from pywinauto import Desktop, keyboard

ARQUIVO_EXCEL = r"C:\Users\jhonatan.pereira\Downloads\etiquetas\Nova pasta\29.05_Banco de dados_padrão_genesis_Solventum_REV2 (4)_COM_GRUPOS.xlsx"
ABA = "2° banco de dados"

PASTA_PDFS = r"C:\Users\jhonatan.pereira\Downloads\etiquetas\pdfs_nicelabel"

COLUNA_DESCRICAO_ETIQUETA = "Descrição da Etiqueta"

LINHA_INICIAL = 2
LINHA_FINAL = 300

VARIAVEIS_NICELABEL = [
    "L$_codigo",
    "L$_lote",
    "L$_validade",
    "EP$_instrucoes_uso_1",
    "EP$_imp_distr_nome",
    "EP$_fabricante_nome",
    "EP$_detentor_registro_nome",
    "EP$_SAC",
]

# Distância entre o nome da variável e o campo valor.
# Se clicar longe ou perto demais, ajuste para 200, 300, 400, 500.
DISTANCIA_CAMPO_VALOR = 350


def normalizar(txt):
    if txt is None:
        return ""

    txt = str(txt).strip()
    txt = txt.replace("EP$_", "EPS_").replace("L$_", "LS_")
    txt = txt.replace("ep$_", "eps_").replace("l$_", "ls_")
    txt = txt.lower()
    txt = unicodedata.normalize("NFKD", txt).encode("ascii", "ignore").decode("ascii")
    txt = re.sub(r"[^a-z0-9]+", "_", txt)
    txt = re.sub(r"_+", "_", txt)

    return txt.strip("_")


def limpar_valor(valor):
    if valor is None:
        return ""

    txt = str(valor).strip()

    if txt.endswith(".0"):
        txt = txt[:-2]

    txt = txt.replace("\r", " ").replace("\n", " ")
    return txt


def nome_arquivo_seguro(txt):
    txt = limpar_valor(txt)
    txt = unicodedata.normalize("NFKD", txt).encode("ascii", "ignore").decode("ascii")
    txt = re.sub(r'[\\/:*?"<>|]+', "_", txt)
    txt = re.sub(r"\s+", "_", txt).strip("_")

    return txt[:140] if txt else "sem_descricao"


def montar_indice_colunas(headers):
    indice = {}

    for i, h in enumerate(headers, start=1):
        if h is not None:
            indice[normalizar(h)] = i

    return indice


def achar_coluna(headers, nome):
    alvo = normalizar(nome)

    for i, h in enumerate(headers, start=1):
        if normalizar(h) == alvo:
            return i

    return None


def pegar_valor(ws, row, indice_colunas, variavel):
    col = indice_colunas.get(normalizar(variavel))

    if not col:
        return ""

    return limpar_valor(ws.cell(row=row, column=col).value)


def conectar_nicelabel():
    janelas = Desktop(backend="uia").windows(title_re=r".*NiceLabel.*", top_level_only=True)

    if not janelas:
        raise Exception("Nenhuma janela do NiceLabel encontrada.")

    print("\nJanelas NiceLabel encontradas:")

    for i, j in enumerate(janelas):
        try:
            print(f"{i}: {j.window_text()}")
        except Exception:
            print(f"{i}: sem título")

    escolha = int(input("Digite o número da janela correta do NiceLabel: "))

    janela = janelas[escolha]
    janela.set_focus()
    time.sleep(1)

    print(f"Janela selecionada: {janela.window_text()}")

    return janela


def colar_texto(valor):
    pyperclip.copy(valor)
    pyautogui.hotkey("ctrl", "a")
    time.sleep(0.08)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.15)


def preencher_variavel(janela, variavel, valor):
    controles = janela.descendants()
    alvo = None
    variavel_norm = normalizar(variavel)

    for c in controles:
        try:
            texto = c.window_text().strip()
            texto_norm = normalizar(texto)

            if texto_norm == variavel_norm:
                alvo = c
                break
        except Exception:
            pass

    if not alvo:
        print(f"Variável não encontrada na tela: {variavel}")
        return False

    rect = alvo.rectangle()

    y = int((rect.top + rect.bottom) / 2)
    x = rect.right + DISTANCIA_CAMPO_VALOR

    print(f"Clicando em {variavel} | X={x}, Y={y}")

    pyautogui.click(x, y)
    time.sleep(0.2)

    colar_texto(valor)

    return True


def salvar_pdf(nome_pdf):
    caminho_pdf = os.path.join(PASTA_PDFS, nome_pdf)

    if os.path.exists(caminho_pdf):
        os.remove(caminho_pdf)

    keyboard.send_keys("^p")
    time.sleep(2)

    keyboard.send_keys("{ENTER}")
    time.sleep(2)

    pyperclip.copy(caminho_pdf)
    keyboard.send_keys("^v")
    time.sleep(0.5)
    keyboard.send_keys("{ENTER}")

    print(f"Aguardando PDF: {caminho_pdf}")

    for _ in range(30):
        if os.path.exists(caminho_pdf):
            print("PDF criado com sucesso.")
            return True

        time.sleep(1)

    print("PDF não apareceu na pasta.")
    return False


def main():
    os.makedirs(PASTA_PDFS, exist_ok=True)

    print("Abrindo planilha...")
    wb = load_workbook(ARQUIVO_EXCEL, data_only=True)
    ws = wb[ABA] if ABA in wb.sheetnames else wb.active

    headers = [cell.value for cell in ws[1]]
    indice_colunas = montar_indice_colunas(headers)

    col_descricao = achar_coluna(headers, COLUNA_DESCRICAO_ETIQUETA)

    if not col_descricao:
        raise Exception(f"Não encontrei a coluna: {COLUNA_DESCRICAO_ETIQUETA}")

    print("\nValidando variáveis na planilha:")

    for v in VARIAVEIS_NICELABEL:
        status = "OK" if normalizar(v) in indice_colunas else "NÃO ACHOU NA PLANILHA"
        print(f"{v}: {status}")

    print("\nDeixe o NiceLabel aberto na tela onde aparecem as variáveis.")
    input("Quando estiver pronto, aperte ENTER...")

    janela = conectar_nicelabel()

    gerados = 0
    falhas = 0

    for row in range(LINHA_INICIAL, LINHA_FINAL + 1):
        descricao = ws.cell(row=row, column=col_descricao).value

        if descricao in [None, ""]:
            continue

        print(f"\nLinha {row}: {descricao}")

        preenchidas = 0

        for variavel in VARIAVEIS_NICELABEL:
            valor = pegar_valor(ws, row, indice_colunas, variavel)

            if valor == "":
                print(f"{variavel}: vazio na planilha. Pulando.")
                continue

            ok = preencher_variavel(janela, variavel, valor)

            if ok:
                print(f"Preencheu {variavel}: {valor}")
                preenchidas += 1

            time.sleep(0.2)

        if preenchidas == 0:
            print("Nenhuma variável preenchida. Pulando impressão.")
            falhas += 1
            continue

        nome_pdf = nome_arquivo_seguro(descricao) + ".pdf"

        print(f"Gerando PDF: {nome_pdf}")

        if salvar_pdf(nome_pdf):
            gerados += 1
        else:
            falhas += 1

        time.sleep(1.5)

    print("\nFinalizado.")
    print(f"Gerados: {gerados}")
    print(f"Falhas: {falhas}")
    print(f"Pasta PDFs: {PASTA_PDFS}")


if __name__ == "__main__":
    main()