import os
import re
import unicodedata
from openpyxl import load_workbook

ARQUIVO_EXCEL = r"C:\Users\jean.silva\GRUPOS DE LAYOUTS\Lote 9\Lote 9_Dia_06.10  Used_For_Groups.xlsx"
PASTA_GRUPOS = r"C:\Users\jean.silva\GRUPOS DE LAYOUTS\Lote 9\Grupos"

ABA = "Planilha1"
COLUNA_RESULTADO = "Grupo Layout"


def normalizar(txt):
    if txt is None:
        return ""
    txt = str(txt).lower()
    txt = unicodedata.normalize("NFKD", txt).encode("ascii", "ignore").decode("ascii")
    txt = re.sub(r"[^a-z0-9]+", " ", txt)
    return re.sub(r"\s+", " ", txt).strip()


def extrair_dimensao(txt):
    txt = normalizar(txt)

    padroes = [
        r"(\d{2,3})\s*mm\s*x\s*(\d{2,3})\s*mm",
        r"(\d{2,3})\s*mm\s*x\s*(\d{2,3})",
        r"(\d{2,3})\s*mm\s*-\s*(\d{2,3})\s*mm",
        r"(\d{2,3})\s*mm\s*-\s*(\d{2,3})",
        r"(\d{2,3})\s*x\s*(\d{2,3})\s*mm",
        r"(\d{2,3})\s*x\s*(\d{2,3})",
    ]

    for p in padroes:
        m = re.search(p, txt)
        if m:
            return f"{m.group(1)}x{m.group(2)}"

    return ""


def limpar_codigo(valor):
    if valor is None:
        return ""

    txt = str(valor).strip()

    if txt.endswith(".0"):
        txt = txt[:-2]

    return txt


def achar_coluna(headers, possibilidades):
    for i, h in enumerate(headers, start=1):
        h_norm = normalizar(h)

        for p in possibilidades:
            if normalizar(p) in h_norm:
                return i

    return None


def carregar_grupos(pasta_grupos):
    grupos = []

    for nome_grupo in os.listdir(pasta_grupos):
        caminho_grupo = os.path.join(pasta_grupos, nome_grupo)

        if not os.path.isdir(caminho_grupo):
            continue

        textos = [nome_grupo]

        for raiz, _, arquivos in os.walk(caminho_grupo):
            for arq in arquivos:
                nome_sem_ext = os.path.splitext(arq)[0]
                textos.append(nome_sem_ext)

        texto_final = normalizar(" ".join(textos))
        dimensao = extrair_dimensao(texto_final)
        codigos = set(re.findall(r"\d{6,20}", texto_final))

        grupos.append({
            "grupo": nome_grupo,
            "texto": texto_final,
            "dimensao": dimensao,
            "codigos": codigos
        })

    return grupos


def criar_indices(grupos):
    indice_codigo = {}
    indice_dimensao = {}

    for g in grupos:
        for codigo in g["codigos"]:
            indice_codigo.setdefault(codigo, []).append(g)

        if g["dimensao"]:
            indice_dimensao.setdefault(g["dimensao"], []).append(g)

    return indice_codigo, indice_dimensao


def localizar_ultima_linha_util(ws, col_codigo, col_nome, col_desc):
    ultima_linha = 1

    for row in range(ws.max_row, 1, -1):
        valores = [
            ws.cell(row=row, column=col_codigo).value,
            ws.cell(row=row, column=col_nome).value if col_nome else None,
            ws.cell(row=row, column=col_desc).value if col_desc else None,
        ]

        if any(v not in [None, ""] for v in valores):
            ultima_linha = row
            break

    return ultima_linha


def melhor_grupo(row_texto, codigo, dimensao, grupos, indice_codigo, indice_dimensao):
    candidatos = []

    if codigo and codigo in indice_codigo:
        candidatos.extend(indice_codigo[codigo])

    if dimensao and dimensao in indice_dimensao:
        candidatos.extend(indice_dimensao[dimensao])

    if not candidatos:
        candidatos = grupos

    candidatos_unicos = {}
    for g in candidatos:
        candidatos_unicos[g["grupo"]] = g

    melhor = ""
    melhor_score = 0

    row_texto_norm = normalizar(row_texto)
    palavras_linha = set(row_texto_norm.split())

    for g in candidatos_unicos.values():
        score = 0

        if codigo and codigo in g["texto"]:
            score += 100

        if dimensao and dimensao == g["dimensao"]:
            score += 60
        elif dimensao and dimensao in g["texto"]:
            score += 40

        palavras_grupo = set(g["texto"].split())
        palavras_em_comum = palavras_linha.intersection(palavras_grupo)

        score += min(len(palavras_em_comum), 20)

        if score > melhor_score:
            melhor_score = score
            melhor = g["grupo"]

    if melhor_score >= 5:
        return melhor, melhor_score
       
    if melhor_score < 60:
        print(
        f"Sem match => "
        f"codigo={codigo} "
        f"dim={dimensao} "
        f"score={melhor_score}"
    )
    return "", melhor_score
 

def main():
    print("Abrindo planilha...")
    wb = load_workbook(ARQUIVO_EXCEL)
    ws = wb[ABA] if ABA in wb.sheetnames else wb.active

    headers = [cell.value for cell in ws[1]]

    col_nome = achar_coluna(headers, ["nome da etiqueta loftware", "nome etiqueta", "etiqueta"])
    col_desc = achar_coluna(headers, ["descricao da etiqueta", "descrição da etiqueta", "descricao", "descrição"])
    col_codigo = achar_coluna(headers, ["l$_codigo", "ls_codigo", "l codigo", "ls codigo", "codigo"])

    if not col_codigo:
        raise Exception("Não encontrei a coluna L$_codigo / ls_codigo / código.")

    if not col_nome and not col_desc:
        raise Exception("Não encontrei coluna de nome/descrição da etiqueta.")

    col_resultado = len(headers) + 1
    ws.cell(row=1, column=col_resultado).value = COLUNA_RESULTADO

    col_score = col_resultado + 1
    ws.cell(row=1, column=col_score).value = "Score Match"

    print("Carregando grupos...")
    grupos = carregar_grupos(PASTA_GRUPOS)

    print(f"Grupos carregados: {len(grupos)}")

    indice_codigo, indice_dimensao = criar_indices(grupos)

    ultima_linha = localizar_ultima_linha_util(ws, col_codigo, col_nome, col_desc)
    total_linhas = ultima_linha - 1

    encontrados = 0
    nao_encontrados = 0

    print(f"Última linha útil encontrada: {ultima_linha}")
    print(f"Iniciando processamento de {total_linhas} linhas...")

    for row in range(2, ultima_linha + 1):
        if row % 50 == 0:
            print(f"Processando linha {row} de {ultima_linha}...")

        nome = ws.cell(row=row, column=col_nome).value if col_nome else ""
        desc = ws.cell(row=row, column=col_desc).value if col_desc else ""
        codigo = limpar_codigo(ws.cell(row=row, column=col_codigo).value)

        texto_linha = f"{nome} {desc} {codigo}"
        dimensao = extrair_dimensao(texto_linha)

        grupo, score = melhor_grupo(
            texto_linha,
            codigo,
            dimensao,
            grupos,
            indice_codigo,
            indice_dimensao
        )

        ws.cell(row=row, column=col_resultado).value = grupo
        ws.cell(row=row, column=col_score).value = score

        if grupo:
            encontrados += 1
        else:
            nao_encontrados += 1

    saida = ARQUIVO_EXCEL.replace(".xlsx", "_COM_GRUPOS.xlsx")

    print("Salvando arquivo...")
    wb.save(saida)

    print("Finalizado!")
    print(f"Encontrados: {encontrados}")
    print(f"Não encontrados: {nao_encontrados}")
    print(f"Arquivo gerado: {saida}")


if __name__ == "__main__":
    main()