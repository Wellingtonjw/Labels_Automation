# -*- coding: utf-8 -*-
"""
Automacao NiceLabel - versao Python
Convertido de NICELABEL_V16_TODAS_LINHAS_CONTINUA_ERROS_CORRIGIDO.ps1

Instalacao:
    pip install pywin32 pythonnet

Execucao:
    Abra este arquivo no VS Code e clique em "Run Python File".
"""

import gc
import os
import re
import sys
import unicodedata
from datetime import datetime
from pathlib import Path

import pythoncom
import win32com.client
# ============================================================
# Esse Código foi convertido de um script PowerShell para Python, mantendo a mesma lógica e funcionalidade.
# Esse código é destinado a automatizar a geração de imagens de etiquetas usando o NiceLabel SDK.
# Funcionaliades a que foram adicionadas Por mim JW
# 1 - Verificação de linhas vazias consecultivas no excel, se 2 linhas vazias então o codigo para.
# 2 - Adicionado a opção de escolher a linha inicial para gerar as imagens, caso queira pular algumas linhas do excel 
# 2.1 - Adicionado a opção de definir linha final de processamento. 
# 3 - Contador de linhas uteis encontradas na planilha excel. Essa linha aparece logo no inicio do processamento.
# 3.1 - Contador de colunas uteis encontradas na planilha excel. Isso diminui muito o tempo de processamento.
# ============================================================

# ============================================================
# CONFIGURACAO
# ============================================================
#Caminho do arquivo Excel que contem os dados das etiquetas
EXCEL_PATH = Path(
    r"C:\Users\jean.silva\Sheets for PDF Generator\LOTE 11\Lote 11_rebranding_2_Banco de dados_05NOV.xlsx")

#caminho da pasta que contem os arquivos de etiquetas do NiceLabel
LABELS_DIR = Path(r"Z:\Etiquetas Clientes BOMI")

#Caminho da pasta de saida dos arquivos PDF gerados
OUTPUT_DIR = Path(r"C:\Users\jean.silva\OutPDFLabelsAutomator\LabelsAtomation")

#caminho da pasta de saida dos arquivos PDF gerados
TEMP_DIR = OUTPUT_DIR / "LOT 11_05NOV"

#caminho da pasta de saida dos arquivos PDF gerados
GRAPHICS_ROOT = TEMP_DIR / "GRAPHICS_TODAS_LINHAS"

#caminho da pasta que contem a SDK do NiceLabel, normalmente em C:\Program Files\NiceLabel\NiceLabel 10\bin.net
SDK_DIR = Path(r"C:\Program Files\NiceLabel\NiceLabel 10\bin.net")
SDK_DLL = SDK_DIR / "SDK.NET.Interface.dll"

#Nome da coluna no Excel que contem o nome da etiqueta do NiceLabel - 
#O Sistema percorrerá essa coluna, os nomes da coluna devem ser iguais aos nomes dos aquivos de etiquetas do NiceLabel (sem a extensão .nlbl)
LABEL_COLUMN_NAME = "Nome da Etiqueta Loftware"
DESCRIPTION_COLUMN_NAME = "Descricao da Etiqueta"

#Variavel para definição de quantidade de linhas vazias antes de parar o processamento, caso existam 2 linhas vazias consecutivas o código para.
MAX_EMPTY_ROWS = 2


def normalize_header(value) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    decomposed = unicodedata.normalize("NFD", text)
    without_accents = "".join(
        ch for ch in decomposed if unicodedata.category(ch) != "Mn"
    )
    normalized = unicodedata.normalize("NFC", without_accents)
    normalized = re.sub(r"\s+", " ", normalized)
    return normalized.lower()


def is_blank(value) -> bool:
    return value is None or str(value).strip() == ""


def excel_text(sheet, row: int, column: int) -> str:
    value = sheet.Cells.Item(row, column).Text
    return "" if value is None else str(value)


def clean_excel_value(value: str) -> str:
    value = re.sub(r"\r\n|\r|\n", " ", value)
    value = re.sub(r"[ \t]{2,}", " ", value)
    return value.strip()


def get_inner_exception_message(exc: Exception) -> str:
    try:
        inner = getattr(exc, "InnerException", None)
        if inner is not None:
            return str(inner.Message)
    except Exception:
        pass
    return ""


def write_utf8_lines(path: Path, lines):
    path.write_text("\n".join(lines), encoding="utf-8")


def load_nicelabel_sdk():
    """Carrega a mesma SDK.NET.Interface.dll utilizada pelo PowerShell."""
    if not SDK_DLL.exists():
        raise FileNotFoundError(f"SDK do NiceLabel nao encontrado: {SDK_DLL}")

    if hasattr(os, "add_dll_directory"):
        os.add_dll_directory(str(SDK_DIR))

    if str(SDK_DIR) not in sys.path:
        sys.path.insert(0, str(SDK_DIR))

    # Para assemblies .NET Framework do NiceLabel no Windows.
    try:
        from pythonnet import load
        load("netfx")
    except RuntimeError:
        # Runtime ja inicializado.
        pass

    import clr
    import System

    System.Reflection.Assembly.LoadFrom(str(SDK_DLL))
    from NiceLabel.SDK import PrintEngineFactory, PrintToGraphicsSettings

    return System, PrintEngineFactory, PrintToGraphicsSettings


def configure_graphics_destination(System, graphics_settings, graphics_dir: Path) -> int:
    """Replica a reflexao do PowerShell para localizar a propriedade de destino."""
    flags = (
        System.Reflection.BindingFlags.Instance
        | System.Reflection.BindingFlags.Public
        | System.Reflection.BindingFlags.NonPublic
    )

    settings_type = graphics_settings.GetType()
    set_count = 0

    for prop in settings_type.GetProperties(flags):
        if not prop.CanWrite:
            continue

        name = str(prop.Name).lower()
        matches = any(
            token in name
            for token in ("folder", "directory", "destination", "output", "path", "file")
        )
        if not matches:
            continue
        if prop.PropertyType != System.String:
            continue

        try:
            prop.SetValue(graphics_settings, str(graphics_dir), None)
            set_count += 1
        except Exception:
            pass

    return set_count


def generate_missing_labels_report(missing_labels):
    if not missing_labels:
        print("\nNenhuma etiqueta faltante.")
        return

    report_path = TEMP_DIR / "ETIQUETAS_NAO_ENCONTRADAS.txt"
    lines = [
        "ETIQUETAS NAO ENCONTRADAS",
        "Gerado em: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "",
        f"Total: {len(missing_labels)}",
        "",
    ]

    print("\nLISTA DE ETIQUETAS NAO ENCONTRADAS:\n")
    for item in missing_labels:
        line = (
            f"Linha Excel {item['ExcelRow']} | Etiqueta: {item['LabelName']} | "
            f"Descricao: {item['Description']} | Caminho: {item['ExpectedPath']}"
        )
        lines.append(line)
        print(f"- Linha {item['ExcelRow']}")
        print(f"  Etiqueta: {item['LabelName']}")
        print(f"  Descricao: {item['Description']}")
        print(f"  Caminho: {item['ExpectedPath']}")
        print()

    write_utf8_lines(report_path, lines)
    print("Relatorio salvo em:")
    print(report_path)


def generate_print_errors_report(print_errors):
    if not print_errors:
        print("\nNenhum erro de renderizacao.")
        return

    report_path = TEMP_DIR / "ERROS_DE_RENDERIZACAO.txt"
    lines = [
        "ERROS DE RENDERIZACAO / PRINT TO GRAPHICS",
        "Gerado em: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "",
        f"Total: {len(print_errors)}",
        "",
    ]

    print("\nERROS DE RENDERIZACAO:\n")
    for item in print_errors:
        line = (
            f"Linha Excel {item['ExcelRow']} | Etiqueta: {item['LabelName']} | "
            f"Descricao: {item['Description']} | Erro: {item['ErrorMessage']} | "
            f"Detalhe: {item['InnerMessage']}"
        )
        lines.append(line)
        print(f"- Linha {item['ExcelRow']}")
        print(f"  Etiqueta: {item['LabelName']}")
        print(f"  Descricao: {item['Description']}")
        print(f"  Erro: {item['ErrorMessage']}")
        if item["InnerMessage"].strip():
            print(f"  Detalhe: {item['InnerMessage']}")
        print()

    write_utf8_lines(report_path, lines)
    print("Relatorio de erros salvo em:")
    print(report_path)


def main():
    excel_app = None
    workbook = None
    sheet = None
    print_engine = None

    pythoncom.CoInitialize()

    print()
    print("========================================")
    print(" NICELABEL - PYTHON - TODAS AS LINHAS")
    print("========================================")
    print()
    print("Saida configurada:")
    print(OUTPUT_DIR)
    print()

    for path in (OUTPUT_DIR, TEMP_DIR, GRAPHICS_ROOT):
        path.mkdir(parents=True, exist_ok=True)

    if not EXCEL_PATH.exists():
        raise FileNotFoundError(f"Arquivo Excel nao encontrado: {EXCEL_PATH}")

    try:
        # ----------------------------------------------------
        # Excel
        # ----------------------------------------------------
        print("Abrindo Excel...")
        excel_app = win32com.client.DispatchEx("Excel.Application")
        excel_app.Visible = False
        excel_app.DisplayAlerts = False

        workbook = excel_app.Workbooks.Open(str(EXCEL_PATH))
        sheet = workbook.Worksheets.Item(1)

        #----------------------------------------------------
        # Contador de colunas
        #----------------------------------------------------
        header_row = sheet.Rows(1)

        last_header_cell = header_row.Find(
            What="*",
            After=header_row.Cells(1, 1),
            LookAt=2,
            LookIn=-4163,
            SearchOrder=2,
            SearchDirection=2,
            MatchCase=False
        )

        if last_header_cell is None:
            raise RuntimeError("Nenhum cabecalho foi encontrado na linha 1.")

        col_count = last_header_cell.Column

        print(f"Planilha encontrada: {col_count} colunas com cabecalho detectadas")
        print()

        # ----------------------------------------------------
        # Cabecalhos
        # ----------------------------------------------------
        headers = {}
        for column in range(1, col_count + 1):
            header = excel_text(sheet, 1, column)
            if not is_blank(header):
                headers[header.strip()] = column

        # ----------------------------------------------------
        # Identifica coluna da etiqueta
        # ----------------------------------------------------
        label_column = None
        for header, column in headers.items():
            if normalize_header(header) == normalize_header(LABEL_COLUMN_NAME):
                label_column = column
                break
        if label_column is None:
            raise RuntimeError("Coluna Nome da Etiqueta Loftware nao encontrada.")

        # ----------------------------------------------------
        # Identifica coluna da descricao
        # ----------------------------------------------------        
        description_column = None

        for header, column in headers.items():
            if normalize_header(header) == normalize_header(DESCRIPTION_COLUMN_NAME):
                description_column = column
                break
        if description_column is None:
            raise RuntimeError("Coluna Descricao da Etiqueta nao encontrada.")


        # ----------------------------------------------------
        # Identifica ultima linha preechida 
        # ---------------------------------------------------- 
        def find_last_data_row(sheet, column_number):
            column_range = sheet.Columns(column_number)

            last_cell = column_range.Find(
                What="*",
                After=column_range.Cells(1, 1),
                LookAt=2,
                LookIn=-4163,
                SearchOrder=1,
                SearchDirection=2,
                MatchCase=False
            )

            if last_cell is None:
                return 1

            return last_cell.Row


        last_label_row = find_last_data_row(sheet, label_column)
        last_description_row = find_last_data_row(sheet, description_column)

        last_used_row = max(
            last_label_row,
            last_description_row
)
        # ----------------------------------------------------
        # Contador de linhas úteis
        # ---------------------------------------------------- 
        row_count = last_used_row

        print()
        print("========================================")
        print(f"TOTAL DE LINHAS ÚTEIS ENCONTRADAS: {row_count}")
        print("========================================")
        print()
        
        # ----------------------------------------------------
        # Linhas para iniciar a geração de imagens, caso queira pular algumas linhas do excel
        # ---------------------------------------------------- 
        #Linha inicial
        while True:
            try:
                start_row = int(
                    input(
                        "A partir de qual linha você deseja começar a gerar as imagens? Começa a partir de 2 (linha 1 é o cabeçalho): "
                    )
                )

                if start_row < 2 or start_row > row_count:
                    print(
                        f"Digite uma linha entre 2 e {row_count}."
                    )
                    continue

                break

            except ValueError:
                print("Digite apenas um número inteiro.")
        # Linha final
        while True:
            try:
                end_row = int(
                    input(
                        f"Até qual linha você deseja gerar as imagens? "
                        f"(máximo {row_count}): "
                    )
                )

                if end_row < start_row:
                    print(
                        f"A linha final deve ser maior ou igual à "
                        f"linha inicial ({start_row})."
                    )
                    continue

                if end_row > row_count:
                    print(
                        f"Digite uma linha entre {start_row} e {row_count}."
                    )
                    continue

                break

            except ValueError:
                print("Digite apenas um número inteiro.")
        # ----------------------------------------------------
        # Listas de controle
        # ---------------------------------------------------- 
        missing_labels = []
        print_errors = []
    
        # ----------------------------------------------------
        # NiceLabel SDK
        # ----------------------------------------------------
        print("Carregando NiceLabel SDK...")
        System, PrintEngineFactory, PrintToGraphicsSettings = load_nicelabel_sdk()

        print("Inicializando Print Engine...")
        PrintEngineFactory.SDKFilesPath = str(SDK_DIR)
        print_engine = PrintEngineFactory.PrintEngine
        print_engine.Initialize()
        print("Print Engine inicializado.")
        print()

        #----------------------------------------------------
        #Variaveis contadoras linhas vazias - Isso para validação de linhas vazias consecutivas
        #----------------------------------------------------
        processed_count = 0
        consecutive_empty_rows = 0
   

        # ----------------------------------------------------
        # Todas as linhas de dados
        # ----------------------------------------------------
        #for que realiza a interação entre linha inicial e linha final (Sem esse for a logica de linha inicial e final não funcionara)
        for excel_row in range(start_row, end_row + 1):
            print()
            print("========================================")
            print(f" PROCESSANDO EXCEL LINHA {excel_row}")
            print("========================================")

            label_name = excel_text(sheet, excel_row, label_column)
            pdf_description = excel_text(sheet, excel_row, description_column)

            # Linha vazia apenas quando os dois campos estao vazios.
            #Logica para contar linhas vazias, se existirem 2 duas linhas vazias o código para. 
            if is_blank(label_name) and is_blank(pdf_description):
                consecutive_empty_rows += 1
                print()
                print(
                    f"Linha {excel_row} vazia. Sequencia atual: "
                    f"{consecutive_empty_rows}"
                )

                if consecutive_empty_rows >= MAX_EMPTY_ROWS:
                    print()
                    print("========================================")
                    print("2 LINHAS VAZIAS CONSECUTIVAS DETECTADAS")
                    print("PROCESSAMENTO ENCERRADO")
                    print("========================================")
                    break
                continue
            else:
                consecutive_empty_rows = 0

            if is_blank(label_name):
                print(f"Linha {excel_row} sem nome de etiqueta. Pulando.")
                continue

            print("\nEtiqueta:")
            print(label_name)
            print("\nDescricao:")
            print(pdf_description)

            graphics_dir = GRAPHICS_ROOT / f"LINHA_{excel_row:04d}"
            graphics_dir.mkdir(parents=True, exist_ok=True)

            label_path = LABELS_DIR / f"{label_name}.nlbl"

            if not label_path.exists():
                missing_labels.append(
                    {
                        "ExcelRow": excel_row,
                        "LabelName": label_name,
                        "Description": pdf_description,
                        "ExpectedPath": str(label_path),
                    }
                )
                print()
                print(f"ETIQUETA NAO ENCONTRADA - LINHA {excel_row}")
                print(f"Nome da etiqueta: {label_name}")
                print(f"Descricao: {pdf_description}")
                print(f"Caminho procurado: {label_path}")
                print("Linha sera ignorada e o processamento continuara.")
                continue

            # ------------------------------------------------
            # Abre etiqueta
            # ------------------------------------------------
            print("\nAbrindo etiqueta:")
            print(label_path)

            label = print_engine.OpenLabel(str(label_path))
            if label is None:
                raise RuntimeError("OpenLabel retornou nulo.")

            print("Etiqueta aberta com sucesso.")
            print(f"Variaveis encontradas na etiqueta: {label.Variables.Count}")

            label_variable_names = set()
            for variable in label.Variables:
                try:
                    label_variable_names.add(str(variable.Name))
                except Exception:
                    pass

            # ------------------------------------------------
            # Variaveis da linha
            # ------------------------------------------------
            variable_lines = []
            filled = 0

            for header, column in headers.items():
                normalized_header = normalize_header(header)

                if normalized_header in (
                    normalize_header(LABEL_COLUMN_NAME),
                    normalize_header(DESCRIPTION_COLUMN_NAME),
                ):
                    continue

                if header not in label_variable_names:
                    continue

                value = excel_text(sheet, excel_row, column)
                if is_blank(value):
                    continue

                value = clean_excel_value(value)
                escaped_value = value.replace('"', '""')
                variable_lines.append(f'{header}="{escaped_value}"')
                filled += 1

            variables_list = os.linesep.join(variable_lines)

            print()
            print(f"Variaveis preparadas: {filled}")
            print("Primeiras variaveis enviadas:")
            for line in variable_lines[:8]:
                print("  " + line)

            variable_file = TEMP_DIR / f"teste_variaveis_linha_{excel_row:04d}_v14.txt"
            variable_file.write_text(variables_list, encoding="utf-8")

            print("\nAplicando valores com OverloadVariableValues()...")
            label.OverloadVariableValues(variables_list, True)
            print("Valores aplicados.")

            # ------------------------------------------------
            # PrintToGraphics
            # ------------------------------------------------
            graphics_settings = PrintToGraphicsSettings()
            graphics_settings.PrintAll = True
            graphics_settings.PrintToFiles = True

            set_count = configure_graphics_destination(
                System, graphics_settings, graphics_dir
            )
            if set_count == 0:
                raise RuntimeError(
                    "Nao foi encontrada uma propriedade de destino de arquivos "
                    "em PrintToGraphicsSettings."
                )

            print(f"\nRenderizando linha {excel_row}...")
            try:
                graphics_result = label.PrintToGraphics(graphics_settings)
                print(f"PrintToGraphics concluido para a linha {excel_row}.")
                processed_count += 1
            except Exception as exc:
                error_message = str(exc)
                inner_message = get_inner_exception_message(exc)
                print_errors.append(
                    {
                        "ExcelRow": excel_row,
                        "LabelName": label_name,
                        "Description": pdf_description,
                        "ErrorMessage": error_message,
                        "InnerMessage": inner_message,
                    }
                )

                print()
                print(f"ERRO AO RENDERIZAR A LINHA {excel_row}")
                print(f"Etiqueta: {label_name}")
                print(f"Descricao: {pdf_description}")
                print(f"Erro: {error_message}")
                if inner_message:
                    print(f"Detalhe: {inner_message}")
                print("Linha sera ignorada e o processamento continuara.")

            label = None

        # ----------------------------------------------------
        # Resumo
        # ----------------------------------------------------
        print()
        print("========================================")
        print(" RESUMO DO PROCESSAMENTO")
        print("========================================")
        print()

        total_rows = row_count - 1
        missing_count = len(missing_labels)

        print(f"Total de linhas de dados: {total_rows}")
        print(f"Linhas processadas com sucesso: {processed_count}")
        print(f"Etiquetas nao encontradas: {missing_count}")
        print(f"Erros de renderizacao: {len(print_errors)}")

        generate_missing_labels_report(missing_labels)
        generate_print_errors_report(print_errors)

        print()
        print("========================================")
        print(" TESTE DE TODAS AS LINHAS FINALIZADO")
        print("========================================")
        print()
        print("Imagens por linha:")
        print(GRAPHICS_ROOT)

    finally:
        try:
            if workbook is not None:
                workbook.Close(False)
        except Exception:
            pass

        try:
            if excel_app is not None:
                excel_app.Quit()
        except Exception:
            pass

        try:
            if print_engine is not None:
                print_engine.Shutdown()
                print("Print Engine encerrado.")
        except Exception:
            pass

        sheet = None
        workbook = None
        excel_app = None
        gc.collect()

        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print()
        print("========================================")
        print("ERRO FATAL")
        print("========================================")
        print(str(exc))

        inner = get_inner_exception_message(exc)
        if inner:
            print("Detalhe:")
            print(inner)
    finally:
        print()
        input("Teste finalizado. Pressione ENTER para fechar...")