
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment

# ==========================================
# CONFIGURAÇÕES
# ==========================================

BASE_DIR = Path(__file__).resolve().parent
ENTRADA = BASE_DIR / "entrada"
SAIDA = BASE_DIR / "saida"

ARQUIVO_ANTERIOR = ENTRADA / "anterior.xlsx"
ARQUIVO_ATUAL = ENTRADA / "atual.xlsx"

# Nome da coluna que identifica cada registro
COLUNA_ID = "ID"

ENTRADA.mkdir(exist_ok=True)
SAIDA.mkdir(exist_ok=True)


def ler_planilha(caminho):
    """Lê uma planilha e retorna cabeçalhos e registros."""

    wb = load_workbook(caminho, data_only=True, read_only=True)
    ws = wb.active

    linhas = ws.iter_rows(values_only=True)
    cabecalhos = next(linhas, None)

    if not cabecalhos:
        wb.close()
        raise ValueError(f"A planilha {caminho.name} está vazia.")

    cabecalhos = [
        str(valor).strip() if valor is not None else ""
        for valor in cabecalhos
    ]

    if any(not coluna for coluna in cabecalhos):
        wb.close()
        raise ValueError(
            f"{caminho.name}: existe coluna sem nome no cabeçalho."
        )

    if len(set(cabecalhos)) != len(cabecalhos):
        wb.close()
        raise ValueError(
            f"{caminho.name}: existem nomes de coluna repetidos."
        )

    if COLUNA_ID not in cabecalhos:
        wb.close()
        raise ValueError(
            f"{caminho.name}: não existe a coluna '{COLUNA_ID}'."
        )

    registros = {}

    for numero_linha, valores in enumerate(linhas, start=2):
        valores = list(valores)

        if all(v is None or str(v).strip() == "" for v in valores):
            continue

        valores += [None] * (len(cabecalhos) - len(valores))
        registro = dict(zip(cabecalhos, valores))
        identificador = registro.get(COLUNA_ID)

        if identificador is None or str(identificador).strip() == "":
            wb.close()
            raise ValueError(
                f"{caminho.name}: ID vazio na linha {numero_linha}."
            )

        chave = str(identificador).strip().casefold()

        if chave in registros:
            wb.close()
            raise ValueError(
                f"{caminho.name}: ID duplicado '{identificador}' "
                f"na linha {numero_linha}."
            )

        registros[chave] = registro

    wb.close()
    return cabecalhos, registros


def normalizar(valor):
    """Compara valores de maneira consistente."""

    if valor is None:
        return ""

    if isinstance(valor, str):
        return valor.strip()

    return valor


def comparar(anteriores, atuais, colunas):
    """Identifica registros novos, removidos e alterados."""

    novos = []
    removidos = []
    alterados = []

    ids_anteriores = set(anteriores)
    ids_atuais = set(atuais)

    for identificador in sorted(ids_atuais - ids_anteriores):
        novos.append(atuais[identificador])

    for identificador in sorted(ids_anteriores - ids_atuais):
        removidos.append(anteriores[identificador])

    colunas_comuns = [
        coluna for coluna in colunas
        if coluna != COLUNA_ID
    ]

    for identificador in sorted(ids_anteriores & ids_atuais):
        antigo = anteriores[identificador]
        atual = atuais[identificador]

        for coluna in colunas_comuns:
            valor_antigo = normalizar(antigo.get(coluna))
            valor_atual = normalizar(atual.get(coluna))

            if valor_antigo != valor_atual:
                alterados.append([
                    identificador,
                    coluna,
                    valor_antigo,
                    valor_atual,
                ])

    return novos, removidos, alterados


def gerar_relatorio(novos, removidos, alterados, colunas):
    """Cria um Excel com o resumo e os detalhes da comparação."""

    wb = Workbook()

    resumo = wb.active
    resumo.title = "Resumo"

    resumo.append(["RELATÓRIO DE COMPARAÇÃO"])
    resumo.append([
        "Data da execução",
        datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    ])
    resumo.append(["Registros novos", len(novos)])
    resumo.append(["Registros removidos", len(removidos)])
    resumo.append(["Campos alterados", len(alterados)])

    ws_novos = wb.create_sheet("Novos")
    ws_novos.append(colunas)
    for registro in novos:
        ws_novos.append([registro.get(c) for c in colunas])

    ws_removidos = wb.create_sheet("Removidos")
    ws_removidos.append(colunas)
    for registro in removidos:
        ws_removidos.append([registro.get(c) for c in colunas])

    ws_alterados = wb.create_sheet("Alterações")
    ws_alterados.append([
        "ID", "Coluna", "Valor anterior", "Valor atual"
    ])
    for alteracao in alterados:
        ws_alterados.append(alteracao)

    for planilha in wb.worksheets:
        planilha.freeze_panes = "A2"
        planilha.auto_filter.ref = planilha.dimensions

        for celula in planilha[1]:
            celula.font = Font(bold=True, color="FFFFFF")
            celula.fill = PatternFill(
                fill_type="solid",
                fgColor="1F4E78"
            )
            celula.alignment = Alignment(
                vertical="center",
                wrap_text=True
            )

        for coluna in planilha.columns:
            largura = max(
                len(str(c.value or "")) for c in coluna
            )
            planilha.column_dimensions[
                coluna[0].column_letter
            ].width = min(largura + 3, 45)

    # Reforça o título da primeira aba
    resumo["A1"].font = Font(bold=True, size=15)

    nome = f"comparacao_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    destino = SAIDA / nome
    wb.save(destino)

    return destino


def main():
    try:
        if not ARQUIVO_ANTERIOR.exists():
            raise FileNotFoundError(
                f"Coloque o arquivo anterior em: {ARQUIVO_ANTERIOR}"
            )

        if not ARQUIVO_ATUAL.exists():
            raise FileNotFoundError(
                f"Coloque o arquivo atual em: {ARQUIVO_ATUAL}"
            )

        colunas_antigas, anteriores = ler_planilha(ARQUIVO_ANTERIOR)
        colunas_atuais, atuais = ler_planilha(ARQUIVO_ATUAL)

        if colunas_antigas != colunas_atuais:
            raise ValueError(
                "As planilhas possuem cabeçalhos diferentes "
                "ou em ordens diferentes. Padronize os cabeçalhos "
                "antes de comparar."
            )

        novos, removidos, alterados = comparar(
            anteriores, atuais, colunas_atuais
        )

        destino = gerar_relatorio(
            novos, removidos, alterados, colunas_atuais
        )

        print("\nCOMPARAÇÃO CONCLUÍDA")
        print(f"Registros novos: {len(novos)}")
        print(f"Registros removidos: {len(removidos)}")
        print(f"Campos alterados: {len(alterados)}")
        print(f"\nRelatório gerado: {destino}")

    except Exception as erro:
        print(f"\nNão foi possível comparar: {erro}")


if __name__ == "__main__":
    main()
