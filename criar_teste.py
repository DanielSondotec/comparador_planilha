
from pathlib import Path
from openpyxl import Workbook

entrada = Path(__file__).resolve().parent / "entrada"
entrada.mkdir(exist_ok=True)


def criar(nome, registros):
    wb = Workbook()
    ws = wb.active
    ws.append(["ID", "Nome", "Departamento", "Status"])

    for registro in registros:
        ws.append(registro)

    wb.save(entrada / nome)


criar("anterior.xlsx", [
    [1, "Ana", "Financeiro", "Ativo"],
    [2, "Bruno", "TI", "Ativo"],
    [3, "Carla", "RH", "Pendente"],
])

criar("atual.xlsx", [
    [1, "Ana", "Financeiro", "Ativo"],
    [2, "Bruno", "TI", "Inativo"],
    [4, "Daniel", "Compras", "Ativo"],
])

print("Planilhas fictícias criadas na pasta entrada.")
