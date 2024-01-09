#BIBLIOTECAS
from datetime import datetime
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename

# SELECIONAR ARQUIVO EXCEL - TK
Tk().withdraw()  # ESCONDE JANELA PRINCIPAL
file_path = askopenfilename(filetypes=[("Excel files", ".xlsx;.xls")])

# CARREGA O ARQUIVO EXCEL - PANDAS
try:
    df = pd.read_excel(file_path)
except pd.errors.ParserError:
    print("Erro ao ler o arquivo. Verifique se é um arquivo Excel válido.")
    exit()

# INSERE NOME DO FORNECEDOR
fornecedor_procurado = input("Digite o nome do fornecedor a ser procurado no histórico: ")

# FILTRA REGISTROS QUE CONTENHAM NOME DO FORNECEDOR SELECIONADO
registros_fornecedor = df[df['Histórico'].str.contains(fornecedor_procurado, case=False)]

if registros_fornecedor.empty:
    print(f"Não foram encontrados registros para o fornecedor '{fornecedor_procurado}' no histórico.")
else:
    print(f"\nRegistros que contêm o fornecedor '{fornecedor_procurado}' no histórico:")
    print(registros_fornecedor)

    # SOLICITA PERIODO DETERMINADO
    data_inicio = input("Digite a data de início do período (no formato dd/mm/aaaa): ")
    data_fim = input("Digite a data de fim do período (no formato dd/mm/aaaa): ")

    try:
        data_inicio_formatada = datetime.strptime(data_inicio, "%d/%m/%Y")
        data_fim_formatada = datetime.strptime(data_fim, "%d/%m/%Y")

        # FILTRA OS REGISTROS NO PERIODO DETERMINADO
        registros_periodo = registros_fornecedor[
            (registros_fornecedor.iloc[:, 0] >= data_inicio_formatada) &
            (registros_fornecedor.iloc[:, 0] <= data_fim_formatada)
            ]

        if registros_periodo.empty:
            print(f"Não foram encontrados registros para o período de '{data_inicio}' a '{data_fim}'.")
        else:
            print(f"\nRegistros para o período de '{data_inicio}' a '{data_fim}':")
            print(registros_periodo)

            # SOMA DEBITO E CREDITO DENTRO DO PERIODO SELECIONADO
            debito_total = registros_periodo['Débito'].sum()
            credito_total = registros_periodo['Crédito'].sum()
            print(
                f"\nTotal de débito para '{fornecedor_procurado}' no período de '{data_inicio}' a '{data_fim}': {debito_total}")
            print(
                f"Total de crédito para '{fornecedor_procurado}' no período de '{data_inicio}' a '{data_fim}': {credito_total}")

    except ValueError:
        print("Datas inseridas no formato incorreto. Use o formato dd/mm/aaaa.")