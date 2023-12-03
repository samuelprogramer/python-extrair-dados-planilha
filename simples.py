import openpyxl

def extrair_dados_planilha(nome_arquivo, nome_planilha):
    # Carregar o arquivo Excel
    workbook = openpyxl.load_workbook(nome_arquivo)

    # Selecionar a planilha pelo nome
    planilha = workbook[nome_planilha]

    # Iterar sobre as linhas da planilha
    for linha in planilha.iter_rows(min_row=2, values_only=True):
        # Supondo que a primeira coluna contém dados importantes
        dado = linha[0]
        print(f"Dado extraído: {dado}")

    # Fechar o arquivo Excel
    workbook.close()

# Substitua 'seu_arquivo.xlsx' pelo nome do seu arquivo Excel
# Substitua 'SuaPlanilha' pelo nome da sua planilha
extrair_dados_planilha('planilhaQuestoes.xlsx', 'Planilha1')
