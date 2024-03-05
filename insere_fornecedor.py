from openpyxl import load_workbook

def copiar_dados(planilha_origem_path, planilha_destino_path, mapeamento_fornecedores, linha_inicial):
    # Carrega a planilha de origem
    wb_origem = load_workbook(filename=planilha_origem_path)
    ws_origem = wb_origem.active

    # Carrega a planilha de destino
    wb_destino = load_workbook(filename=planilha_destino_path)
    ws_destino = wb_destino.active

    # Itera sobre as linhas da planilha de origem
    for row_origem in ws_origem.iter_rows(min_row=linha_inicial, values_only=True):
        # Encontra a próxima linha vazia na planilha de destino
        linha_destino = ws_destino.max_row + 1

        # Copia os dados da planilha de origem para a planilha de destino, apenas nas colunas correspondentes
        for col_origem, col_destino in mapeamento_fornecedores.items():
            # Encontra o índice da coluna na planilha de origem
            index_col_origem = list(mapeamento_fornecedores.keys()).index(col_origem)

            # Obtém o valor da célula na coluna de origem
            valor_origem = row_origem[index_col_origem]

            # Define o valor na célula na coluna de destino
            ws_destino.cell(row=linha_destino, column=col_destino, value=valor_origem)

    # Salva as alterações na planilha de destino
    wb_destino.save(planilha_destino_path)

    print("Dados copiados com sucesso.")

# Mapeamento de colunas entre a planilha de origem e a planilha de destino
mapeamento_fornecedores = {
    'CODIGO': 1,
    'RAZAO': 2,
    'ENDERECO': 6,
    'CIDADE': 8,
    'ESTADO': 9,
    'CEP': 10,
    'CNPJ': 11
}

# Chama a função para copiar os dados da planilha de origem para a planilha de destino
copiar_dados('C:\\Users\\User\\Documents\\Docs - Importacao - Cigam\\FORNECEDORES.xlsx', 'C:\\Users\\User\\Documents\\Docs - Importacao - Cigam\\EMPRESAS_MODELOONDEPERMANECE O COD ANTERI(1) - Principal.xlsx', mapeamento_fornecedores, 2)
