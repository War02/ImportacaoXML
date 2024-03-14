import pandas as pd

def insere_codigo(planilha_principal_path, planilha_geral_path):
    # Leitura dos arquivos do Excel, especificando o tipo de dado para CNPJ_CPF como string
    planilha_principal = pd.read_excel(planilha_principal_path, dtype={'CNPJ_CPF': str})
    planilha_geral = pd.read_excel(planilha_geral_path, dtype={'CNPJ_CPF': str})

    codigos_clientes = []
    for cnpj in planilha_principal['CNPJ_CPF']:
        linha_correspondente = planilha_geral[planilha_geral['CNPJ_CPF'] == cnpj]
        if not linha_correspondente.empty:
            codigo_cliente = linha_correspondente.iloc[0]['Inscricao']
            codigos_clientes.append(codigo_cliente)
        else:
            codigos_clientes.append(None)

    planilha_principal['Inscricao'] = codigos_clientes

    planilha_principal.to_excel(planilha_principal_path, index=False)

insere_codigo('\\\\192.168.10.51\\Dados\\Departamental\\TI\\!CIGAM\\Migracao\\Dados Usados para Migração\\Docs - Importacao - Cigam\\EMPRESAS_MODELOONDEPERMANECE O COD ANTERI - PRINCIPAL - Copia.xlsx', '\\\\192.168.10.51\\dados\\Departamental\\TI\\!CIGAM\\Migracao\\Dados Coletados para Migração\\Inscricao.xlsx')