import pandas as pd

#Irá consultar a planilha geral, através do CNPJ, para implementar o código do cliente à planilha principal
def insere_codigo(planilha_principal_path, planilha_geral_path):

    planilha_principal = pd.read_excel(planilha_principal_path)
    planilha_geral = pd.read_excel(planilha_geral_path)

    codigos_clientes = []
    for cnpj in planilha_principal['CNPJ_CPF']:
        linha_correspondente = planilha_geral[planilha_geral['CNPJ'] == cnpj]
        if not linha_correspondente.empty:
            codigo_cliente = linha_correspondente.iloc[0]['CODIGO']
            codigos_clientes.append(codigo_cliente)
        else:
            codigos_clientes.append(None)

    planilha_principal['Empresa'] = codigos_clientes

    planilha_principal.to_excel(planilha_principal_path, index=False)

insere_codigo('C:\\Users\\User\\Documents\\Docs - Importacao - Cigam\\EMPRESAS_MODELOONDEPERMANECE O COD ANTERI(1) - Principal.xlsx', '\\\\192.168.10.51\\Dados\\Departamental\\TI\\!CIGAM\\Implantação\\CAD GERAL.xlsx')