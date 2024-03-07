import pandas as pd

# Identifica as empresas que são inscritas e adiciona como 1 na planilha principal

todos_clientes = pd.read_excel('C:\\Users\\User\\Documents\\Docs - Importacao - Cigam\\EMPRESAS_MODELOONDEPERMANECE O COD ANTERI - PRINCIPAL.xlsx')
clientes_com_inscricao = pd.read_excel('\\\\192.168.10.51\\Dados\\Departamental\\TI\\!CIGAM\\Migracao\\Dados Coletados para Migração\\INSCRICAO E TIPO DE PESSOA.xlsx')

for codigo_cliente in todos_clientes['Empresa']:
    if codigo_cliente in clientes_com_inscricao['CODIGO'].values:
        todos_clientes.loc[todos_clientes['Empresa'] == codigo_cliente, 'Inscrito'] = 1

todos_clientes.to_excel('C:\\Users\\User\\Documents\\Docs - Importacao - Cigam\\EMPRESAS_MODELOONDEPERMANECE O COD ANTERI - PRINCIPAL.xlsx', index=False)
