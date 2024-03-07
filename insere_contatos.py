import pandas as pd

# Verifica se a empresa possui contato na planilha principal, se não possuir, busca em outra planilha e o adiciona
# na planilha principal

try:
    planilha_principal = pd.read_excel('C:\\Users\\User\\Documents\\Docs - Importacao - Cigam\\EMPRESAS_MODELOONDEPERMANECE O COD ANTERI - PRINCIPAL.xlsx')
    planilha_contatos = pd.read_excel('\\\\192.168.10.51\\Dados\\Departamental\\TI\\!CIGAM\\Migracao\\Dados Coletados para Migração\\contatos.xlsx')

    empresa_sem_telefone = planilha_principal[planilha_principal['Fone'].isnull()]

    for index, empresa in empresa_sem_telefone.iterrows():
        codigo_empresa = empresa['Empresa']
        contato_empresa = planilha_contatos[planilha_contatos['Codigo'] == codigo_empresa ]
        if not contato_empresa.empty:
            telefone = contato_empresa.iloc[0]['Fone']
            planilha_principal.at[index, 'Fone'] = telefone

    planilha_principal.to_excel('C:\\Users\\User\\Documents\\Docs - Importacao - Cigam\\EMPRESAS_MODELOONDEPERMANECE O COD ANTERI - PRINCIPAL.xlsx', index=False)

except Exception as e:
    print(f"Ocorreu um erro: {e}")

# planilha_contatos = pd.read_excel('\\\\192.168.10.51\\Dados\\Departamental\\TI\\!CIGAM\\Migracao\\Dados Coletados para Migração\\contatos.xlsx')
#
# planilha_contatos['Fone'] = planilha_contatos['Fone'].astype(str).str.lstrip('0')
#
# planilha_contatos.to_excel('\\\\192.168.10.51\\Dados\\Departamental\\TI\\!CIGAM\\Migracao\\Dados Coletados para Migração\\contatos.xlsx')