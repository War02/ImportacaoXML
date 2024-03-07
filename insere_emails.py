import pandas as pd

# Busca na planilha principal, pelo código da empresa e, o compara com outra planilha para dectar o e-mail do responsavel
# salvando as informações na planilha de contatos
try:
    planilha_empresas = pd.read_excel('C:\\Users\\User\\Documents\\Docs - Importacao - Cigam\\EMPRESAS_MODELOONDEPERMANECE O COD ANTERI - TESTE.xlsx')
    planilha_emails = pd.read_excel('\\\\192.168.10.51\\Dados\\Departamental\\TI\\!CIGAM\\Migracao\\Dados Coletados para Migração\\EMAILCLIENTES.xlsx')
    #planilha_importacao = pd.read_csv('C:\\Users\\User\\Documents\\Docs - Importacao - Cigam\\CONTATO-EMAIL_MODELOONDEPERMANECE O COD ANTERI.csv')

    nome_coluna_planilha_empresas = 'Empresa'
    nome_coluna_planilha_emails = 'Codigo'

    planilha_mesclada = pd.merge(planilha_empresas, planilha_emails, left_on=nome_coluna_planilha_empresas, right_on=nome_coluna_planilha_emails, how='inner')

    colunas_desejadas = ['Codigo', 'email', 'email-nf']
    planilha_final = planilha_mesclada[colunas_desejadas]

    planilha_final.to_excel('C:\\Users\\User\\Documents\\Docs - Importacao - Cigam\\CORRESPONDENCIAS.xlsx', index=False)

except Exception as e:
    print(f"Ocorreu um erro: {e}")