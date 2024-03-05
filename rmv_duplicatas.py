import pandas as pd

def remover_duplicatas_pelo_cnpj(planilha_path):
    planilha = pd.read_excel(planilha_path)

    planilha_sem_duplicatas = planilha.drop_duplicates(subset='CNPJ_CPF', keep='first')

    planilha_sem_duplicatas.to_excel(planilha_path, index=False)

remover_duplicatas_pelo_cnpj('C:\\Users\\User\\Documents\\Docs - Importacao - Cigam\\EMPRESAS_MODELOONDEPERMANECE O COD ANTERI(1) - Principal.xlsx')
