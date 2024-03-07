import pandas as pd

# Busca remover duplicatas, tomando como base algum identificador

def remover_duplicatas_pelo_cnpj(planilha_path):
    planilha = pd.read_csv(planilha_path)

    planilha_sem_duplicatas = planilha.drop_duplicates(subset='CD_EMPRESA', keep='first')

    planilha_sem_duplicatas.to_csv(planilha_path, index=False)

remover_duplicatas_pelo_cnpj('C:\\Users\\User\\Documents\\Docs - Importacao - Cigam\\CONTATO-EMAIL_MODELOONDEPERMANECE O COD ANTERI.csv')
