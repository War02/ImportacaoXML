import pandas as pd

# Verifica se todos os nomes que estão contidos em uma tabela, estão contidos na outra;

def comparar_nome(planilha1_path, planilha2_path):
    planilha1 = pd.read_excel(planilha1_path)
    planilha2 = pd.read_excel(planilha2_path)

    nomes_planilha1 = set(planilha1['CNPJ_CPF'])
    nomes_planilha2 = set(planilha2['CNPJ_CPF'])

    nome_somente_planilha1 = nomes_planilha1 - nomes_planilha2

    nome_somente_planilha2 = nomes_planilha2 - nomes_planilha1

    if not nome_somente_planilha1 and not nome_somente_planilha2:
        print("\nTodos os nomes existem")

    else:
        print("\nNomes somente planilha 1:")
        for nome in nome_somente_planilha1:
            print(nome)

        print("\nNomes somente planilha 2:")
        for nome in nome_somente_planilha1:
            print(nome)

comparar_nome('\\\\192.168.10.51\\Dados\\Departamental\\TI\\!CIGAM\\Migracao\\Dados Usados para Migração\\Docs - Importacao - Cigam\\EMPRESAS_MODELOONDEPERMANECE O COD ANTERI - PRINCIPAL - Copia.xlsx', '\\\\192.168.10.51\\Dados\\Departamental\\TI\\!CIGAM\\Migracao\\Dados Coletados para Migração\\Inscricao.xlsx')