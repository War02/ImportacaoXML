import pandas as pd

df = pd.read_excel('C:\\Users\\User\\Documents\\Docs - Importacao - Cigam\\EMPRESAS_MODELOONDEPERMANECE O COD ANTERI(1) - Principal - Copia.xlsx')

df['Endereco'] = df['Endereco'].str.split(',', expand=True)[0]

df.to_excel('C:\\Users\\User\\Documents\\Docs - Importacao - Cigam\\EMPRESAS_MODELOONDEPERMANECE O COD ANTERI(1) - Principal - Copia.xlsx', index=False)
