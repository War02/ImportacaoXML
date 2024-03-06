import pandas as pd
#
# # Carregue o arquivo Excel
# df = pd.read_excel('C:\\Users\\User\\Documents\\Docs - Importacao - Cigam\\EMPRESAS_MODELOONDEPERMANECE O COD ANTERI(1) - Principal - Copia.xlsx')
#
# # Remova os caracteres após a primeira vírgula na coluna 'Sua_Coluna'
# df['Endereco'] = df['Endereco'].str.split(',', expand=True)[0]
#
# # Salve o DataFrame de volta no arquivo Excel
# df.to_excel('C:\\Users\\User\\Documents\\Docs - Importacao - Cigam\\EMPRESAS_MODELOONDEPERMANECE O COD ANTERI(1) - Principal - Copia.xlsx', index=False)


# Carregue o arquivo Excel
df = pd.read_excel('C:\\Users\\User\\Documents\\Docs - Importacao - Cigam\\EMPRESAS_MODELOONDEPERMANECE O COD ANTERI(1) - Principal - Copia.xlsx')

# Iterar sobre os valores da coluna e imprimir aqueles que excedem 60 caracteres
for index, valor in df['Endereco'].items():
    if len(str(valor)) > 40:
        print(f"A linha {index} excede 60 caracteres: {valor}")