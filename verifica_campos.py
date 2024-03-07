import pandas as pd

# Verifica se os tamanhos dos campos estão de acordo com os parâmetros da importação

df = pd.read_excel('C:\\Users\\User\\Documents\\Docs - Importacao - Cigam\\EMPRESAS_MODELOONDEPERMANECE O COD ANTERI(1) - Principal - Copia.xlsx')

# Iterar sobre os valores da coluna e imprimir aqueles que excedem 60 caracteres
for index, valor in df['Endereco'].items():
    if len(str(valor)) > 40:
        print(f"A linha {index} excede 60 caracteres: {valor}")