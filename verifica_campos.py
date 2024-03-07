import pandas as pd

# Verifica se os tamanhos dos campos estão de acordo com os parâmetros da importação

df = pd.read_excel('C:\\Users\\User\\Documents\\Docs - Importacao - Cigam\\EMPRESAS_MODELOONDEPERMANECE O COD ANTERI - Principal.xlsx')

for index, valor in df['Endereco'].items():
    if len(str(valor)) > 40:
        print(f"A linha {index} excede 40 caracteres: {valor}")