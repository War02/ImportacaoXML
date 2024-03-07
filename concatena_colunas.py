import pandas as pd

# Concatena duas colunas gerando uma nova e separando os itens por vírgula

planilha_emails = pd.read_excel('C:\\Users\\User\\Documents\\Docs - Importacao - Cigam\\CORRESPONDENCIASFORNECEDORES.xlsx')

planilha_emails['EMAIL'] = planilha_emails['email'].astype(str) + ',' + planilha_emails['email-nf'].astype(str)

planilha_emails.to_excel('C:\\Users\\User\\Documents\\Docs - Importacao - Cigam\\CORRESPONDENCIASFORNECEDORES.xlsx', index=False)