import pandas as pd
import xml.etree.ElementTree as ET

def find_text(root, namespace, path):
    try:
        return root.find(path, namespace).text
    except AttributeError:
        return None

xml_path = 'C:\\Users\\User\\Downloads\\TesteXml\\35190104774509000142550050000603691000603697-nfe.xml'
tree = ET.parse(xml_path)
root = tree.getroot()

ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

data = {
    'xNome': find_text(root, ns, './/nfe:dest/nfe:xNome'),
    'fone': find_text(root, ns, './/nfe:dest/nfe:enderDest/nfe:fone'),
    'xLgr': find_text(root, ns, './/nfe:dest/nfe:enderDest/nfe:xLgr'),
    'xBairro': find_text(root, ns, './/nfe:dest/nfe:enderDest/nfe:xBairro'),
    'xMun': find_text(root, ns, './/nfe:dest/nfe:enderDest/nfe:xMun'),
    'uf': find_text(root, ns, './/nfe:dest/nfe:enderDest/nfe:UF'),
    'cep': find_text(root, ns, './/nfe:dest/nfe:enderDest/nfe:CEP'),
    'cnpj': find_text(root, ns, './/nfe:dest/nfe:CNPJ'),
    'cPais': find_text(root, ns, './/nfe:dest/nfe:enderDest/nfe:cPais'),
    'nro': find_text(root, ns, './/nfe:dest/nfe:enderDest/nfe:nro')
}
column_mapping = {
    'xNome': 'Nome Completo',
    'fone': 'Fone',
    'xLgr': 'Endereco',
    'xBairro': 'Bairro',
    'xMun': 'Municipio',
    'uf': 'UF',
    'cep': 'CEP',
    'cnpj': 'CNPJ_CPF',
    'cPais': 'Codigo Pais',
    'nro': 'Numero'
}

# Caminho para o arquivo Excel existente
excel_path = 'C:\\Users\\User\\Documents\\EMPRESAS_MODELOONDEPERMANECE O COD ANTERI.xlsx'

# Carrega a planilha existente para um DataFrame
df = pd.read_excel(excel_path)

# Cria um novo DataFrame com os dados do XML mapeados para as colunas correspondentes
new_row = {column_mapping[key]: value for key, value in data.items()}
new_df = pd.DataFrame([new_row])

# Adiciona a nova linha ao final do DataFrame original, mantendo as colunas existentes
df = pd.concat([df, new_df], ignore_index=True, sort=False)

# Salva o DataFrame atualizado de volta para o arquivo Excel
df.to_excel(excel_path, index=False)

print("Nova linha adicionada com sucesso ao arquivo Excel existente.")
