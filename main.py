import xml.etree.ElementTree as ET
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

def find_text(root, namespace, path):
    try:
        return root.find(path, namespace).text
    except AttributeError:
        return None

xml_path = 'C:\\Users\\User\\Downloads\\TesteXml\\35190104774509000142550050000603701000603701-nfe.xml'
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
excel_path = 'C:\\Users\\User\\Documents\\EMPRESAS_MODELOONDEPERMANECE O COD ANTERI(1).xlsx'

wb = load_workbook(filename=excel_path)
ws = wb.active

next_row = ws.max_row + 1

for key, value in data.items():
    column_name = column_mapping.get(key)
    if column_name:
        for cell in ws[1]:
            if cell.value == column_name:
                ws[f"{cell.column_letter}{next_row}"] = value
                break

wb.save(excel_path)

print("Dados adicionados com sucesso à planilha.")
