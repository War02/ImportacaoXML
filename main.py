import os
import xml.etree.ElementTree as ET
from openpyxl import load_workbook

# Percorre pastas com nomes personalizados, identifica os arquivos xml, captura somente os dados desejados
# contidos nas tags e os salva em uma planilha
def extrair_informacao_xml(arquivo_xml):
    try:
        tree = ET.parse(arquivo_xml)
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
        return data
    except ET.ParseError:
        return None

def find_text(root, namespace, path):
    try:
        return root.find(path, namespace).text
    except AttributeError:
        return None

def buscar_em_pasta(pasta):
    excel_path = 'C:\\Users\\User\\Documents\\EMPRESAS_MODELOONDEPERMANECE O COD ANTERI(Final).xlsx'

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

    wb = load_workbook(filename=excel_path)
    ws = wb.active

    for pasta_atual, _, arquivos in os.walk(pasta):
        for arquivo in arquivos:
            if arquivo.endswith('.xml'):
                caminho_arquivo = os.path.join(pasta_atual, arquivo)
                data = extrair_informacao_xml(caminho_arquivo)
                if data:
                    cnpj = data['cnpj']
                    for row in ws.iter_rows(min_row=2, max_col=1, max_row=ws.max_row, values_only=True):
                        if row[0] == cnpj:
                            for key, value in data.items():
                                column_name = column_mapping.get(key)
                                if column_name:
                                    for cell in ws[1]:
                                        if cell.value == column_name:
                                            col_letter = cell.column_letter
                                            ws[f"{col_letter}{row[0].row}"] = value
                                            break
                            break
                    else:
                        next_row = ws.max_row + 1
                        for key, value in data.items():
                            column_name = column_mapping.get(key)
                            if column_name:
                                for cell in ws[1]:
                                    if cell.value == column_name:
                                        col_letter = cell.column_letter
                                        ws[f"{col_letter}{next_row}"] = value
                                        break
    wb.save(excel_path)

    print("Dados dos arquivos XML adicionados com sucesso à planilha.")

for ano in range(2019, 2025):
    for mes in range(1, 13):
        nome_pasta = f"{ano}{mes:02d}"
        pasta_atual = os.path.join('C:\\Users\\User\\Downloads\\RaizTesteXml', nome_pasta)
        if os.path.exists(pasta_atual):
            buscar_em_pasta(pasta_atual)

# fazer a tratativa para quando o número estiver junto com endereço