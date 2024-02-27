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
    'cnpj': find_text(root, ns, './/nfe:dest/nfe:enderDest/nfe:CNPJ'),
    'cPais': find_text(root, ns, './/nfe:dest/nfe:enderDest/nfe:cPais'),
    'nro': find_text(root, ns, './/nfe:dest/nfe:enderDest/nfe:nro')
}

df = pd.DataFrame(data, index=[0])

print(df)
