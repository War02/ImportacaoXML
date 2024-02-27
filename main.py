import pandas as pd
import xml.etree.cElementTree as ET
from xml.dom import minidom
from Tools.scripts.dutree import display

xml_path = 'C:\\Users\\User\\Downloads\\TesteXml\\35190104774509000142550050000603691000603697-nfe.xml'
tree = ET.parse(xml_path)
root = tree.getroot()

ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

xNome = root.find('.//nfe:dest/nfe:xNome', ns)
fone = root.find('.//nfe:dest/nfe:fone', ns)
xLgr = root.find('.//nfe:dest/nfe:xLgr', ns)
xBairro = root.find('.//nfe:dest/nfe:xBairro', ns)
xMun = root.find('.//nfe:dest/nfe:xMun', ns)
uf = root.find('.//nfe:dest/nfe:xUF', ns)
cep = root.find('.//nfe:dest/nfe:CEP', ns)
cnpj = root.find('.//nfe:dest/nfe:CNPJ', ns)
cPais = root.find('.//nfe:dest/nfe:cPais', ns)
nro = root.find('.//nfe:dest/nfe:nro', ns)

if xNome is not None:
    print(xNome.text)
else:
    print("Tag 'xNome' não encontrada dentro da tag 'dest'")
