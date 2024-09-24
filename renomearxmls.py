import os
import xml.etree.ElementTree as ET

# Caminho da pasta que contém os arquivos XML
caminho_pasta = 'xmls'

# Função para extrair o valor do atributo NOME-COMPLETO
def extrair_nome_completo(arquivo_xml):
    try:
        tree = ET.parse(arquivo_xml)
        root = tree.getroot()
        nome_completo = root.find(".//DADOS-GERAIS").attrib['NOME-COMPLETO']
        return nome_completo
    except Exception as e:
        print(f"Erro ao processar o arquivo {arquivo_xml}: {e}")
        return None

# Função para gerar um nome de arquivo único
def gerar_nome_unico(caminho_pasta, nome_base):
    nome_arquivo = f"{nome_base}.xml"
    contador = 2
    while os.path.exists(os.path.join(caminho_pasta, nome_arquivo)):
        nome_arquivo = f"{nome_base}{contador}.xml"
        contador += 1
    return nome_arquivo

# Renomear os arquivos XML
for arquivo in os.listdir(caminho_pasta):
    if arquivo.endswith('.xml'):
        caminho_arquivo = os.path.join(caminho_pasta, arquivo)
        nome_completo = extrair_nome_completo(caminho_arquivo)
        if nome_completo:
            nome_arquivo_unico = gerar_nome_unico(caminho_pasta, nome_completo)
            caminho_novo_arquivo = os.path.join(caminho_pasta, nome_arquivo_unico)
            os.rename(caminho_arquivo, caminho_novo_arquivo)
            print(f"Arquivo {arquivo} renomeado para {nome_arquivo_unico}")

print("Processo de renomeação concluído.")
