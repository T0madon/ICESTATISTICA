import os
from openpyxl import load_workbook
import unidecode

def remover_acentos(texto):
    return unidecode.unidecode(texto) if texto else texto

def encontrar_docente(nome):
    pasta = "FINANCIADOS"
    # Lista todos os arquivos na pasta especificada
    arquivos = os.listdir(pasta)
    
    # Itera sobre cada arquivo na pasta
    for arquivo in arquivos:
        # Verifica se o arquivo é uma planilha Excel
        if arquivo.endswith('.xlsx'):
            # Carrega a planilha
            wb = load_workbook(os.path.join(pasta, arquivo))
            # Acessa a primeira planilha (índice 0)
            planilha = wb.worksheets[0]
            
            # Itera sobre as células das colunas I (nome do docente), B (nome do projeto) e G (valor do projeto) a partir da linha 3
            for docente, projeto, valor in zip(planilha['I'][2:], planilha['B'][2:], planilha['G'][2:]):
                nome_docente = remover_acentos(docente.value)
                nome_projeto = remover_acentos(projeto.value)
                valor_projeto = remover_acentos(valor.value)
                if nome_docente == remover_acentos(nome):
                    # Remove o "R$" do valor do projeto
                    valor_projeto = valor_projeto.replace('R$', '').strip()
                    print(f'O docente {nome} foi encontrado no arquivo: {arquivo}. Projeto: {nome_projeto}, Valor: {valor_projeto}')
                    # Aqui você pode adicionar o valor a uma lista, dicionário ou outra estrutura de dados, conforme necessário
                    # Por exemplo, para adicionar a uma lista:
                    # valores.append(float(valor_projeto))

# Nome do docente que você quer buscar
nome_docente = "Andressa Novatski"

# Chama a função para buscar o docente na pasta "FINANCIADOS"
encontrar_docente(nome_docente)
