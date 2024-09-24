import os
import pandas as pd
import unidecode
import psycopg2

# Defina o nome do pesquisador
nome_pesquisador = "Andressa Novatski"

# Diretório onde estão os arquivos .xlsx
diretorio = "PRODUTIVIDADE"

# Função para gerar array de anos
def gerar_anos(inicio, termino):
    anos = []
    ano_inicio = inicio.year
    ano_termino = termino.year
    for ano in range(ano_inicio, ano_termino + 1):
        anos.append(str(ano))
    return anos

# Inicializa a lista de anos de pesquisa
anos_pesquisa = []

# Conexão com o banco de dados
# Conexão com o banco de dados
def inserir_dados(id_professor, ano, departamento):
    # Verifica se já existe uma linha com o mesmo id_professor e ano
    cursor.execute('''SELECT 1 FROM public.produtividade 
                        WHERE id_professor = %s AND ano = %s;''', (id_professor, ano))
    row = cursor.fetchone()
    if row:
        print(f"Já existe uma linha para o professor {id_professor} no ano {ano}.")
    else:
        # Insere os dados apenas se não houver linha duplicada
        cursor.execute('''INSERT INTO public.produtividade (id_professor, ano, departamento) 
                            VALUES (%s, %s, %s);''', (id_professor, ano, departamento))
        conn.commit()
        print(f"Dados inseridos para o professor {id_professor} no ano {ano}.")

    cursor.close()
    conn.close()



# Percorre todos os arquivos .xlsx no diretório especificado
for arquivo in os.listdir(diretorio):
    if arquivo.endswith(".xlsx"):
        # Carrega o arquivo Excel
        caminho_arquivo = os.path.join(diretorio, arquivo)
        df = pd.read_excel(caminho_arquivo)

        # Remove acentos dos nomes dos pesquisadores
        df['PESQUISADOR'] = df['PESQUISADOR'].apply(lambda x: unidecode.unidecode(str(x)))

        # Filtra as linhas onde o nome do pesquisador corresponde
        df_pesquisador = df[df['PESQUISADOR'] == unidecode.unidecode(nome_pesquisador)]

        # Para cada linha do pesquisador, calcula os anos de pesquisa
        for index, row in df_pesquisador.iterrows():
            data_inicio = pd.to_datetime(row['INÍCIO'], format='%d/%m/%Y')
            data_termino = pd.to_datetime(row['TÉRMINO'], format='%d/%m/%Y')
            anos = gerar_anos(data_inicio, data_termino)
            anos_pesquisa.extend(anos)

            # Inserir os dados no banco para cada ano
            for ano in anos:
                id_professor = row['PESQUISADOR']  # Ajuste conforme necessário
                departamento = "Seu Departamento"  # Ajuste conforme necessário
                inserir_dados(id_professor, ano, departamento)

# Remove anos duplicados
anos_pesquisa = list(set(anos_pesquisa))
print(anos_pesquisa)
