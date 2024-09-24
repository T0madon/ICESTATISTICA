import os
import pandas as pd
from unidecode import unidecode  # Importa a função unidecode para remover acentos

# Obtém o diretório atual onde está o código
diretorio_atual = os.path.dirname(os.path.abspath(__file__))

# Pasta onde estão os dados
pasta_bolsasic = "BOLSASIC"

# Percorre os anos dentro da pasta BOLSASIC
for ano in range(2019, 2025):  # Modifique conforme os anos desejados
    pasta_ano = os.path.join(diretorio_atual, pasta_bolsasic, str(ano))
    if os.path.exists(pasta_ano):
        # Lista os arquivos na pasta do ano
        arquivos = os.listdir(pasta_ano)
        # Itera sobre os arquivos
        for arquivo in arquivos:
            # Verifica se o arquivo é uma planilha (pode adaptar essa verificação conforme o tipo de arquivo)
            if arquivo.endswith('.xlsx') or arquivo.endswith('.xls'):
                # Lê a planilha
                caminho_arquivo = os.path.join(pasta_ano, arquivo)
                
                # Lê a planilha, especificando que o cabeçalho está na linha 8 (0-indexed)
                df = pd.read_excel(caminho_arquivo, header=8)
                
                # Remove os acentos dos nomes na coluna 'ORIENTADOR(A)' e 'SUBPROJETO DO' da planilha
                df['ORIENTADOR(A)'] = df['ORIENTADOR(A)'].astype(str).apply(unidecode)
                df['SUBPROJETO DO'] = df['SUBPROJETO DO'].astype(str).apply(unidecode)
                
                # Filtra as linhas onde o nome do professor corresponde ao nome desejado
                df_filtrado = df[df['ORIENTADOR(A)'] == unidecode(nome)]
                
                # Itera sobre as linhas filtradas
                for index, row in df_filtrado.iterrows():
                    # Imprime o nome do trabalho e o ano da pasta
                    print("Professor:", row['ORIENTADOR(A)'])
                    print("Trabalho:", row['SUBPROJETO DO'])
                    print("Ano do trabalho:", ano)
                    titulobolsa = row['SUBPROJETO DO']
                    anobolsa = ano
                    print()  # Adiciona uma linha em branco entre as impressões
                    cursor.execute("SELECT COUNT(*) FROM bolsas WHERE nome = %s", (titulobolsa,))
                    titulobolsacomparar = cursor.fetchone()[0]
                    if titulobolsacomparar < 1:
                        cursor.execute('''INSERT into public.projetos (id_professor,nome,anopubli,departamento) 
                values('%s','%s','%s','%s','%s');
                    ''' % ((idprofessor),(titulobolsa),(anobolsa),(nomedepart)))
                        connection.commit()