import pandas as pd

# Carregar as planilhas, começando a leitura a partir da linha 3
planilha_colaboradores = pd.read_excel('COLABORADORES_SECATE.xlsx', header=2)
planilha_efetivos = pd.read_excel('Efetivos_SECATE.xlsx', header=2)

# Extrair nomes das colunas 'nm_pessoa_completo'
nomes_colaboradores = planilha_colaboradores['nm_pessoa_completo'].tolist()
nomes_efetivos = planilha_efetivos['nm_pessoa_completo'].tolist()

# Unir as listas de nomes, garantindo que não haja repetições
todos_os_nomes = list(set(nomes_colaboradores + nomes_efetivos))

# Criar um DataFrame com os nomes únicos
df_nomes_unicos = pd.DataFrame({'Nomes': todos_os_nomes})

# Salvar a nova planilha
df_nomes_unicos.to_excel('Nomes_Unicos.xlsx', index=False)

print("Planilha 'Nomes_Unicos.xlsx' criada com sucesso.")
