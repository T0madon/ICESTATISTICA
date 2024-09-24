import psycopg2

# Conecta ao servidor PostgreSQL (se ainda não conectado)
try:
    connection = psycopg2.connect(
        dbname="postgres",  # Nome do banco de dados padrão
        user="postgres",    # Nome de usuário padrão
        password="Gumattos2",
        host="localhost",
        port="5432"
    )
    connection.autocommit = True
    cursor = connection.cursor()

    # Cria o banco de dados "banco2"
    cursor.execute("CREATE DATABASE banco2;")
    print("Banco de dados 'banco2' criado com sucesso.")

except psycopg2.Error as e:
    print("Erro ao conectar ou criar o banco de dados:", e)

finally:
    # Fecha a conexão com o servidor PostgreSQL
    if connection:
        connection.close()
