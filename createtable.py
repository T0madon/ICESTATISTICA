import psycopg2

def conecta_db():
    try:
        con = psycopg2.connect(database="banco2",
                                host="localhost",
                                user="postgres",
                                password="Gumattos2",
                                port="5432")
        return con
    except psycopg2.Error as e:
        print("Erro ao conectar ao banco de dados:", e)
        return None

def criar_db(sql):
    con = conecta_db()
    if con:
        cur = con.cursor()
        try:
            cur.execute(sql)
            con.commit()
            print("Tabela criada com sucesso.")
        except psycopg2.Error as e:
            print("Erro ao criar tabela:", e)
        finally:
            con.close()

# Criação da tabela de professores
sql_professores = '''
    DROP TABLE IF EXISTS public.professores;
    CREATE TABLE public.professores (
        id_professor  UUID PRIMARY KEY,         
        nome          character varying(100), 
        anosUEPG      character varying(100), 
        graduacao     character varying(20),
        departamento  character varying(500), 
        status        character varying(12),
        tide          character varying(5),
        setor         character varying(10)
    );
'''
criar_db(sql_professores)

# Criação da tabela de artigos
sql_artigos = '''
    DROP TABLE IF EXISTS public.artigos;
    CREATE TABLE public.artigos (
        id_professor  character varying(5000),         
        nome          character varying(1000), 
        anopubli      character varying(10), 
        tipo          character varying(30),
        departamento  character varying(500)
    );
'''
criar_db(sql_artigos)

# Criação da tabela de projetos
sql_projetos = '''
    DROP TABLE IF EXISTS public.projetos;
    CREATE TABLE public.projetos (
        id_professor  character varying(5000),         
        nome          character varying(1000), 
        anopubli      character varying(10), 
        tipo          character varying(30),
        departamento  character varying(500)
    );
'''
criar_db(sql_projetos)

# Criação da tabela de orientações
sql_orientacoes = '''
    DROP TABLE IF EXISTS public.orientacoes;
    CREATE TABLE public.orientacoes (
        id_professor  character varying(500),
        nome          character varying(1000),          
        anoconclusao  character varying(10), 
        tipo          character varying(30),
        departamento  character varying(500)
    );
'''
criar_db(sql_orientacoes)

# Criação da tabela de publicações em congresso
sql_congressos = '''
    DROP TABLE IF EXISTS public.congressos;
    CREATE TABLE public.congressos (
        id_professor  character varying(5000),
        nome          character varying(1000),          
        anoconclusao  character varying(10), 
        tipo          character varying(30),
        departamento  character varying(500)
    );
'''
criar_db(sql_congressos)

# Criação da tabela de bolsas
sql_bolsas = '''
    DROP TABLE IF EXISTS public.bolsas;
    CREATE TABLE public.bolsas (
        id_professor  character varying(500),
        nome          character varying(1000),          
        ano           character varying(10), 
        departamento  character varying(500)
    );
'''
criar_db(sql_bolsas)

# Criação da tabela de financiados
sql_financiados = '''
    DROP TABLE IF EXISTS public.financiados;
    CREATE TABLE public.financiados (
        id_professor  character varying(500),         
        nome          character varying(1000), 
        anopubli      character varying(10), 
        valor         FLOAT,
        departamento  character varying(500)
    );
'''
criar_db(sql_financiados)

# Criação da tabela de produtividade
sql_produtividade = '''
    DROP TABLE IF EXISTS public.produtividade;
    CREATE TABLE public.produtividade (
        id_professor  character varying(500),         
        ano           character varying(10), 
        departamento  character varying(500)
    );
'''
criar_db(sql_produtividade)