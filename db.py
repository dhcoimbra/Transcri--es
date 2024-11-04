import psycopg2
import os

# Função para conectar ao banco de dados PostgreSQL
def conectar():
    try:
        conn = psycopg2.connect(
            host=os.getenv('DB_HOST', 'localhost'),
            database=os.getenv('DB_NAME', 'transcricoes_db'),
            user=os.getenv('DB_USER', 'postgres'),
            password=os.getenv('DB_PASSWORD', '1234'),
            port=os.getenv('DB_PORT', '5432')
        )
        conn.set_client_encoding('UTF8')
        return conn
    except psycopg2.Error as e:
        print(f"Erro ao conectar ao banco de dados: {e}")
        raise


# Função para criar a tabela de transcrições
def criar_tabela_transcricoes():
    try:
        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS transcricoes (
            hash_arquivo TEXT PRIMARY KEY,
            transcricao TEXT,
            from_field TEXT,
            to_field TEXT,
            timestamp TIMESTAMP
        );
    """)
        conn.commit()
        cursor.close()
        conn.close()
        #print("Tabela 'transcricoes' criada ou já existe.")
    except psycopg2.Error as e:
        print(f"Erro ao criar a tabela: {e}")
        raise


# Função para salvar uma nova transcrição
def salvar_transcricao(hash_arquivo, transcricao, from_field, to_field, timestamp):
    try:
        conn = conectar()
        cursor = conn.cursor()

        # Verificar os dados a serem inseridos
        #print(f"Inserindo no banco: filename={audio_filename}, transcription={transcription[:30]}...")

        cursor.execute(
        """
        INSERT INTO transcricoes (hash_arquivo, transcricao, from_field, to_field, timestamp)
        VALUES (%s, %s, %s, %s, %s)
        ON CONFLICT (hash_arquivo) DO NOTHING
        """,
        (hash_arquivo, transcricao, from_field, to_field, timestamp)
    )
        conn.commit()

        #print(f"Transcrição salva no banco de dados: {audio_filename}")
        cursor.close()
        conn.close()

    except psycopg2.Error as e:
        print(f"Erro ao salvar transcrição no banco de dados: {e}")
        raise



# Função para buscar uma transcrição já existente
def buscar_transcricao(hash_arquivo):
    try:
        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("SELECT transcricao FROM transcricoes WHERE hash_arquivo = %s", (hash_arquivo,))
        result = cursor.fetchone()
        cursor.close()
        conn.close()
        return result[0] if result else None
    except psycopg2.Error as e:
        print(f"Erro ao buscar transcrição no banco de dados: {e}")
        raise
