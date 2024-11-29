#import psycopg2
import os
import sqlite3

# Função para conectar ao banco de dados PostgreSQL
"""def conectar():
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
        raise"""


# Função para criar a tabela de transcrições
def criar_tabela_transcricoes_ANTIGO(): #função antiga
    try:
        conn = sqlite3.connect('transcricoes_db.db')
        #conn = conectar()
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
    except sqlite3.Error as e:
        print(f"Erro ao criar a tabela: {e}")
        raise

def criar_tabela_transcricoes():
    # Caminho onde o banco de dados será armazenado
    caminho_bd = os.path.join('config', 'transcricoes_db.db')
    
    # Verifica se o banco de dados já existe
    if os.path.exists(caminho_bd):
        print(f"O banco de dados '{caminho_bd}' já existe. Nenhuma ação foi realizada.")
        return
    
    # Se não existir, cria a tabela
    try:
        # Certifique-se de que a pasta 'config' existe
        os.makedirs('config', exist_ok=True)
        
        conn = sqlite3.connect(caminho_bd)
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
        print("Tabela 'transcricoes' criada com sucesso.")
    except sqlite3.Error as e:
        print(f"Erro ao criar a tabela: {e}")
        raise
    
    
# Função para salvar uma nova transcrição
def salvar_transcricao(hash_arquivo, transcricao, from_field, to_field, timestamp):
    try:
        #conn = conectar()
        conn = sqlite3.connect('config/transcricoes_db.db')
        cursor = conn.cursor()

        # Verificar os dados a serem inseridos
        #print(f"Inserindo no banco: filename={audio_filename}, transcription={transcription[:30]}...")
        
        cursor.execute(
        """
        INSERT OR IGNORE INTO transcricoes (hash_arquivo, transcricao, from_field, to_field, timestamp)
        VALUES (?, ?, ?, ?, ?)
        """,
        (hash_arquivo, transcricao, from_field, to_field, timestamp)
    )
        conn.commit()

        #print(f"Transcrição salva no banco de dados: {audio_filename}")
        cursor.close()
        conn.close()

    except sqlite3.Error as e:
        print(f"Erro ao salvar transcrição no banco de dados: {e}")
        raise



# Função para buscar uma transcrição já existente
def buscar_transcricao(hash_arquivo):
    try:
        #conn = conectar()
        conn = sqlite3.connect('config/transcricoes_db.db')
        cursor = conn.cursor()
        cursor.execute("SELECT transcricao FROM transcricoes WHERE hash_arquivo = ?", (hash_arquivo,))
        result = cursor.fetchone()
        cursor.close()
        conn.close()
        return result[0] if result else None
    except sqlite3.Error as e:
        print(f"Erro ao buscar transcrição no banco de dados: {e}")
        raise
