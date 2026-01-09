# database.py
from sqlalchemy import create_engine, text
import config

# Variável global da conexão
engine = None

def get_engine():
    """Retorna a engine atual ou cria uma nova se não existir"""
    global engine
    if engine is None:
        reconectar()
    return engine

def reconectar():
    """Força a recriação da conexão com os dados atuais do config.py"""
    global engine
    try:
        if engine:
            engine.dispose() # Fecha a anterior
        
        # Cria nova com os dados atuais do config
        engine = create_engine(config.get_connection_string(), fast_executemany=True)
        print(f"Conectado ao banco: {config.DB_NAME} em {config.DB_SERVER}")
    except Exception as e:
        print(f"Erro ao configurar conexão: {e}")

def executar_comando(sql_cmd):
    """Executa comandos SQL diretos (Delete, Update, Insert manual)"""
    with get_engine().begin() as conn:
        conn.execute(text(sql_cmd))

def toggle_constraints(enable=True):
    acao = "CHECK" if enable else "NOCHECK"
    executar_comando(f"EXEC sp_msforeachtable 'ALTER TABLE ? {acao} CONSTRAINT ALL'")

def limpar_tabela(nome_tabela, reset_identity=False):
    try:
        executar_comando(f"DELETE FROM {nome_tabela}")
        if reset_identity:
            executar_comando(f"DBCC CHECKIDENT ('{nome_tabela}', RESEED, 0)")
        print(f"Tabela {nome_tabela} limpa com sucesso.")
    except Exception as e:
        print(f"Erro ao limpar {nome_tabela}: {e}")

def inserir_bulk(df, nome_tabela, manter_id=True):
    if df.empty: return

    with get_engine().connect() as conn:
        transaction = conn.begin()
        try:
            if manter_id:
                conn.execute(text(f"SET IDENTITY_INSERT {nome_tabela} ON"))
            
            df.to_sql(nome_tabela, con=conn, if_exists='append', index=False, chunksize=500)
            
            if manter_id:
                conn.execute(text(f"SET IDENTITY_INSERT {nome_tabela} OFF"))
            
            transaction.commit()
            print(f"Importado: {len(df)} registros em {nome_tabela}")
        except Exception as e:
            transaction.rollback()
            raise e