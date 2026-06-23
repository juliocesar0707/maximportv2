# database.py
from sqlalchemy import create_engine, text
import config
import urllib.parse
import pyodbc
import os
import pandas as pd
from datetime import datetime

# Variável global da conexão
engine = None

def detectar_driver():
    """Identifica o driver ODBC disponível"""
    try:
        drivers = [d for d in pyodbc.drivers() if 'SQL Server' in d]
    except:
        return 'SQL Server'
    
    preferencias = [
        'ODBC Driver 18 for SQL Server',
        'ODBC Driver 17 for SQL Server',
        'ODBC Driver 13 for SQL Server',
        'SQL Server Native Client 11.0',
        'SQL Server'
    ]
    
    for pref in preferencias:
        if pref in drivers:
            return pref
    
    if drivers: return drivers[0]
    return 'SQL Server'

def get_connection_string(banco=None):
    driver = detectar_driver()
    db_target = banco if banco else config.DB_NAME
    
    params = urllib.parse.quote_plus(
        f"DRIVER={{{driver}}};"
        f"SERVER={config.DB_SERVER};"
        f"DATABASE={db_target};"
        f"UID={config.DB_USER};"
        f"PWD={config.DB_PASS};"
        f"TrustServerCertificate=yes;"
    )
    return f"mssql+pyodbc:///?odbc_connect={params}"

def get_engine():
    global engine
    if engine is None:
        reconectar()
    return engine

def reconectar():
    global engine
    try:
        if engine:
            engine.dispose()
        conn_str = get_connection_string()
        # Removido fast_executemany temporariamente se estiver causando problemas com IDENTITY
        engine = create_engine(conn_str) 
        print(f"Conectado ao banco: {config.DB_NAME} em {config.DB_SERVER}")
    except Exception as e:
        print(f"Erro ao configurar conexão: {e}")
        raise e

def listar_bancos_disponiveis(servidor):
    """Lista bancos usando método híbrido de autenticação"""
    driver = detectar_driver()
    configs = [
        f"UID={config.DB_USER};PWD={config.DB_PASS};TrustServerCertificate=yes;",
        f"Trusted_Connection=yes;TrustServerCertificate=yes;"
    ]
    
    for auth in configs:
        try:
            params = urllib.parse.quote_plus(f"DRIVER={{{driver}}};SERVER={servidor};DATABASE=master;{auth}")
            eng = create_engine(f"mssql+pyodbc:///?odbc_connect={params}")
            with eng.connect() as conn:
                result = conn.execute(text("SELECT name FROM sys.databases WHERE name NOT IN ('master','tempdb','model','msdb') AND state_desc='ONLINE' ORDER BY name"))
                return [row[0] for row in result]
        except:
            continue
    raise Exception("Não foi possível listar os bancos (Falha de Login).")

def executar_comando(sql_cmd):
    with get_engine().begin() as conn:
        conn.execute(text(sql_cmd))

def toggle_constraints(enable=True):
    tabelas_alvo = ['cliente', 'produto', 'produto_empresa', 'vendaPgto', 'fornecedor', 'ncm', 'proncm']
    
    if enable:
        print("🔒 Reativando GATILHOS e CHECAGENS...")
        cmd_fk = "CHECK CONSTRAINT ALL"
        cmd_tr = "ENABLE TRIGGER ALL"
    else:
        print("🔓 Desativando GATILHOS (Triggers) e FKs para importação...")
        cmd_fk = "NOCHECK CONSTRAINT ALL"
        cmd_tr = "DISABLE TRIGGER ALL"

    with get_engine().begin() as conn:
        try:
            conn.execute(text(f"EXEC sp_msforeachtable 'ALTER TABLE ? {cmd_fk}'"))
            conn.execute(text(f"EXEC sp_msforeachtable 'ALTER TABLE ? {cmd_tr}'"))
        except Exception as e:
            print(f"Aviso no método geral: {e}")

        for tab in tabelas_alvo:
            try:
                conn.execute(text(f"ALTER TABLE {tab} {cmd_fk}"))
                conn.execute(text(f"ALTER TABLE {tab} {cmd_tr}"))
            except: pass

def limpar_tabela(nome_tabela, reset_identity=False):
    try:
        executar_comando(f"DELETE FROM {nome_tabela}")
        if reset_identity:
            try:
                executar_comando(f"DBCC CHECKIDENT ('{nome_tabela}', RESEED, 0)")
            except: pass
        print(f"Tabela {nome_tabela} limpa.")
    except Exception as e:
        print(f"Erro ao limpar {nome_tabela}: {e}")

def inserir_bulk(df, nome_tabela, manter_id=True):
    if df.empty: return

    # Usa .connect() com controle manual de transação para garantir IDENTITY_INSERT
    with get_engine().connect() as conn:
        transaction = conn.begin()
        try:
            if manter_id:
                # O comando SET IDENTITY_INSERT precisa estar na mesma sessão
                conn.execute(text(f"SET IDENTITY_INSERT {nome_tabela} ON"))
            
            # chunksize ajustado para estabilidade
            df.to_sql(nome_tabela, con=conn, if_exists='append', index=False, chunksize=200)
            
            if manter_id:
                conn.execute(text(f"SET IDENTITY_INSERT {nome_tabela} OFF"))
            
            transaction.commit()
            print(f"Importado: {len(df)} registros em {nome_tabela}")
        except Exception as e:
            transaction.rollback()
            print(f"Erro no bulk insert em {nome_tabela}. Tentando recuperação linha a linha para capturar erros...")
            _inserir_linha_a_linha(df, nome_tabela, manter_id)

def _inserir_linha_a_linha(df, nome_tabela, manter_id):
    erros = []
    sucesso_count = 0
    with get_engine().connect() as conn:
        if manter_id:
            try:
                conn.execute(text(f"SET IDENTITY_INSERT {nome_tabela} ON"))
            except:
                pass
            
        for index, row in df.iterrows():
            transaction = conn.begin()
            try:
                # Insere apenas uma linha
                pd.DataFrame([row]).to_sql(nome_tabela, con=conn, if_exists='append', index=False)
                transaction.commit()
                sucesso_count += 1
            except Exception as e:
                transaction.rollback()
                linha_com_erro = row.to_dict()
                # Salva o texto do erro até 500 caracteres para ficar claro
                linha_com_erro['_erro_banco'] = str(e).split('\n')[0][:500]
                erros.append(linha_com_erro)
                
        if manter_id:
            try:
                conn.execute(text(f"SET IDENTITY_INSERT {nome_tabela} OFF"))
            except:
                pass
            
    print(f"Resumo da recuperação em {nome_tabela}: {sucesso_count} linhas salvas, {len(erros)} falhas.")
    
    if erros:
        # Gera arquivo Excel de log
        df_erros = pd.DataFrame(erros)
        os.makedirs("logs", exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        arquivo_erro = f"logs/erros_{nome_tabela}_{timestamp}.xlsx"
        df_erros.to_excel(arquivo_erro, index=False)
        print(f"🚨 ATENÇÃO: {len(erros)} registros não puderam ser importados.")
        print(f"🚨 LOG SALVO EM: {arquivo_erro} para correção no Excel e reimportação.")