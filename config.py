# config.py
import urllib.parse

# CONFIGURAÇÕES DO SQL SERVER (DESTINO)
DB_SERVER = 'LOCALHOST'  # Ou nome do servidor da Maxdata
DB_NAME = 'centerf'  # Nome do banco de dados
DB_USER = 'sa'
DB_PASS = 'macro01'

# Variável global que guardará o caminho do arquivo selecionado
# Agora começa vazia ou None
ARQUIVO_SELECIONADO = None

def get_connection_string():
    params = urllib.parse.quote_plus(
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={DB_SERVER};"
        f"DATABASE={DB_NAME};"
        f"UID={DB_USER};"
        f"PWD={DB_PASS}"
    )
    return f"mssql+pyodbc:///?odbc_connect={params}"