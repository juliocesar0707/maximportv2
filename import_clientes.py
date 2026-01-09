# import_clientes.py
import pandas as pd
import database as db
import utils
import datetime

def executar_importacao(caminho_excel, is_fornecedor=False, limpar_base=False):
    tipo_str = "Fornecedor" if is_fornecedor else "Cliente"
    print(f"--- Iniciando Importação de {tipo_str} ---")

    try:
        df_origem = pd.read_excel(caminho_excel, sheet_name=0, dtype=str)
    except Exception as e:
        print(f"Erro leitura Excel: {e}")
        return

    # Limpeza (Cuidado com o ID 1 que costuma ser Consumidor Final)
    if limpar_base and not is_fornecedor:
        db.executar_comando("DELETE FROM cliente WHERE cliId > 1 AND cliTipoCad NOT IN (5,6)")

    df_cli = pd.DataFrame()

    # Mapeamento
    if not is_fornecedor:
        df_cli['cliId'] = df_origem['id'].fillna(0).astype(int)
    
    # Lógica do checkbox chkFornecedor do Delphi
    # 0 = Cliente, 1 = Fornecedor (Baseado no Delphi)
    tipo_cad = 1 if is_fornecedor else 0 
    
    df_cli['cliTipoCad'] = tipo_cad
    df_cli['cliNome'] = df_origem['nome'].apply(lambda x: utils.tratar_string(x, 50))
    df_cli['cliFantasia'] = df_origem['fantasia'].apply(lambda x: utils.tratar_string(x, 50))
    df_cli['cliCpfCgc'] = df_origem['cpf_cnpj'].apply(lambda x: utils.tratar_string(utils.remove_char(x), 20))
    df_cli['cliDatCad'] = datetime.datetime.now()
    
    # Endereço
    df_cli['cliFatEnd'] = df_origem['endereco'].apply(lambda x: utils.tratar_string(x, 50))
    df_cli['cliFatBairro'] = df_origem['bairro'].apply(lambda x: utils.tratar_string(x, 20))
    df_cli['cliFatCidade'] = df_origem['cidade'].apply(lambda x: utils.tratar_string(x, 30))
    df_cli['cliFatUf'] = df_origem['uf'].apply(lambda x: utils.tratar_string(x, 2))
    df_cli['cliEmail'] = df_origem['email'].apply(lambda x: utils.tratar_string(x, 50))
    df_cli['cliFone'] = df_origem['telefone'].apply(lambda x: utils.tratar_string(utils.remove_char(x), 15))

    # Inserção
    # Se for fornecedor, geralmente não forçamos o ID (Identity Insert OFF), 
    # mas o seu código Delphi mantinha IDs para clientes.
    manter_id = not is_fornecedor
    
    db.inserir_bulk(df_cli, 'cliente', manter_id=manter_id)
    print(f"--- Fim {tipo_str} ---")