# import_clientes.py
import pandas as pd
import database as db
import utils
import datetime

def executar_importacao(caminho_excel, mapa_colunas=None, is_fornecedor=False, limpar_base=False):
    tipo_str = "Fornecedor" if is_fornecedor else "Cliente"
    print(f"--- Iniciando Importação de {tipo_str} ---")

    try:
        # Lê o Excel como texto para evitar problemas de formatação
        df_origem = pd.read_excel(caminho_excel, sheet_name=0, dtype=str)
    except Exception as e:
        print(f"Erro leitura Excel: {e}")
        return

    # 1. LIMPEZA
    if limpar_base and not is_fornecedor:
        print("Limpando clientes antigos...")
        db.executar_comando("DELETE FROM cliente WHERE cliId > 1 AND cliTipoCad NOT IN (5,6)")

    # 2. PREPARAÇÃO DOS DADOS
    df_cli = pd.DataFrame()

    # Função auxiliar inteligente:
    # Se tiver mapa, usa o mapa. Se não (Fornecedor direto), tenta achar pelo nome padrão ou retorna vazio.
    def pegar_valor(campo_db, default=''):
        col_excel = None
        
        # 1. Tenta pegar pelo Mapeamento (Prioridade)
        if mapa_colunas and campo_db in mapa_colunas:
            col_excel = mapa_colunas[campo_db]
        
        # 2. Se achou a coluna e ela existe no Excel
        if col_excel and col_excel in df_origem.columns:
            serie = df_origem[col_excel].fillna(default)
            return serie
        
        # 3. Retorna coluna com valor padrão se não achar nada
        return pd.Series([default] * len(df_origem))

    # --- MAPEAMENTO DOS CAMPOS ---

    # 1. ID (cliId)
    # Se for Cliente e mapeou o ID, usa ele.
    if not is_fornecedor:
        col_id = mapa_colunas.get('cliId') if mapa_colunas else None
        if col_id and col_id in df_origem.columns:
            df_cli['cliId'] = pd.to_numeric(df_origem[col_id], errors='coerce').fillna(0).astype(int)
        else:
            # Se não mapeou, não cria a coluna (o banco gera auto-incremento)
            pass 
    
    # 2. Dados Principais (Mapeados na Tela)
    df_cli['cliNome']     = pegar_valor('cliNome').apply(lambda x: utils.tratar_string(x, 50))
    df_cli['cliFantasia'] = pegar_valor('cliFantasia').apply(lambda x: utils.tratar_string(x, 50))
    df_cli['cliCpfCgc']   = pegar_valor('cliCpfCgc').apply(lambda x: utils.tratar_string(utils.remove_char(x), 20))
    df_cli['cliFone']     = pegar_valor('cliFone').apply(lambda x: utils.tratar_string(utils.remove_char(x), 15))
    df_cli['cliEmail']    = pegar_valor('cliEmail').apply(lambda x: utils.tratar_string(x, 50))
    
    # 3. Endereço (Mapeados na Tela)
    df_cli['cliFatEnd']    = pegar_valor('cliFatEnd').apply(lambda x: utils.tratar_string(x, 50))
    df_cli['cliFatBairro'] = pegar_valor('cliFatBairro').apply(lambda x: utils.tratar_string(x, 20))
    df_cli['cliFatCidade'] = pegar_valor('cliFatCidade').apply(lambda x: utils.tratar_string(x, 30))
    df_cli['cliFatUf']     = pegar_valor('cliFatUf').apply(lambda x: utils.tratar_string(x, 2))

    # 4. Campos "Invisíveis" (Necessários para o Delphi não dar erro)
    # Preenchemos com vazio/zero pois não foram mapeados na tela
    df_cli['cliRgInsc']       = ''
    df_cli['cliFatEndNumero'] = '' 
    df_cli['cliFatCep']       = ''
    
    # Endereço de Cobrança (Cópia vazia)
    df_cli['cliCobEnd']       = ''
    df_cli['cliCobBairro']    = ''
    df_cli['cliCobEndNumero'] = ''
    df_cli['cliCobCidade']    = ''
    df_cli['cliCobUf']        = ''
    df_cli['cliCobCep']       = ''
    
    # Outros
    df_cli['CliLimitCred']    = 0
    df_cli['zzz_CliObsVend']  = ''
    df_cli['CliFax']          = ''
    df_cli['cliCelular']      = ''
    
    # Contatos extras
    df_cli['CliContNome1']  = ''
    df_cli['CliContDepto1'] = ''
    df_cli['CliContFone1']  = ''
    df_cli['CliCadNomePai'] = ''
    df_cli['CliCadNomeMae'] = ''

    # 5. Campos de Controle
    df_cli['cliDatCad']  = datetime.datetime.now()
    df_cli['cliTipoCad'] = 1 if is_fornecedor else 0
    df_cli['cliTipo']    = 1 if is_fornecedor else 0

    # Remove linhas onde o nome é vazio (evita sujeira)
    df_cli = df_cli[df_cli['cliNome'] != '']

    # 3. INSERÇÃO
    manter_id = False
    
    # Só ativamos IDENTITY_INSERT se a coluna cliId existir no DataFrame e tiver dados
    if 'cliId' in df_cli.columns:
        if (df_cli['cliId'] > 10).any():
            manter_id = True
        else:
            # Se a coluna existe mas só tem zeros, remove ela pra deixar o banco gerar
            df_cli = df_cli.drop(columns=['cliId'])

    print(f"Inserindo {len(df_cli)} registros... (Manter ID original: {manter_id})")
    
    # Grava no banco
    db.inserir_bulk(df_cli, 'cliente', manter_id=manter_id)
    print(f"--- Fim Importação {tipo_str} ---")