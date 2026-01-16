# import_clientes.py
import pandas as pd
import database as db
import utils
import datetime

def executar_importacao(caminho_excel, mapa_colunas=None, is_fornecedor=False, limpar_base=False):
    # Define o rótulo apenas para o log
    tipo_str = "Fornecedor" if is_fornecedor else "Cliente"
    print(f"--- Iniciando Importação de {tipo_str} (Modo Delphi) ---")

    try:
        # Lê o Excel como texto para evitar problemas de formatação
        df_origem = pd.read_excel(caminho_excel, sheet_name=0, dtype=str)
    except Exception as e:
        print(f"Erro leitura Excel: {e}")
        return

    # ---------------------------------------------------------
    # 1. LIMPEZA SEGURA (Igual ao Delphi)
    # ---------------------------------------------------------
    if limpar_base and not is_fornecedor:
        print("Limpando clientes antigos (PRESERVANDO USUÁRIO ADMIN E SISTEMA)...")
        # SQL exato do Delphi para não travar o sistema:
        # Mantém cliId 1 (Admin) e tipos 5/6 (Usuários internos)
        db.executar_comando("DELETE FROM cliente WHERE cliId <> 1 AND cliTipoCad <> 5 AND cliTipoCad <> 6")

    # ---------------------------------------------------------
    # 2. PREPARAÇÃO DOS DADOS
    # ---------------------------------------------------------
    df_cli = pd.DataFrame()

    def pegar_valor(campo_db, default=''):
        """Tenta pegar do mapa, se não existir, retorna default."""
        col_excel = None
        if mapa_colunas and campo_db in mapa_colunas:
            col_excel = mapa_colunas[campo_db]
        
        if col_excel and col_excel in df_origem.columns:
            return df_origem[col_excel].fillna(default)
        return pd.Series([default] * len(df_origem))

    # --- A. TRATAMENTO DO ID (cliId) ---
    
    # Se for CLIENTE, precisamos do ID para manter o histórico
    if not is_fornecedor:
        col_id = mapa_colunas.get('cliId') if mapa_colunas else None
        if col_id and col_id in df_origem.columns:
            # Converte para int
            df_cli['cliId'] = pd.to_numeric(df_origem[col_id], errors='coerce').fillna(0).astype(int)
            
            # Filtra IDs inválidos (zero ou negativo)
            df_cli = df_cli[df_cli['cliId'] > 0]
            
            # PROTEÇÃO CRÍTICA CONTRA TRAVAMENTO:
            # O banco já tem o ID 1 (Admin/Sistema). Se tentarmos importar o ID 1 do Excel,
            # vai dar erro de chave duplicada. Removemos o ID 1 da importação aqui.
            if 1 in df_cli['cliId'].values:
                print("AVISO: O Cliente ID 1 do Excel foi ignorado pois o ID 1 já é do ADMIN do sistema.")
                df_cli = df_cli[df_cli['cliId'] != 1]
        else:
            print("AVISO CRÍTICO: Importação de CLIENTE exige coluna 'cliId' mapeada!")
            return
    
    # --- B. CAMPOS PRINCIPAIS ---

    # cliTipo (Pessoa Física=0 / Jurídica=1 - Exemplo genérico, ajuste conforme necessidade)
    # Se for fornecedor, o Delphi tende a usar padrão 1, mas aqui tentamos pegar do Excel
    if is_fornecedor:
        df_cli['cliTipo'] = pegar_valor('cliTipo', '1').apply(lambda x: int(x) if str(x).isdigit() else 1)
    else:
        df_cli['cliTipo'] = pegar_valor('cliTipo', '0').apply(lambda x: int(x) if str(x).isdigit() else 0)

    df_cli['cliNome']     = pegar_valor('cliNome').apply(lambda x: utils.tratar_string(x, 50))
    df_cli['cliCpfCgc']   = pegar_valor('cliCpfCgc').apply(lambda x: utils.tratar_string(utils.remove_char(str(x)), 20))
    df_cli['cliRgInsc']   = pegar_valor('cliRgInsc').apply(lambda x: utils.tratar_string(x, 20))
    df_cli['cliFantasia'] = pegar_valor('cliFantasia').apply(lambda x: utils.tratar_string(x, 50))
    df_cli['cliEmail']    = pegar_valor('cliEmail').apply(lambda x: utils.tratar_string(x, 50))

    # --- C. ENDEREÇOS ---
    df_cli['cliFatEnd']       = pegar_valor('cliFatEnd').apply(lambda x: utils.tratar_string(x, 50))
    df_cli['cliFatBairro']    = pegar_valor('cliFatBairro').apply(lambda x: utils.tratar_string(x, 20))
    df_cli['cliFatEndNumero'] = pegar_valor('cliFatEndNumero').apply(lambda x: utils.tratar_string(x, 10))
    df_cli['cliFatCidade']    = pegar_valor('cliFatCidade').apply(lambda x: utils.tratar_string(x, 30))
    df_cli['cliFatUf']        = pegar_valor('cliFatUf').apply(lambda x: utils.tratar_string(x, 2))
    df_cli['cliFatCep']       = pegar_valor('cliFatCep').apply(lambda x: utils.tratar_string(utils.remove_char(str(x)), 9))
    
    # IBGE (Opcional, mas bom ter)
    df_cli['cliFatCidCodIBGE']= pegar_valor('cliFatCidCodIBGE', None) 

    # Cobrança (Copia dados se necessário ou pega do Excel)
    df_cli['cliCobEnd']       = pegar_valor('cliCobEnd').apply(lambda x: utils.tratar_string(x, 50))
    df_cli['cliCobBairro']    = pegar_valor('cliCobBairro').apply(lambda x: utils.tratar_string(x, 20))
    df_cli['cliCobEndNumero'] = pegar_valor('cliCobEndNumero').apply(lambda x: utils.tratar_string(x, 10))
    df_cli['cliCobCidade']    = pegar_valor('cliCobCidade').apply(lambda x: utils.tratar_string(x, 30))
    df_cli['cliCobUf']        = pegar_valor('cliCobUf').apply(lambda x: utils.tratar_string(x, 2))
    df_cli['cliCobCep']       = pegar_valor('cliCobCep').apply(lambda x: utils.tratar_string(utils.remove_char(str(x)), 9))
    df_cli['cliCobCidCodIBGE']= pegar_valor('cliCobCidCodIBGE', None)

    # --- D. FINANCEIRO E OBS ---
    df_cli['CliLimitCred']    = pegar_valor('CliLimitCred', '0').apply(lambda x: float(str(x).replace(',','.')) if x else 0)
    df_cli['zzz_CliObsVend']  = pegar_valor('zzz_CliObsVend').apply(lambda x: utils.tratar_string(x, 255))
    
    # --- E. CONTATO ---
    df_cli['CliFone']    = pegar_valor('CliFone').apply(lambda x: utils.tratar_string(utils.remove_char(str(x)), 10))
    df_cli['CliFax']     = pegar_valor('CliFax').apply(lambda x: utils.tratar_string(utils.remove_char(str(x)), 10))
    df_cli['cliCelular'] = pegar_valor('cliCelular').apply(lambda x: utils.tratar_string(utils.remove_char(str(x)), 10))

    df_cli['CliContNome1']  = pegar_valor('CliContNome1').apply(lambda x: utils.tratar_string(x, 50))
    df_cli['CliContDepto1'] = pegar_valor('CliContDepto1').apply(lambda x: utils.tratar_string(x, 20))
    df_cli['CliContFone1']  = pegar_valor('CliContFone1').apply(lambda x: utils.tratar_string(utils.remove_char(str(x)), 10))
    
    # Filiação (Pai/Mãe)
    df_cli['CliCadNomePai'] = pegar_valor('CliCadNomePai').apply(lambda x: utils.tratar_string(x, 50))
    df_cli['CliCadNomeMae'] = pegar_valor('CliCadNomeMae').apply(lambda x: utils.tratar_string(x, 50))

    # --- F. CAMPOS DE CONTROLE (REGRA DE NEGÓCIO IMPORTANTE) ---
    
    # Aqui implementamos a lógica exata solicitada:
    # 0 = CLIENTE
    # 1 = FORNECEDOR
    df_cli['cliTipoCad'] = 1 if is_fornecedor else 0
    
    df_cli['cliDatCad']  = datetime.datetime.now()

    # Validação final: Remove linhas sem Nome
    df_cli = df_cli[df_cli['cliNome'] != '']

    # ---------------------------------------------------------
    # 3. LÓGICA DE INSERÇÃO (IDENTITY_INSERT)
    # ---------------------------------------------------------
    
    manter_id_original = False

    if not is_fornecedor:
        # Se é Cliente, ativamos o IDENTITY_INSERT para usar os IDs do Excel
        manter_id_original = True
        
        # Verificação final de segurança
        if 'cliId' not in df_cli.columns:
            print("Erro: ID não encontrado no DataFrame.")
            return
    else:
        # Se é Fornecedor, removemos o cliId para o banco gerar (Auto Incremento)
        if 'cliId' in df_cli.columns:
            df_cli = df_cli.drop(columns=['cliId'])

    print(f"Inserindo {len(df_cli)} registros... (Manter ID original: {manter_id_original})")
    print(f"Tipo de Cadastro (cliTipoCad) definido como: {1 if is_fornecedor else 0}")
    
    # Chama a função de inserção no banco
    db.inserir_bulk(df_cli, 'cliente', manter_id=manter_id_original)
    
    print(f"--- Fim Importação {tipo_str} ---")