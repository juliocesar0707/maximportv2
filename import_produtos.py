# import_produtos.py
import pandas as pd
import numpy as np
import database as db
import utils
from sqlalchemy import text

def sincronizar_ncms(df_base):
    """
    1. Verifica os NCMs que estão no Excel.
    2. Insere os inexistentes na tabela 'proncm'.
    3. Retorna um dicionário { 'NCM_CODIGO': NCM_ID } para vincular no produto.
    """
    print("--- Sincronizando Tabela de NCMs (proncm) ---")
    
    # Pega lista única de NCMs da planilha (remove vazios)
    ncms_excel = df_base['zzz_proCodigoNcm'].dropna().unique()
    ncms_excel = [str(x).strip() for x in ncms_excel if str(x).strip() != '']
    
    if not ncms_excel:
        return {}

    engine = db.get_engine()
    
    # 1. Busca NCMs já existentes no banco para não duplicar
    lista_sql = str(tuple(ncms_excel)).replace(',)', ')') 
    if len(ncms_excel) == 1: 
        lista_sql = f"('{ncms_excel[0]}')"

    sql_busca = f"SELECT NCMcodigoNCM, ncmid FROM proncm WHERE NCMcodigoNCM IN {lista_sql}"
    
    try:
        df_existentes = pd.read_sql(sql_busca, engine)
        ncms_banco = set(df_existentes['NCMcodigoNCM'].astype(str).str.strip())
    except Exception as e:
        print(f"Erro ao buscar NCMs: {e}")
        ncms_banco = set()

    # 2. Identifica quais são novos
    ncms_novos = [n for n in ncms_excel if n not in ncms_banco]

    # 3. Insere os novos (se houver)
    if ncms_novos:
        print(f"Cadastrando {len(ncms_novos)} novos NCMs...")
        with engine.begin() as conn:
            for codigo in ncms_novos:
                try:
                    # Inserindo apenas o código. 
                    sql_insert = "INSERT INTO proncm (NCMcodigoNCM) VALUES (:cod)"
                    conn.execute(text(sql_insert), {"cod": codigo})
                except Exception as e:
                    print(f"Erro ao inserir NCM {codigo}: {e}")

    # 4. Busca tudo novamente para gerar o dicionário de IDs atualizado
    try:
        df_final = pd.read_sql(sql_busca, engine)
        # Cria dicionário: Chave = Código NCM, Valor = ID (ncmid)
        mapa_ids = dict(zip(df_final['NCMcodigoNCM'].astype(str).str.strip(), df_final['ncmid']))
        return mapa_ids
    except Exception as e:
        print(f"Erro ao mapear IDs de NCM: {e}")
        return {}

def executar_importacao(caminho_excel, mapa_colunas, limpar_base=False):
    print("--- Iniciando Importação (Produto + Empresa + Fiscal) ---")
    
    try:
        df_origem = pd.read_excel(caminho_excel, sheet_name=0, dtype=str)
    except Exception as e:
        print(f"Erro ao ler Excel: {e}")
        return

    # 1. LIMPEZA
    if limpar_base:
        print("Limpando tabelas...")
        try:
            db.limpar_tabela('prolote', reset_identity=True)
            db.executar_comando("DELETE FROM produto_empresa")
            db.executar_comando("DELETE FROM produto WHERE proId > 1") 
            db.executar_comando("DBCC CHECKIDENT ('produto', RESEED, 1)")
            db.limpar_tabela('produtoUn', reset_identity=True)
        except Exception as e:
            print(f"Erro limpeza: {e}")

    # 2. PREPARAÇÃO DOS DADOS
    def pegar_valor(campo_db, funcao_tratamento=None, valor_padrao=''):
        col_excel = mapa_colunas.get(campo_db)
        if col_excel and col_excel in df_origem.columns:
            serie = df_origem[col_excel]
            if funcao_tratamento:
                return serie.apply(lambda x: funcao_tratamento(x) if pd.notnull(x) else valor_padrao)
            return serie.fillna(valor_padrao)
        return pd.Series([valor_padrao] * len(df_origem))

    df_base = pd.DataFrame()
    
    # Campos Universais
    df_base['proDescricao'] = pegar_valor('proDescricao', lambda x: utils.tratar_string(x, 50))
    df_base['zzz_proCodigo'] = pegar_valor('zzz_proCodigo', lambda x: utils.tratar_string(x, 20))
    
    # NCM: Lemos o CÓDIGO do Excel para uma coluna temporária
    df_base['zzz_proCodigoNcm'] = pegar_valor('zzz_proCodigoNcm', lambda x: utils.tratar_string(utils.remove_char(x), 8))

    # --- PROCESSAMENTO NCM ---
    # Sincroniza com a tabela 'proncm' e pega os IDs
    mapa_ncm_ids = sincronizar_ncms(df_base)
    
    # Cria a coluna de ID (FK) mapeando o código. Se não achar, fica 0.
    df_base['proncmid'] = df_base['zzz_proCodigoNcm'].map(mapa_ncm_ids).fillna(0).astype(int)

    # Campos Comerciais
    df_base['proUn'] = pegar_valor('proUn', lambda x: utils.tratar_string(x, 2), valor_padrao='UN')
    df_base['proCusto'] = pegar_valor('zzz_proCusto', utils.tratar_moeda, valor_padrao=0)
    df_base['proVenda'] = pegar_valor('zzz_proVenda', utils.tratar_moeda, valor_padrao=0)
    df_base['proEstoqueAtual'] = pegar_valor('proEstoqueAtual', utils.tratar_moeda, valor_padrao=0)
    df_base['proEstoqueMin'] = pegar_valor('zzz_proEstoqueMin', utils.tratar_moeda, valor_padrao=0)
    
    # --- CST e CSOSN (Nomes Corrigidos) ---
    df_base['proCodcst2'] = pegar_valor('proCodcst2', lambda x: utils.tratar_string(x, 3), valor_padrao='')
    df_base['proCodCSOSN'] = pegar_valor('proCodCSOSN', lambda x: utils.tratar_string(x, 4), valor_padrao='')

    df_base['proCodigoEmpresa'] = df_base['zzz_proCodigo'] 

    # 3. SEPARAÇÃO (FIXO vs AUTOMÁTICO)
    col_id_excel = mapa_colunas.get('proId')
    ids_originais = df_origem[col_id_excel] if (col_id_excel and col_id_excel in df_origem.columns) else pd.Series([np.nan] * len(df_origem))
    ids_numericos = pd.to_numeric(ids_originais, errors='coerce')
    mask_numerico = ids_numericos.notnull()
    
    # LISTA DE COLUNAS PARA INSERT NA TABELA produto_empresa
    # Agora incluindo proCodcst2 e proCodCSOSN
    cols_empresa = [
        'proId', 'proUn', 'proCusto', 'proVenda', 'proEstoqueAtual', 
        'proEstoqueMin', 'proCodigoEmpresa', 'proCodcst2', 'proCodCSOSN'
    ]

    # --- GRUPO 1: IDs FIXOS ---
    df_fixo = df_base[mask_numerico].copy()
    if not df_fixo.empty:
        df_fixo['proId'] = ids_numericos[mask_numerico].astype(int)
        
        # Ajuste: Inserimos 'proncmid'
        df_prod_fixo = df_fixo[['proId', 'proDescricao', 'zzz_proCodigo', 'proncmid']].copy()
        
        print(f"Inserindo {len(df_prod_fixo)} produtos FIXOS...")
        db.inserir_bulk(df_prod_fixo, 'produto', manter_id=True)
        
        # Dados Comerciais
        df_emp_fixo = df_fixo[cols_empresa].copy()
        df_emp_fixo.rename(columns={'proCodigoEmpresa': 'proCodigo'}, inplace=True)
        df_emp_fixo['empId'] = 1 
        db.inserir_bulk(df_emp_fixo, 'produto_empresa', manter_id=False)

    # --- GRUPO 2: IDs AUTOMÁTICOS ---
    df_auto = df_base[~mask_numerico].copy()
    if not df_auto.empty:
        df_auto['zzz_proCodigo'] = df_auto['zzz_proCodigo'].replace('', 'AUTO_' + df_auto.index.astype(str))
        
        # Ajuste: Inserimos 'proncmid'
        df_prod_auto = df_auto[['proDescricao', 'zzz_proCodigo', 'proncmid']].copy()
        
        print(f"Inserindo {len(df_prod_auto)} produtos AUTOMÁTICOS...")
        db.inserir_bulk(df_prod_auto, 'produto', manter_id=False)
        
        print("Sincronizando IDs gerados...")
        lista_refs = tuple(df_auto['zzz_proCodigo'].unique())
        if len(lista_refs) > 0:
            query = f"SELECT zzz_proCodigo, proId FROM produto WHERE zzz_proCodigo IN {lista_refs}"
            try:
                df_ids_gerados = pd.read_sql(query, db.get_engine())
                df_auto_final = pd.merge(df_auto, df_ids_gerados, on='zzz_proCodigo', how='inner')
                
                df_emp_auto = df_auto_final[cols_empresa].copy()
                df_emp_auto.rename(columns={'proCodigoEmpresa': 'proCodigo'}, inplace=True)
                df_emp_auto['empId'] = 1
                
                db.inserir_bulk(df_emp_auto, 'produto_empresa', manter_id=False)
            except Exception as e:
                print(f"Erro ao sincronizar IDs automáticos: {e}")

    # 4. AUXILIARES
    print("Processando Unidades...")
    sql_unidades = """
    INSERT INTO produtoUn (unpUn, unpDescricao)
    SELECT DISTINCT proUn, proUn FROM produto_empresa 
    WHERE proUn IS NOT NULL AND proUn <> '' 
    AND proUn NOT IN (SELECT unpUn FROM produtoUn)
    """
    db.executar_comando(sql_unidades)

    print("--- Fim Importação ---")