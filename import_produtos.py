# import_produtos.py
import pandas as pd
import numpy as np
import database as db
import utils

def executar_importacao(caminho_excel, mapa_colunas, limpar_base=False):
    print("--- Iniciando Importação Normalizada (Produto + Empresa) ---")
    
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
    # Função auxiliar para ler colunas do Excel
    def pegar_valor(campo_db, funcao_tratamento=None, valor_padrao=''):
        col_excel = mapa_colunas.get(campo_db)
        if col_excel and col_excel in df_origem.columns:
            serie = df_origem[col_excel]
            if funcao_tratamento:
                return serie.apply(lambda x: funcao_tratamento(x) if pd.notnull(x) else valor_padrao)
            return serie.fillna(valor_padrao)
        return pd.Series([valor_padrao] * len(df_origem))

    # --- MONTAGEM DO DATAFRAME UNIVERSAL ---
    df_base = pd.DataFrame()
    
    # Campos Universais (Vão para tabela PRODUTO)
    df_base['proDescricao'] = pegar_valor('proDescricao', lambda x: utils.tratar_string(x, 50))
    # Usamos zzz_proCodigo na tabela produto apenas como "ponte" para recuperar o ID depois
    df_base['zzz_proCodigo'] = pegar_valor('zzz_proCodigo', lambda x: utils.tratar_string(x, 20))
    df_base['zzz_proCodigoNcm'] = pegar_valor('zzz_proCodigoNcm', lambda x: utils.tratar_string(utils.remove_char(x), 8))

    # Campos Comerciais (Vão para tabela PRODUTO_EMPRESA)
    df_base['proUn'] = pegar_valor('proUn', lambda x: utils.tratar_string(x, 2), valor_padrao='UN')
    df_base['proCusto'] = pegar_valor('zzz_proCusto', utils.tratar_moeda, valor_padrao=0)
    df_base['proVenda'] = pegar_valor('zzz_proVenda', utils.tratar_moeda, valor_padrao=0)
    df_base['proEstoqueAtual'] = pegar_valor('proEstoqueAtual', utils.tratar_moeda, valor_padrao=0)
    df_base['proEstoqueMin'] = pegar_valor('zzz_proEstoqueMin', utils.tratar_moeda, valor_padrao=0)
    
    # A Referência também vai para produto_empresa (coluna proCodigo)
    df_base['proCodigoEmpresa'] = df_base['zzz_proCodigo'] 

    # 3. SEPARAÇÃO (FIXO vs AUTOMÁTICO)
    col_id_excel = mapa_colunas.get('proId')
    if col_id_excel and col_id_excel in df_origem.columns:
        ids_originais = df_origem[col_id_excel]
    else:
        ids_originais = pd.Series([np.nan] * len(df_origem))

    ids_numericos = pd.to_numeric(ids_originais, errors='coerce')
    mask_numerico = ids_numericos.notnull()

    # --- PROCESSAMENTO GRUPO 1: IDs FIXOS (Numéricos) ---
    df_fixo = df_base[mask_numerico].copy()
    if not df_fixo.empty:
        df_fixo['proId'] = ids_numericos[mask_numerico].astype(int)
        
        # Etapa 1.1: Inserir em PRODUTO (Dados Básicos)
        df_prod_fixo = df_fixo[['proId', 'proDescricao', 'zzz_proCodigo', 'zzz_proCodigoNcm']].copy()
        print(f"Inserindo {len(df_prod_fixo)} produtos FIXOS na tabela PRODUTO...")
        db.inserir_bulk(df_prod_fixo, 'produto', manter_id=True)
        
        # Etapa 1.2: Inserir em PRODUTO_EMPRESA (Dados Comerciais)
        df_emp_fixo = df_fixo[['proId', 'proUn', 'proCusto', 'proVenda', 'proEstoqueAtual', 'proEstoqueMin', 'proCodigoEmpresa']].copy()
        df_emp_fixo.rename(columns={'proCodigoEmpresa': 'proCodigo'}, inplace=True)
        df_emp_fixo['empId'] = 1 # Empresa Fixa
        
        print(f"Inserindo dados comerciais (Preço/Estoque/UN) na tabela PRODUTO_EMPRESA...")
        db.inserir_bulk(df_emp_fixo, 'produto_empresa', manter_id=False)

    # --- PROCESSAMENTO GRUPO 2: IDs AUTOMÁTICOS (Alfanuméricos/Vazios) ---
    df_auto = df_base[~mask_numerico].copy()
    if not df_auto.empty:
        # Prepara Referência para servir de chave de busca (Preenche vazios com 'SEM_REF')
        # Precisamos garantir que zzz_proCodigo tenha algo único para recuperarmos o ID depois
        df_auto['zzz_proCodigo'] = df_auto['zzz_proCodigo'].replace('', 'AUTO_' + df_auto.index.astype(str))
        
        # Etapa 2.1: Inserir em PRODUTO (Sem ID, deixa o banco criar)
        df_prod_auto = df_auto[['proDescricao', 'zzz_proCodigo', 'zzz_proCodigoNcm']].copy()
        
        print(f"Inserindo {len(df_prod_auto)} produtos AUTOMÁTICOS na tabela PRODUTO...")
        db.inserir_bulk(df_prod_auto, 'produto', manter_id=False)
        
        # Etapa 2.2: RECUPERAR OS IDs GERADOS PELO BANCO
        # Usamos a referência (zzz_proCodigo) para achar o proId que acabou de ser criado
        print("Sincronizando IDs gerados pelo banco...")
        
        # Pega a lista de referências que acabamos de inserir
        lista_refs = tuple(df_auto['zzz_proCodigo'].unique())
        
        if len(lista_refs) > 0:
            # Busca no banco: Quem é o ID da referência 'X'?
            # Nota: Se houver muitas refs, fazemos em chunks, mas aqui simplificado:
            query = f"SELECT zzz_proCodigo, proId FROM produto WHERE zzz_proCodigo IN {lista_refs}"
            
            # Se a lista for muito grande, o python pode reclamar da query. 
            # Em produção ideal, faríamos batch, mas para < 5000 itens vai tranquilo.
            try:
                df_ids_gerados = pd.read_sql(query, db.get_engine())
                
                # Faz o MERGE (VLOOKUP) para trazer o proId de volta pro DataFrame
                df_auto_final = pd.merge(df_auto, df_ids_gerados, on='zzz_proCodigo', how='inner')
                
                # Etapa 2.3: Inserir em PRODUTO_EMPRESA com os IDs certos
                df_emp_auto = df_auto_final[['proId', 'proUn', 'proCusto', 'proVenda', 'proEstoqueAtual', 'proEstoqueMin', 'proCodigoEmpresa']].copy()
                df_emp_auto.rename(columns={'proCodigoEmpresa': 'proCodigo'}, inplace=True)
                df_emp_auto['empId'] = 1
                
                print(f"Inserindo {len(df_emp_auto)} vínculos comerciais na tabela PRODUTO_EMPRESA...")
                db.inserir_bulk(df_emp_auto, 'produto_empresa', manter_id=False)
                
            except Exception as e:
                print(f"Erro ao sincronizar IDs automáticos: {e}")

    # 4. ROTINAS AUXILIARES
    print("Processando Unidades (Tabela produtoUn)...")
    # Agora a unidade vem da tabela produto_empresa, não mais de produto
    sql_unidades = """
    INSERT INTO produtoUn (unpUn, unpDescricao)
    SELECT DISTINCT proUn, proUn FROM produto_empresa 
    WHERE proUn IS NOT NULL AND proUn <> '' 
    AND proUn NOT IN (SELECT unpUn FROM produtoUn)
    """
    db.executar_comando(sql_unidades)

    print("--- Fim Importação ---")