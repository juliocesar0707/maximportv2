# import_financeiro.py
import pandas as pd
import database as db
import utils
import datetime

def executar_importacao(caminho_excel, limpar_base=False):
    print("--- Iniciando Importação do FINANCEIRO ---")

    # 1. LEITURA DO EXCEL
    # Ajuste 'sheet_name' se os dados financeiros estiverem em outra aba
    try:
        df_origem = pd.read_excel(caminho_excel, sheet_name=0, dtype=str)
    except Exception as e:
        print(f"Erro ao ler Excel: {e}")
        return

    # 2. LIMPEZA DA BASE (Opcional)
    if limpar_base:
        # CUIDADO: Isso apaga todo o histórico financeiro
        print("Limpando tabela FINANCEIRO...")
        db.executar_comando("DELETE FROM financeiro")
        # Se tiver tabela filha (ex: financeiro_baixa), limpar aqui também

    df_fin = pd.DataFrame()

    # 3. MAPEAMENTO E TRATAMENTO
    # Adapte as chaves do df_origem['...'] para os nomes do SEU Excel

    # Link com Cliente/Fornecedor (Essencial)
    df_fin['pgtClienteId'] = df_origem['id_cliente'].fillna(0).astype(int)
    
    # Valores Monetários
    df_fin['pgtValor'] = df_origem['valor_original'].apply(utils.tratar_moeda)
    df_fin['pgtValorJuros'] = df_origem['juros'].apply(utils.tratar_moeda)
    
    # Tratamento de Datas (Crucial para o Financeiro)
    # dayfirst=True garante que 01/02 seja 1º de Fev, não 2 de Jan
    df_fin['pgtData'] = pd.to_datetime(df_origem['data_emissao'], dayfirst=True, errors='coerce').fillna(datetime.datetime.now())
    
    # Nota: No seu Delphi estava escrito 'pgtVecmto' (possível erro de digitação no legado)
    # Verifique no SQL se a coluna é 'pgtVencimento' ou 'pgtVecmto'. Mantive a do Delphi.
    df_fin['pgtVecmto'] = pd.to_datetime(df_origem['data_vencimento'], dayfirst=True, errors='coerce')
    df_fin['pgtDataQuitou'] = pd.to_datetime(df_origem['data_pagamento'], dayfirst=True, errors='coerce')

    # Identificação do Documento
    df_fin['pgtNumDoc'] = df_origem['numero_doc'].apply(lambda x: utils.tratar_string(x, 20))
    df_fin['pgtNossoNumero'] = df_origem['nosso_numero'].apply(lambda x: utils.tratar_string(x, 20))
    df_fin['pgtObs'] = df_origem['obs'].apply(lambda x: utils.tratar_string(x, 100)) # Ajuste tamanho se precisar

    # Tipos e Status (Regras de Negócio)
    
    # Tipo Conta: 'R' = Receber, 'P' = Pagar
    # Se não tiver no Excel, assume 'R' (padrão do seu Delphi)
    if 'tipo_conta' in df_origem.columns:
        df_fin['pgtTipoConta'] = df_origem['tipo_conta'].apply(lambda x: utils.tratar_string(x, 1).upper())
    else:
        df_fin['pgtTipoConta'] = 'R'

    # Pago: 'S' = Sim, 'N' = Não
    if 'pago' in df_origem.columns:
        df_fin['pgtPago'] = df_origem['pago'].apply(lambda x: utils.tratar_string(x, 1).upper())
    else:
        # Lógica inteligente: Se tem data de quitação, está pago
        df_fin['pgtPago'] = df_fin['pgtDataQuitou'].apply(lambda x: 'S' if pd.notnull(x) else 'N')

    # Dados Bancários (Opcional)
    df_fin['pgtBanco'] = df_origem['banco'].apply(lambda x: utils.tratar_string(x, 10)) if 'banco' in df_origem else ''
    df_fin['pgtAgencia'] = df_origem['agencia'].apply(lambda x: utils.tratar_string(x, 10)) if 'agencia' in df_origem else ''
    df_fin['pgtContaC'] = df_origem['conta'].apply(lambda x: utils.tratar_string(x, 15)) if 'conta' in df_origem else ''

    # Campos Padrão do Legado
    df_fin['empId'] = 1
    # Tipos de pagamento (Vista/Prazo) - Defaults do Delphi
    df_fin['pgtTipoVista'] = 0 
    df_fin['pgtTipoPrazo'] = 3 # Ex: 3 = Boleto (conforme seu btnInfo do Delphi)

    # 4. INSERÇÃO NO BANCO
    # Financeiro geralmente tem autoincremento no ID principal (ex: pgtId), 
    # então NÃO enviamos o ID, deixamos o SQL Server gerar.
    print(f"Inserindo {len(df_fin)} registros financeiros...")
    
    # Removendo linhas onde Cliente é 0 ou Data de Vencimento é NaT (Not a Time) para evitar erro
    df_fin = df_fin[df_fin['pgtClienteId'] > 0]
    df_fin = df_fin[df_fin['pgtVecmto'].notnull()]

    db.inserir_bulk(df_fin, 'financeiro', manter_id=False)

    print("--- Fim Importação Financeira ---")