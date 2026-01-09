# utils.py
import re
import pandas as pd

def remove_char(texto):
    """Remove tudo que não é número (para CNPJ, CPF, Telefone)"""
    if pd.isna(texto): return ''
    # Converte para string primeiro para evitar erro em números
    return re.sub(r'[^0-9]', '', str(texto))

def tratar_string(texto, tamanho_max):
    """Corta a string para caber no campo do banco e trata Nulos"""
    if pd.isna(texto): return ''
    s = str(texto).strip()
    return s[:tamanho_max]

def tratar_moeda(valor):
    """
    Converte valores monetários para float de forma inteligente.
    Resolve problemas de Ponto vs Vírgula.
    """
    if pd.isna(valor) or valor == '':
        return 0.0
    
    # Se já for número (float/int), retorna direto
    if isinstance(valor, (int, float)):
        return float(valor)
        
    s = str(valor).strip()
    
    # Remove símbolos de moeda e espaços extras
    s = s.replace('R$', '').replace('r$', '').strip()
    
    # --- LÓGICA DE DETECÇÃO DE FORMATO ---
    
    # Caso 1: Formato Brasileiro Completo (ex: 1.500,50)
    # Tem ponto E vírgula. O ponto é milhar, a vírgula é decimal.
    if '.' in s and ',' in s:
        s = s.replace('.', '') # Remove milhar
        s = s.replace(',', '.') # Transforma vírgula em ponto para o Python entender
    
    # Caso 2: Apenas Vírgula (ex: 15,83 ou 1500,00)
    # Assume padrão brasileiro: troca vírgula por ponto.
    elif ',' in s:
        s = s.replace(',', '.')
        
    # Caso 3: Apenas Ponto (ex: 15.830 ou 1000.00) -> O SEU CASO
    # O código antigo removia o ponto achando que era milhar (15830).
    # Agora: Se tiver apenas UM ponto, assumimos que é decimal (15.83).
    elif '.' in s:
        # Se tiver múltiplos pontos (ex: 1.000.000), aí sim é milhar
        if s.count('.') > 1:
            s = s.replace('.', '')
        else:
            # Mantém o ponto, pois é decimal (15.830 vira 15.83)
            pass 
            
    try:
        return float(s)
    except Exception as e:
        # Se falhar, imprime no log para ajudar a debugar
        print(f"Erro ao converter valor '{valor}': {e}")
        return 0.0