# main.py
import sys
import os
import tkinter as tk            # Biblioteca de Interface Gráfica
from tkinter import filedialog  # Módulo de diálogo de arquivo (OpenDialog)

import config
import database as db
import import_produtos
import import_clientes
import import_financeiro

def selecionar_arquivo_gui():
    """Abre a janela nativa do Windows para escolher o arquivo"""
    print("Abrindo janela de seleção de arquivo...")
    
    # Cria uma janela raiz do Tkinter e a esconde (para não abrir uma janela vazia feia)
    root = tk.Tk()
    root.withdraw() 
    root.attributes('-topmost', True) # Garante que a janela apareça na frente de tudo

    caminho = filedialog.askopenfilename(
        title="Selecione a Planilha de Importação (MaxData)",
        filetypes=[
            ("Arquivos Excel", "*.xlsx *.xls"),
            ("Todos os arquivos", "*.*")
        ]
    )
    
    root.destroy() # Destrói a instância do Tkinter após o uso
    return caminho

def menu():
    caminho_atual = config.ARQUIVO_SELECIONADO if config.ARQUIVO_SELECIONADO else "Nenhum arquivo selecionado!"
    
    print("\n=== MAX IMPORT PYTHON (Analista Maxdata) ===")
    print(f"Banco Destino: {config.DB_NAME} em {config.DB_SERVER}")
    print(f"Arquivo Origem: {caminho_atual}")
    print("-" * 40)
    print("0. [TROCAR ARQUIVO] - Selecionar outra planilha")
    print("1. Importar PRODUTOS (Completo)")
    print("2. Importar CLIENTES")
    print("3. Importar FORNECEDORES")
    print("4. Importar FINANCEIRO")
    print("5. Sair")
    
    return input("Escolha uma opção: ")

def main():
    # Passo 1: Força a seleção do arquivo logo ao abrir o programa
    arquivo = selecionar_arquivo_gui()
    
    if not arquivo:
        print("Nenhum arquivo selecionado. O programa será encerrado.")
        return # Sai do programa se cancelar
        
    config.ARQUIVO_SELECIONADO = arquivo
    
    # Loop do Menu
    while True:
        opcao = menu()
        
        if opcao == '5':
            print("Saindo do sistema...")
            break

        # Opção extra para trocar de arquivo sem fechar o programa
        if opcao == '0':
            novo_arquivo = selecionar_arquivo_gui()
            if novo_arquivo:
                config.ARQUIVO_SELECIONADO = novo_arquivo
            continue # Volta para o menu com o novo arquivo

        # Preparação global de segurança
        try:
            db.toggle_constraints(False) 
            
            # Usa sempre a variável config.ARQUIVO_SELECIONADO que veio da GUI
            arquivo_atual = config.ARQUIVO_SELECIONADO

            if opcao == '1':
                limpar = input("Deseja limpar a base de PRODUTOS antes? (s/n): ").lower() == 's'
                import_produtos.executar_importacao(arquivo_atual, limpar_base=limpar)
                
            elif opcao == '2':
                limpar = input("Deseja limpar a base de CLIENTES antes? (s/n): ").lower() == 's'
                import_clientes.executar_importacao(arquivo_atual, is_fornecedor=False, limpar_base=limpar)
                
            elif opcao == '3':
                print("Importando Fornecedores...")
                import_clientes.executar_importacao(arquivo_atual, is_fornecedor=True, limpar_base=False)

            elif opcao == '4':
                print("--- ATENÇÃO ---")
                limpar = input("Tem certeza que deseja limpar a tabela FINANCEIRO antes? (s/n): ").lower() == 's'
                import_financeiro.executar_importacao(arquivo_atual, limpar_base=limpar)

            else:
                print("Opção inválida!")

        except Exception as e:
            print(f"\n❌ ERRO CRÍTICO: {e}")
            
        finally:
            db.toggle_constraints(True)
            print("\n✅ Processo finalizado.")

if __name__ == "__main__":
    main()