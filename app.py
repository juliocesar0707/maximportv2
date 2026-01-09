# app.py
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
import threading
import sys
import io
import os
import pandas as pd

# Módulos do Projeto
import config
import database as db
import import_produtos
import import_clientes
import import_financeiro
import ui_mapeamento

# --- REDIRECIONAMENTO DE LOG ---
class TextRedirector(io.StringIO):
    def __init__(self, text_widget):
        self.text_widget = text_widget
    def write(self, str):
        try:
            self.text_widget.after(0, self._append_text, str)
        except: pass
    def _append_text(self, str):
        try:
            self.text_widget.insert(END, str)
            self.text_widget.see(END)
        except: pass
    def flush(self): pass

# --- APP PRINCIPAL ---
class MaxImportApp(ttk.Window):
    def __init__(self):
        super().__init__(themename="cosmo")
        self.title("Max Import 2.0 - Maxdata Sistemas")
        self.geometry("980x750")
        
        if sys.platform.startswith("win") and os.path.exists("icone.ico"):
            self.iconbitmap("icone.ico")
        
        # Variáveis
        self.caminho_excel = ttk.StringVar()
        self.db_server = ttk.StringVar(value=config.DB_SERVER)
        self.db_name = ttk.StringVar(value=config.DB_NAME)
        self.progress_val = ttk.DoubleVar(value=0)

        self.criar_interface()

    def criar_interface(self):
        # Header
        header = ttk.Frame(self, padding=15, bootstyle="primary")
        header.pack(fill=X)
        ttk.Label(header, text="Max Import - Ferramenta de Migração", font=("Segoe UI", 18, "bold"), bootstyle="inverse-primary").pack(side=LEFT)

        # Container Principal
        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(fill=BOTH, expand=True)

        # --- ÁREA DE CONFIGURAÇÃO DE BANCO (TOPO) ---
        frame_db = ttk.Labelframe(main_frame, text="1. Conexão com Banco de Dados", padding=10)
        frame_db.pack(fill=X, pady=5)

        # Coluna 1: Servidor + Botão de Busca
        col_server = ttk.Frame(frame_db)
        col_server.pack(side=LEFT, fill=X, expand=True, padx=5)
        
        ttk.Label(col_server, text="Servidor:").pack(anchor=W)
        frm_input_srv = ttk.Frame(col_server)
        frm_input_srv.pack(fill=X)
        
        ttk.Entry(frm_input_srv, textvariable=self.db_server).pack(side=LEFT, fill=X, expand=True)
        ttk.Button(frm_input_srv, text="🔍", bootstyle="outline-secondary", command=self.listar_bancos_gui, width=3).pack(side=LEFT, padx=(2,0))
        
        # Coluna 2: Banco (Combobox)
        col_db = ttk.Frame(frame_db)
        col_db.pack(side=LEFT, fill=X, expand=True, padx=5)
        
        ttk.Label(col_db, text="Selecione o Banco:").pack(anchor=W)
        self.cbo_bancos = ttk.Combobox(col_db, textvariable=self.db_name, state="normal")
        self.cbo_bancos.pack(fill=X)

        # Botão Salvar
        ttk.Button(frame_db, text="Salvar Conexão", command=self.atualizar_conexao, bootstyle="secondary").pack(side=LEFT, padx=5, pady=(18,0))

        # --- ÁREA DE ARQUIVO ---
        frame_file = ttk.Labelframe(main_frame, text="2. Seleção de Arquivo (Excel)", padding=10)
        frame_file.pack(fill=X, pady=5)

        ttk.Entry(frame_file, textvariable=self.caminho_excel, state="readonly").pack(side=LEFT, fill=X, expand=True, padx=(0, 5))
        ttk.Button(frame_file, text="📂 Escolher Planilha", command=self.selecionar_arquivo, bootstyle="info").pack(side=RIGHT)

        # --- ÁREA DE AÇÕES (IMPORTAÇÃO E LIMPEZA) ---
        frame_acoes = ttk.Labelframe(main_frame, text="3. Ações e Ferramentas", padding=15)
        frame_acoes.pack(fill=X, pady=10)

        # Lado Esquerdo: Importações
        lbl_imp = ttk.Label(frame_acoes, text="IMPORTAÇÃO:", font=("Segoe UI", 9, "bold"), bootstyle="success")
        lbl_imp.grid(row=0, column=0, sticky=W, padx=5)

        btn_prod = ttk.Button(frame_acoes, text="📦 PRODUTOS", bootstyle="success", command=lambda: self.preparar_importacao(1))
        btn_prod.grid(row=1, column=0, padx=5, pady=5, sticky=EW)

        btn_cli = ttk.Button(frame_acoes, text="👥 CLIENTES", bootstyle="primary", command=lambda: self.preparar_importacao(2))
        btn_cli.grid(row=1, column=1, padx=5, pady=5, sticky=EW)

        btn_forn = ttk.Button(frame_acoes, text="🏭 FORNECEDORES", bootstyle="warning", command=lambda: self.preparar_importacao(3))
        btn_forn.grid(row=1, column=2, padx=5, pady=5, sticky=EW)

        btn_fin = ttk.Button(frame_acoes, text="💰 FINANCEIRO", bootstyle="info", command=lambda: self.preparar_importacao(4))
        btn_fin.grid(row=1, column=3, padx=5, pady=5, sticky=EW)

        # Separador Visual
        ttk.Separator(frame_acoes, orient=VERTICAL).grid(row=0, column=4, rowspan=2, sticky=NS, padx=20)

        # Lado Direito: Limpeza
        lbl_del = ttk.Label(frame_acoes, text="MANUTENÇÃO:", font=("Segoe UI", 9, "bold"), bootstyle="danger")
        lbl_del.grid(row=0, column=5, sticky=W, padx=5)

        btn_limpar = ttk.Button(frame_acoes, text="🗑️ LIMPAR DADOS...", bootstyle="danger", command=self.abrir_menu_limpeza)
        btn_limpar.grid(row=1, column=5, padx=5, pady=5, sticky=EW)

        # Ajuste de Grid
        frame_acoes.columnconfigure(0, weight=1)
        frame_acoes.columnconfigure(1, weight=1)
        frame_acoes.columnconfigure(2, weight=1)
        frame_acoes.columnconfigure(3, weight=1)
        frame_acoes.columnconfigure(5, weight=1)

        # --- LOG E PROGRESSO ---
        lbl_log = ttk.Label(main_frame, text="Log de Execução:", font=("Segoe UI", 9, "bold"))
        lbl_log.pack(anchor=W, pady=(10, 0))

        self.txt_log = ttk.Text(main_frame, height=12, font=("Consolas", 9))
        self.txt_log.pack(fill=BOTH, expand=True, pady=5)
        sys.stdout = TextRedirector(self.txt_log)

        self.barra_progresso = ttk.Progressbar(main_frame, variable=self.progress_val, bootstyle="striped-animated")
        self.barra_progresso.pack(fill=X, pady=5)
        
        ttk.Label(self, text="Maxdata Sistemas © 2025", font=("Segoe UI", 8), bootstyle="secondary").pack(side=BOTTOM, pady=5)

    # --- FUNÇÕES LÓGICAS ---

    def listar_bancos_gui(self):
        server = self.db_server.get()
        if not server:
            messagebox.showwarning("Aviso", "Digite o nome/IP do servidor primeiro.")
            return
            
        try:
            print(f"Buscando bancos no servidor {server}...")
            lista = db.listar_bancos_disponiveis(server)
            self.cbo_bancos['values'] = lista
            print(f"Encontrados {len(lista)} bancos.")
            if lista:
                self.cbo_bancos.current(0)
                self.cbo_bancos.event_generate("<<ComboboxSelected>>")
        except Exception as e:
            messagebox.showerror("Erro de Conexão", str(e))
            print(f"Erro: {e}")

    def atualizar_conexao(self):
        config.DB_SERVER = self.db_server.get()
        config.DB_NAME = self.db_name.get()
        print("--- Atualizando Conexão ---")
        db.reconectar()
        messagebox.showinfo("Conexão", f"Parâmetros atualizados!\nServidor: {config.DB_SERVER}\nBanco: {config.DB_NAME}")

    def selecionar_arquivo(self):
        arquivo = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if arquivo:
            self.caminho_excel.set(arquivo)

    def abrir_menu_limpeza(self):
        top = ttk.Toplevel(self)
        top.title("Zona de Perigo - Limpeza de Dados")
        top.geometry("400x350")
        
        ttk.Label(top, text="O que você deseja apagar?", font=("Segoe UI", 12, "bold"), bootstyle="danger").pack(pady=20)
        ttk.Label(top, text="Essa ação não pode ser desfeita.", font=("Segoe UI", 9)).pack()

        ttk.Button(top, text="Limpar APENAS Produtos", bootstyle="outline-danger", 
                   command=lambda: self.executar_limpeza_thread(1, top)).pack(fill=X, padx=30, pady=5)
        
        ttk.Button(top, text="Limpar APENAS Clientes", bootstyle="outline-danger", 
                   command=lambda: self.executar_limpeza_thread(2, top)).pack(fill=X, padx=30, pady=5)
        
        ttk.Button(top, text="Limpar APENAS Financeiro", bootstyle="outline-danger", 
                   command=lambda: self.executar_limpeza_thread(3, top)).pack(fill=X, padx=30, pady=5)

        ttk.Separator(top).pack(fill=X, padx=20, pady=10)

        ttk.Button(top, text="☢️ LIMPAR BANCO COMPLETO", bootstyle="danger", 
                   command=lambda: self.executar_limpeza_thread(99, top)).pack(fill=X, padx=30, pady=10)

    def executar_limpeza_thread(self, opcao, janela_popup):
        janela_popup.destroy()
        if not messagebox.askyesno("Confirmar Exclusão", "Tem certeza absoluta? Os dados serão perdidos para sempre."):
            return
        t = threading.Thread(target=self._limpeza_worker, args=(opcao,))
        t.start()

    def _limpeza_worker(self, opcao):
        self.barra_progresso.start(20)
        self.alternar_interface("disabled")
        try:
            db.toggle_constraints(False)
            if opcao == 1 or opcao == 99:
                print("Limpando Produtos e Estoque...")
                db.limpar_tabela('prolote', reset_identity=True)
                db.executar_comando("DELETE FROM produto_empresa") 
                db.executar_comando("DELETE FROM produto") 
                db.limpar_tabela('produtoUn', reset_identity=True)
            if opcao == 2 or opcao == 99:
                print("Limpando Clientes...")
                db.executar_comando("DELETE FROM cliente WHERE cliId > 1")
            if opcao == 3 or opcao == 99:
                print("Limpando Financeiro...")
                db.limpar_tabela('financeiro')
            messagebox.showinfo("Sucesso", "Limpeza concluída!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro na limpeza: {e}")
            print(f"Erro: {e}")
        finally:
            db.toggle_constraints(True)
            self.barra_progresso.stop()
            self.alternar_interface("normal")

    def preparar_importacao(self, opcao):
        arquivo = self.caminho_excel.get()
        if not arquivo:
            messagebox.showwarning("Aviso", "Selecione um arquivo Excel primeiro!")
            return
        if not os.path.exists(arquivo):
            messagebox.showerror("Erro", "Arquivo não encontrado!")
            return

        # Produtos (1) e Clientes (2) abrem mapa.
        if opcao in [1, 2]:
            try:
                df_header = pd.read_excel(arquivo, sheet_name=0, nrows=0)
                colunas = list(df_header.columns)
                tipo = "PRODUTO" if opcao == 1 else "CLIENTE"
                
                dialogo = ui_mapeamento.DialogoMapeamento(self, colunas, tipo_importacao=tipo)
                self.wait_window(dialogo)
                
                if dialogo.resultado:
                    t = threading.Thread(target=self.processar_thread, args=(opcao, dialogo.resultado))
                    t.start()
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao ler Excel: {e}")
        else:
            if messagebox.askyesno("Confirmar", "Iniciar importação direta?"):
                t = threading.Thread(target=self.processar_thread, args=(opcao, None))
                t.start()

    def processar_thread(self, opcao, mapa_colunas):
        self.barra_progresso.start(10)
        self.alternar_interface("disabled")
        try:
            db.reconectar()
            db.toggle_constraints(False)
            arquivo = self.caminho_excel.get()

            if opcao == 1:
                import_produtos.executar_importacao(arquivo, mapa_colunas, limpar_base=False)
            elif opcao == 2:
                # CORREÇÃO AQUI: Passando mapa_colunas
                import_clientes.executar_importacao(arquivo, mapa_colunas=mapa_colunas, is_fornecedor=False, limpar_base=False)
            elif opcao == 3:
                import_clientes.executar_importacao(arquivo, mapa_colunas=None, is_fornecedor=True, limpar_base=False)
            elif opcao == 4:
                import_financeiro.executar_importacao(arquivo, limpar_base=False)

            messagebox.showinfo("Sucesso", "Importação Finalizada!")

        except Exception as e:
            print(f"ERRO FATAL: {e}")
            messagebox.showerror("Erro", f"Falha na importação: {e}")
        finally:
            db.toggle_constraints(True)
            self.barra_progresso.stop()
            self.alternar_interface("normal")

    def alternar_interface(self, estado):
        pass 

if __name__ == "__main__":
    app = MaxImportApp()
    app.mainloop()