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

# --- REDIRECIONAMENTO DE LOG (Mantido Intacto) ---
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
        # MUDANÇA VISUAL: Tema 'superhero' (Dark Moderno) ou 'flatly' (Clean Light)
        # Vamos usar 'superhero' para o visual "Super Bonito" e Tech.
        super().__init__(themename="superhero") 
        self.title("Max Import 2.0 - Suite de Migração")
        self.geometry("1024x800")
        self.place_window_center()
        
        if sys.platform.startswith("win") and os.path.exists("icone.ico"):
            self.iconbitmap("icone.ico")
        
        # Variáveis
        self.caminho_excel = ttk.StringVar()
        self.db_server = ttk.StringVar(value=config.DB_SERVER)
        self.db_name = ttk.StringVar(value=config.DB_NAME)
        self.progress_val = ttk.DoubleVar(value=0)

        # Estilos Customizados
        style = ttk.Style()
        style.configure('Big.TButton', font=('Segoe UI', 11, 'bold'))
        style.configure('Card.TFrame', background=style.colors.bg)

        self.criar_interface()

    def criar_interface(self):
        # --- HEADER ---
        header = ttk.Frame(self, padding=20, bootstyle="secondary")
        header.pack(fill=X)
        
        # Layout Flexível no Header
        h_container = ttk.Frame(header, bootstyle="secondary")
        h_container.pack(fill=X)
        
        ttk.Label(h_container, text="MAX IMPORT", font=("Segoe UI", 24, "bold"), bootstyle="inverse-secondary").pack(side=LEFT)
        ttk.Label(h_container, text=" | Ferramenta de Migração Inteligente", font=("Segoe UI", 14), bootstyle="inverse-secondary").pack(side=LEFT, pady=(8,0))
        ttk.Label(h_container, text="v2.0", font=("Consolas", 10), bootstyle="inverse-secondary").pack(side=RIGHT, pady=(10,0))

        # --- CONTAINER PRINCIPAL ---
        # Adicionei um padding maior para o conteúdo "respirar"
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(fill=BOTH, expand=True)

        # --- SEÇÃO 1: CONEXÃO (Card Style) ---
        # Labelframe com estilo 'info' para destacar a borda
        frame_db = ttk.Labelframe(main_frame, text=" 📡 Conectividade ", padding=15, bootstyle="info")
        frame_db.pack(fill=X, pady=(0, 15))

        # Grid interno para alinhar campos
        frame_db.columnconfigure(0, weight=1)
        frame_db.columnconfigure(1, weight=1)
        frame_db.columnconfigure(2, weight=0)

        # Campo Servidor
        lbl_srv = ttk.Label(frame_db, text="SERVIDOR SQL", font=("Segoe UI", 8, "bold"), bootstyle="info")
        lbl_srv.grid(row=0, column=0, sticky=W, padx=5)
        
        inp_srv_frame = ttk.Frame(frame_db)
        inp_srv_frame.grid(row=1, column=0, sticky=EW, padx=5, pady=(5,0))
        
        ttk.Entry(inp_srv_frame, textvariable=self.db_server, font=("Segoe UI", 10)).pack(side=LEFT, fill=X, expand=True)
        ttk.Button(inp_srv_frame, text="🔍", bootstyle="info-outline", command=self.listar_bancos_gui).pack(side=RIGHT, padx=(5,0))

        # Campo Banco
        lbl_db = ttk.Label(frame_db, text="BANCO DE DADOS", font=("Segoe UI", 8, "bold"), bootstyle="info")
        lbl_db.grid(row=0, column=1, sticky=W, padx=5)
        
        self.cbo_bancos = ttk.Combobox(frame_db, textvariable=self.db_name, state="normal", font=("Segoe UI", 10))
        self.cbo_bancos.grid(row=1, column=1, sticky=EW, padx=5, pady=(5,0))

        # Botão Conectar
        btn_conectar = ttk.Button(frame_db, text="💾 SALVAR & CONECTAR", command=self.atualizar_conexao, bootstyle="info", width=20)
        btn_conectar.grid(row=1, column=2, padx=10, pady=(5,0), ipady=2)

        # --- SEÇÃO 2: ARQUIVO (Card Style) ---
        frame_file = ttk.Labelframe(main_frame, text=" 📂 Origem dos Dados ", padding=15, bootstyle="primary")
        frame_file.pack(fill=X, pady=(0, 15))

        file_container = ttk.Frame(frame_file)
        file_container.pack(fill=X)

        ttk.Entry(file_container, textvariable=self.caminho_excel, state="readonly", font=("Segoe UI", 10)).pack(side=LEFT, fill=X, expand=True, padx=(0, 10))
        ttk.Button(file_container, text="SELECIONAR EXCEL", command=self.selecionar_arquivo, bootstyle="primary-outline", cursor="hand2").pack(side=RIGHT)

        # --- SEÇÃO 3: DASHBOARD DE AÇÕES ---
        # Substituí o Grid simples por um layout mais visual
        frame_acoes = ttk.Labelframe(main_frame, text=" 🚀 Painel de Controle ", padding=15, bootstyle="light")
        frame_acoes.pack(fill=X, pady=(0, 15))

        # Container interno para centralizar botões
        grid_acoes = ttk.Frame(frame_acoes)
        grid_acoes.pack(fill=X)
        
        # Colunas com peso igual para botões ficarem do mesmo tamanho
        for i in range(4): grid_acoes.columnconfigure(i, weight=1)
        grid_acoes.columnconfigure(4, weight=0) # Separador
        grid_acoes.columnconfigure(5, weight=1) # Limpeza

        # Botões Grandes e Coloridos
        btn_prod = ttk.Button(grid_acoes, text=" PRODUTOS", image="", compound=LEFT, bootstyle="success", style='Big.TButton', command=lambda: self.preparar_importacao(1))
        btn_prod.grid(row=0, column=0, padx=5, sticky=EW, ipady=10)

        btn_cli = ttk.Button(grid_acoes, text=" CLIENTES", bootstyle="primary", style='Big.TButton', command=lambda: self.preparar_importacao(2))
        btn_cli.grid(row=0, column=1, padx=5, sticky=EW, ipady=10)

        btn_forn = ttk.Button(grid_acoes, text=" FORNECEDORES", bootstyle="warning", style='Big.TButton', command=lambda: self.preparar_importacao(3))
        btn_forn.grid(row=0, column=2, padx=5, sticky=EW, ipady=10)

        btn_fin = ttk.Button(grid_acoes, text=" FINANCEIRO", bootstyle="info", style='Big.TButton', command=lambda: self.preparar_importacao(4))
        btn_fin.grid(row=0, column=3, padx=5, sticky=EW, ipady=10)

        # Separador Vertical
        ttk.Separator(grid_acoes, orient=VERTICAL).grid(row=0, column=4, sticky=NS, padx=15)

        # Botão de Perigo
        btn_limpar = ttk.Button(grid_acoes, text=" MANUTENÇÃO", bootstyle="danger-outline", style='Big.TButton', command=self.abrir_menu_limpeza)
        btn_limpar.grid(row=0, column=5, padx=5, sticky=EW, ipady=10)

        # --- SEÇÃO 4: LOG E STATUS ---
        # Um frame bonito para o terminal
        frame_log = ttk.Frame(main_frame)
        frame_log.pack(fill=BOTH, expand=True)

        lbl_log = ttk.Label(frame_log, text="> Console de Execução", font=("Consolas", 10, "bold"), bootstyle="secondary")
        lbl_log.pack(anchor=W, pady=(0, 5))

        # Texto com fundo escuro (automático do tema superhero) e fonte verde/branca
        self.txt_log = ttk.Text(frame_log, height=10, font=("Consolas", 9), relief=FLAT, padx=10, pady=10)
        self.txt_log.pack(fill=BOTH, expand=True)
        
        # Redireciona print para o widget
        sys.stdout = TextRedirector(self.txt_log)

        # Barra de Progresso Striped (Listrada)
        self.barra_progresso = ttk.Progressbar(main_frame, variable=self.progress_val, bootstyle="success-striped", mode='indeterminate')
        self.barra_progresso.pack(fill=X, pady=(15, 5), ipady=2)
        
        # Footer
        ttk.Label(self, text="  Maxdata Sistemas © 2025  ", font=("Segoe UI", 9), bootstyle="inverse-secondary").pack(side=BOTTOM, fill=X)

    # --- FUNÇÕES LÓGICAS (MANTIDAS 100% IDÊNTICAS) ---

    def listar_bancos_gui(self):
        server = self.db_server.get()
        if not server:
            messagebox.showwarning("Aviso", "Digite o nome/IP do servidor primeiro.")
            return
            
        try:
            print(f"Buscando bancos no servidor {server}...")
            # Feedback visual rápido
            self.config(cursor="wait")
            self.update()
            
            lista = db.listar_bancos_disponiveis(server)
            self.cbo_bancos['values'] = lista
            print(f"Encontrados {len(lista)} bancos.")
            if lista:
                self.cbo_bancos.current(0)
                self.cbo_bancos.event_generate("<<ComboboxSelected>>")
        except Exception as e:
            messagebox.showerror("Erro de Conexão", str(e))
            print(f"Erro: {e}")
        finally:
            self.config(cursor="")

    def atualizar_conexao(self):
        config.DB_SERVER = self.db_server.get()
        config.DB_NAME = self.db_name.get()
        print("--- Atualizando Conexão ---")
        try:
            db.reconectar()
            messagebox.showinfo("Sucesso", f"Conectado com sucesso!\n\nServidor: {config.DB_SERVER}\nBanco: {config.DB_NAME}")
        except Exception as e:
            messagebox.showerror("Falha", f"Não foi possível conectar:\n{e}")

    def selecionar_arquivo(self):
        arquivo = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if arquivo:
            self.caminho_excel.set(arquivo)
            print(f"Arquivo selecionado: {os.path.basename(arquivo)}")

    def abrir_menu_limpeza(self):
        # Toplevel estilizada
        top = ttk.Toplevel(self)
        top.title("Zona de Manutenção")
        top.geometry("450x420")
        top.place_window_center()
        
        # Header de Perigo
        head = ttk.Frame(top, padding=20, bootstyle="danger")
        head.pack(fill=X)
        ttk.Label(head, text="⚠️ CUIDADO", font=("Segoe UI", 16, "bold"), bootstyle="inverse-danger").pack()
        ttk.Label(head, text="Ações de exclusão são irreversíveis", font=("Segoe UI", 10), bootstyle="inverse-danger").pack()

        content = ttk.Frame(top, padding=20)
        content.pack(fill=BOTH, expand=True)

        ttk.Button(content, text="Limpar APENAS Produtos", bootstyle="outline-danger", 
                   command=lambda: self.executar_limpeza_thread(1, top)).pack(fill=X, pady=5, ipady=5)
        
        ttk.Button(content, text="Limpar APENAS Clientes", bootstyle="outline-danger", 
                   command=lambda: self.executar_limpeza_thread(2, top)).pack(fill=X, pady=5, ipady=5)
        
        ttk.Button(content, text="Limpar APENAS Financeiro", bootstyle="outline-danger", 
                   command=lambda: self.executar_limpeza_thread(3, top)).pack(fill=X, pady=5, ipady=5)

        ttk.Separator(content).pack(fill=X, pady=15)

        btn_full = ttk.Button(content, text="☢️ RESETAR BANCO COMPLETO", bootstyle="danger", 
                   command=lambda: self.executar_limpeza_thread(99, top))
        btn_full.pack(fill=X, pady=5, ipady=10)

    def executar_limpeza_thread(self, opcao, janela_popup):
        janela_popup.destroy()
        if not messagebox.askyesno("Confirmar Exclusão", "Tem certeza absoluta? Os dados serão apagados permanentemente."):
            return
        t = threading.Thread(target=self._limpeza_worker, args=(opcao,))
        t.start()

    def _limpeza_worker(self, opcao):
        self.barra_progresso.start(10) # Velocidade da animação
        self.alternar_interface("disabled")
        try:
            db.toggle_constraints(False)
            if opcao == 1 or opcao == 99:
                print("Limpando Produtos e Estoque...")
                db.limpar_tabela('prolote', reset_identity=True)
                db.executar_comando("DELETE FROM produto_empresa") 
                db.executar_comando("DELETE FROM produto WHERE proId > 1") 
                db.limpar_tabela('produtoUn', reset_identity=True)
            
            if opcao == 2 or opcao == 99:
                print("Limpando Clientes (Preservando Admin e Sistema)...")
                sql_delphi = "DELETE FROM cliente WHERE cliId <> 1 AND cliTipoCad <> 5 AND cliTipoCad <> 6"
                db.executar_comando(sql_delphi)

            if opcao == 3 or opcao == 99:
                print("Limpando Financeiro...")
                db.limpar_tabela('financeiro')
            
            print("--- Limpeza Concluída ---")
            messagebox.showinfo("Sucesso", "Limpeza realizada com sucesso!")
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
            messagebox.showwarning("Atenção", "Por favor, selecione um arquivo Excel primeiro.")
            return
        if not os.path.exists(arquivo):
            messagebox.showerror("Erro", "O arquivo especificado não foi encontrado.")
            return

        # Produtos (1) e Clientes (2) abrem mapa.
        if opcao in [1, 2]:
            try:
                # Leitura rápida só do cabeçalho
                df_header = pd.read_excel(arquivo, sheet_name=0, nrows=0)
                colunas = list(df_header.columns)
                tipo = "PRODUTO" if opcao == 1 else "CLIENTE"
                
                # Abre janela de mapeamento (agora estilizada)
                dialogo = ui_mapeamento.DialogoMapeamento(self, colunas, tipo_importacao=tipo)
                self.wait_window(dialogo)
                
                if dialogo.resultado:
                    t = threading.Thread(target=self.processar_thread, args=(opcao, dialogo.resultado))
                    t.start()
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao ler cabeçalho do Excel: {e}")
        else:
            if messagebox.askyesno("Confirmar Importação", "Deseja iniciar a importação direta dos dados?"):
                t = threading.Thread(target=self.processar_thread, args=(opcao, None))
                t.start()

    def processar_thread(self, opcao, mapa_colunas):
        self.barra_progresso.start(15)
        self.alternar_interface("disabled")
        try:
            db.reconectar()
            db.toggle_constraints(False)
            arquivo = self.caminho_excel.get()

            if opcao == 1:
                import_produtos.executar_importacao(arquivo, mapa_colunas, limpar_base=False)
            elif opcao == 2:
                import_clientes.executar_importacao(arquivo, mapa_colunas=mapa_colunas, is_fornecedor=False, limpar_base=False)
            elif opcao == 3:
                import_clientes.executar_importacao(arquivo, mapa_colunas=None, is_fornecedor=True, limpar_base=False)
            elif opcao == 4:
                import_financeiro.executar_importacao(arquivo, limpar_base=False)

            messagebox.showinfo("Processo Finalizado", "A importação foi concluída com sucesso!")

        except Exception as e:
            print(f"ERRO FATAL: {e}")
            messagebox.showerror("Erro Crítico", f"Falha durante a importação:\n{e}")
        finally:
            db.toggle_constraints(True)
            self.barra_progresso.stop()
            self.alternar_interface("normal")

    def alternar_interface(self, estado):
        # Opcional: Desabilitar botões durante processamento
        pass 

if __name__ == "__main__":
    app = MaxImportApp()
    app.mainloop()