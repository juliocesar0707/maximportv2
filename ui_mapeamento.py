# ui_mapeamento.py
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import messagebox

class DialogoMapeamento(ttk.Toplevel):
    def __init__(self, parent, colunas_excel, tipo_importacao="PRODUTO"):
        super().__init__(parent)
        self.title(f"Mapeamento Inteligente - {tipo_importacao}")
        self.geometry("950x780")
        self.place_window_center()
        self.resultado = None 
        
        # --- HEADER ---
        header_frame = ttk.Frame(self, padding=20, bootstyle="secondary")
        header_frame.pack(fill=X)
        
        icon = "📦" if tipo_importacao == "PRODUTO" else "👥"
        ttk.Label(header_frame, text=f"{icon} Mapeamento de Colunas", font=("Segoe UI", 18, "bold"), bootstyle="inverse-secondary").pack(side=LEFT)
        ttk.Label(header_frame, text="Relacione os campos do Excel com o Sistema", font=("Segoe UI", 10), bootstyle="inverse-secondary").pack(side=RIGHT, pady=(10,0))
        
        # --- ÁREA DE SCROLL ---
        self.container = ttk.Frame(self, padding=2)
        self.container.pack(fill=BOTH, expand=True, padx=20, pady=10)
        
        self.canvas = ttk.Canvas(self.container, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self.container, orient="vertical", command=self.canvas.yview, bootstyle="rounded")
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side=LEFT, fill=BOTH, expand=True)
        self.scrollbar.pack(side=RIGHT, fill=Y)

        self.colunas_excel = ["(Ignorar / Não Importar)"] + colunas_excel
        self.combos = {}

        # Dicionários de Campos (Mantido a lógica original)
        if tipo_importacao == "PRODUTO":
            self.campos_sistema = {
                'proId': 'ID Produto (Fixo ou Automático)',
                'zzz_proCodigo': 'Referência / Cód. Fabricante',
                'cdbCodigo': 'Código de Barras (EAN/GTIN)',
                'proDescricao': 'Descrição do Produto',
                'zzz_proCodigoNcm': 'NCM (Código Fiscal)',
                'proUn': 'Unidade (UN, KG, CX)',
                'zzz_proCusto': 'Preço de Custo (R$)',
                'zzz_proVenda': 'Preço de Venda (R$)',
                'proEstoqueAtual': 'Estoque Atual',
                'zzz_proEstoqueMin': 'Estoque Mínimo',
                'proCodcst2': 'CST ICMS (Ex: 00, 60)',
                'proCodCSOSN': 'CSOSN (Simples Nacional)'
            }
        elif tipo_importacao == "CLIENTE":
            self.campos_sistema = {
                'cliId': 'ID Cliente (Código)',
                'cliNome': 'Nome / Razão Social',
                'cliFantasia': 'Nome Fantasia / Apelido',
                'cliTipo': 'Tipo Pessoa (0=Física, 1=Jurídica)',
                'cliCpfCgc': 'CPF / CNPJ',
                'cliRgInsc': 'RG / Inscrição Estadual',
                'cliDatCad': 'Data de Cadastro',
                'cliFatCep': 'CEP (Faturamento)',
                'cliFatEnd': 'Endereço (Faturamento)',
                'cliFatEndNumero': 'Número (Faturamento)',
                'cliFatBairro': 'Bairro (Faturamento)',
                'cliFatCidade': 'Cidade (Faturamento)',
                'cliFatUf': 'UF (Faturamento)',
                'cliFatCidCodIBGE': 'Código IBGE Cidade',
                'cliCobCep': 'CEP (Cobrança)',
                'cliCobEnd': 'Endereço (Cobrança)',
                'cliCobEndNumero': 'Número (Cobrança)',
                'cliCobBairro': 'Bairro (Cobrança)',
                'cliCobCidade': 'Cidade (Cobrança)',
                'cliCobUf': 'UF (Cobrança)',
                'cliEmail': 'Email',
                'CliFone': 'Telefone Fixo',
                'cliCelular': 'Celular / WhatsApp',
                'CliFax': 'Fax / Outro',
                'CliLimitCred': 'Limite de Crédito (R$)',
                'zzz_CliObsVend': 'Observações',
                'CliContNome1': 'Nome Contato',
                'CliContDepto1': 'Departamento',
                'CliContFone1': 'Telefone Contato',
                'CliCadNomePai': 'Nome do Pai',
                'CliCadNomeMae': 'Nome da Mãe'
            }

        # --- CABEÇALHO DA LISTA ---
        # Usando Frames coloridos para cabeçalho
        head_grid = ttk.Frame(self.scrollable_frame, padding=5)
        head_grid.grid(row=0, column=0, columnspan=2, sticky=EW, pady=(0,10))
        
        lbl_dest = ttk.Label(head_grid, text="  CAMPO NO BANCO DE DADOS", font=("Segoe UI", 9, "bold"), bootstyle="inverse-info", width=40, anchor=W)
        lbl_dest.pack(side=LEFT, fill=X, expand=True, padx=1)
        
        lbl_orig = ttk.Label(head_grid, text="  COLUNA NA SUA PLANILHA", font=("Segoe UI", 9, "bold"), bootstyle="inverse-warning", width=40, anchor=W)
        lbl_orig.pack(side=LEFT, fill=X, expand=True, padx=1)

        row = 1
        for campo_db, label_amigavel in self.campos_sistema.items():
            # Label do Campo Sistema
            ttk.Label(self.scrollable_frame, text=f"• {label_amigavel}", font=("Segoe UI", 10)).grid(row=row, column=0, sticky=W, padx=15, pady=8)
            
            # Combobox Estilizado
            cbox = ttk.Combobox(self.scrollable_frame, values=self.colunas_excel, state="readonly", width=45, font=("Segoe UI", 9))
            cbox.grid(row=row, column=1, sticky=EW, padx=10, pady=8)
            self.combos[campo_db] = cbox
            
            # --- Lógica de Auto-Seleção (Mantida) ---
            cbox.set("(Ignorar / Não Importar)")
            for col_excel in colunas_excel:
                nm_db = campo_db.lower()
                nm_ex = str(col_excel).lower()
                match = False
                
                # Regras Genéricas
                if nm_db == nm_ex: match = True
                
                # Produto
                if campo_db == 'zzz_proCodigo' and nm_ex in ['referencia', 'ref', 'codigo', 'cod', 'fabricante']: match = True
                if campo_db == 'cdbCodigo' and nm_ex in ['barras', 'ean', 'gtin', 'codigo de barras', 'cod barras', 'cod. barras']: match = True
                if campo_db == 'proDescricao' and nm_ex in ['nome', 'descricao', 'descrição', 'produto']: match = True
                if 'custo' in nm_db and 'custo' in nm_ex: match = True
                if 'venda' in nm_db and 'venda' in nm_ex: match = True
                if 'un' in nm_db and 'unidade' in nm_ex: match = True
                
                # Cliente
                if campo_db == 'cliId' and nm_ex in ['id', 'codigo', 'código']: match = True
                if campo_db == 'cliNome' and nm_ex in ['nome', 'razao', 'razão', 'cliente']: match = True
                if campo_db == 'cliFantasia' and nm_ex in ['fantasia', 'apelido']: match = True
                if campo_db == 'cliCpfCgc' and ('cpf' in nm_ex or 'cnpj' in nm_ex): match = True
                if campo_db == 'cliTipo' and ('tipo' in nm_ex): match = True
                
                # Endereço
                if 'cep' in nm_db and 'cep' in nm_ex: match = True
                if 'bairro' in nm_db and 'bairro' in nm_ex: match = True
                if 'cidade' in nm_db and 'cidade' in nm_ex: match = True
                if 'uf' in nm_db and 'uf' in nm_ex: match = True
                if campo_db == 'cliFatEnd' and ('rua' in nm_ex or 'endereco' in nm_ex): match = True
                if 'numero' in nm_db and ('num' in nm_ex or 'nº' in nm_ex): match = True
                
                # Contato
                if 'email' in nm_db and 'email' in nm_ex: match = True
                if 'fone' in nm_db and 'telefone' in nm_ex: match = True
                if 'celular' in nm_db and 'celular' in nm_ex: match = True
                if 'obs' in nm_db and 'obs' in nm_ex: match = True

                if match:
                    cbox.set(col_excel)
                    break
            
            # Separador sutil
            ttk.Separator(self.scrollable_frame, bootstyle="secondary").grid(row=row+1, column=0, columnspan=2, sticky=EW, padx=10, pady=0)
            
            row += 2

        # --- FOOTER ---
        btn_frame = ttk.Frame(self, padding=20, bootstyle="light")
        btn_frame.pack(fill=X, side=BOTTOM)
        
        btn_confirmar = ttk.Button(btn_frame, text="✅ CONFIRMAR E IMPORTAR", bootstyle="success", command=self.confirmar)
        btn_confirmar.pack(fill=X, ipady=10)
        
        # Habilitar scroll com mousewheel
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def confirmar(self):
        mapa = {}
        for campo_db, cbox in self.combos.items():
            valor = cbox.get()
            if valor and valor != "(Ignorar / Não Importar)":
                mapa[campo_db] = valor
        
        self.resultado = mapa
        self.destroy()