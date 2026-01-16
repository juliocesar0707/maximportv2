# ui_mapeamento.py
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import messagebox

class DialogoMapeamento(ttk.Toplevel):
    def __init__(self, parent, colunas_excel, tipo_importacao="PRODUTO"):
        super().__init__(parent)
        self.title(f"Mapeamento de Colunas - {tipo_importacao}")
        self.geometry("900x750") # Aumentei um pouco a altura
        self.resultado = None 
        
        ttk.Label(self, text="Relacione as colunas para importação", font=("Segoe UI", 14, "bold"), bootstyle="primary").pack(pady=10)
        
        # Container com Scrollbar (para caber todos os campos novos)
        self.container = ttk.Frame(self)
        self.container.pack(fill=BOTH, expand=True, padx=10, pady=5)
        
        self.canvas = ttk.Canvas(self.container)
        self.scrollbar = ttk.Scrollbar(self.container, orient="vertical", command=self.canvas.yview)
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

        if tipo_importacao == "PRODUTO":
            self.campos_sistema = {
                # Mapeamento Híbrido
                'proId': 'ID Produto (Se numérico = Fixo, Se letra = Automático)',
                'zzz_proCodigo': 'Referência / Código de Barras',
                'proDescricao': 'Descrição do Produto',
                'zzz_proCodigoNcm': 'NCM (Código Fiscal)',
                
                # Campos Fiscais e Valores
                'proUn': 'Unidade (UN, KG, CX)',
                'zzz_proCusto': 'Preço de Custo',
                'zzz_proVenda': 'Preço de Venda',
                'proEstoqueAtual': 'Estoque Atual',
                'zzz_proEstoqueMin': 'Estoque Mínimo',
                'proCodcst2': 'CST ICMS (Ex: 00, 60, 40)',
                'proCodCSOSN': 'CSOSN (Simples Nacional)'
            }
        elif tipo_importacao == "CLIENTE":
            self.campos_sistema = {
                # Dados Principais
                'cliId': 'ID Cliente (Código)',
                'cliNome': 'Nome / Razão Social',
                'cliFantasia': 'Nome Fantasia / Apelido',
                'cliTipo': 'Tipo Pessoa (0=Física, 1=Jurídica)',
                'cliCpfCgc': 'CPF / CNPJ',
                'cliRgInsc': 'RG / Inscrição Estadual',
                'cliDatCad': 'Data de Cadastro',
                
                # Endereço Faturamento
                'cliFatCep': 'CEP (Faturamento)',
                'cliFatEnd': 'Endereço / Rua (Faturamento)',
                'cliFatEndNumero': 'Número (Faturamento)',
                'cliFatBairro': 'Bairro (Faturamento)',
                'cliFatCidade': 'Cidade (Faturamento)',
                'cliFatUf': 'UF (Faturamento)',
                'cliFatCidCodIBGE': 'Código IBGE Cidade',

                # Endereço Cobrança
                'cliCobCep': 'CEP (Cobrança)',
                'cliCobEnd': 'Endereço / Rua (Cobrança)',
                'cliCobEndNumero': 'Número (Cobrança)',
                'cliCobBairro': 'Bairro (Cobrança)',
                'cliCobCidade': 'Cidade (Cobrança)',
                'cliCobUf': 'UF (Cobrança)',

                # Contato e Financeiro
                'cliEmail': 'Email',
                'CliFone': 'Telefone Fixo',
                'cliCelular': 'Celular / WhatsApp',
                'CliFax': 'Fax / Outro Telefone',
                'CliLimitCred': 'Limite de Crédito (R$)',
                'zzz_CliObsVend': 'Observações do Cadastro',
                
                # Filiação e Pessoas de Contato
                'CliContNome1': 'Nome da Pessoa de Contato',
                'CliContDepto1': 'Departamento do Contato',
                'CliContFone1': 'Telefone do Contato',
                'CliCadNomePai': 'Nome do Pai',
                'CliCadNomeMae': 'Nome da Mãe'
            }

        # Cabeçalhos
        ttk.Label(self.scrollable_frame, text="DESTINO (BANCO DE DADOS)", font=("Segoe UI", 10, "bold"), bootstyle="inverse-secondary").grid(row=0, column=0, sticky=EW, padx=5, pady=5)
        ttk.Label(self.scrollable_frame, text="ORIGEM (SEU EXCEL)", font=("Segoe UI", 10, "bold"), bootstyle="inverse-success").grid(row=0, column=1, sticky=EW, padx=5, pady=5)

        row = 1
        for campo_db, label_amigavel in self.campos_sistema.items():
            ttk.Label(self.scrollable_frame, text=label_amigavel, font=("Segoe UI", 10)).grid(row=row, column=0, sticky=W, padx=10, pady=5)
            
            cbox = ttk.Combobox(self.scrollable_frame, values=self.colunas_excel, state="readonly", width=40)
            cbox.grid(row=row, column=1, sticky=EW, padx=10, pady=5)
            self.combos[campo_db] = cbox
            
            # --- Inteligência de Auto-Seleção (Deixando o trabalho mais fácil) ---
            cbox.set("(Ignorar / Não Importar)")
            for col_excel in colunas_excel:
                nm_db = campo_db.lower()
                nm_ex = str(col_excel).lower()
                
                match = False
                
                # Regras Genéricas
                if nm_db == nm_ex: match = True
                
                # Regras Específicas Produto
                if campo_db == 'zzz_proCodigo' and nm_ex in ['referencia', 'ref', 'codigo', 'barras', 'cod']: match = True
                if campo_db == 'proDescricao' and nm_ex in ['nome', 'descricao', 'descrição', 'produto']: match = True
                if 'custo' in nm_db and 'custo' in nm_ex: match = True
                if 'venda' in nm_db and 'venda' in nm_ex: match = True
                if 'un' in nm_db and 'unidade' in nm_ex: match = True
                
                # Regras Específicas Cliente
                if campo_db == 'cliId' and nm_ex in ['id', 'codigo', 'código']: match = True
                if campo_db == 'cliNome' and nm_ex in ['nome', 'razao', 'razão', 'cliente']: match = True
                if campo_db == 'cliFantasia' and nm_ex in ['fantasia', 'apelido']: match = True
                if campo_db == 'cliCpfCgc' and ('cpf' in nm_ex or 'cnpj' in nm_ex or 'documento' in nm_ex): match = True
                if campo_db == 'cliTipo' and ('tipo' in nm_ex or 'fisica' in nm_ex): match = True
                
                # Endereço
                if 'cep' in nm_db and 'cep' in nm_ex: match = True
                if 'bairro' in nm_db and 'bairro' in nm_ex: match = True
                if 'cidade' in nm_db and 'cidade' in nm_ex: match = True
                if 'uf' in nm_db and 'uf' in nm_ex: match = True
                if campo_db == 'cliFatEnd' and ('rua' in nm_ex or 'endereco' in nm_ex or 'endereço' in nm_ex): match = True
                if 'numero' in nm_db and ('num' in nm_ex or 'nº' in nm_ex): match = True
                
                # Contato
                if 'email' in nm_db and 'email' in nm_ex: match = True
                if 'fone' in nm_db and 'telefone' in nm_ex: match = True
                if 'celular' in nm_db and 'celular' in nm_ex: match = True
                if 'obs' in nm_db and ('obs' in nm_ex or 'observacao' in nm_ex): match = True

                if match:
                    cbox.set(col_excel)
                    break
            row += 1

        # Botão Confirmar (fixo na parte inferior)
        btn_frame = ttk.Frame(self, padding=10)
        btn_frame.pack(fill=X, side=BOTTOM)
        ttk.Separator(btn_frame).pack(fill=X, pady=5)
        
        btn_confirmar = ttk.Button(btn_frame, text="CONFIRMAR E PROCESSAR", bootstyle="success", command=self.confirmar)
        btn_confirmar.pack(fill=X, ipady=5)
        
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
        
        # Validação simples
        if 'proDescricao' in self.campos_sistema and 'proDescricao' not in mapa and 'cliNome' not in mapa:
             # Se for produto exige descricao, se for cliente exige nome. 
             # Simplificação: verifica se o mapa está muito vazio
             pass
            
        self.resultado = mapa
        self.destroy()