# ui_mapeamento.py
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import messagebox

class DialogoMapeamento(ttk.Toplevel):
    def __init__(self, parent, colunas_excel, tipo_importacao="PRODUTO"):
        super().__init__(parent)
        self.title(f"Mapeamento de Colunas - {tipo_importacao}")
        self.geometry("750x650")
        self.resultado = None 
        
        ttk.Label(self, text="Relacione as colunas para importação", font=("Segoe UI", 14, "bold"), bootstyle="primary").pack(pady=15)
        
        self.frame_main = ttk.Frame(self, padding=20)
        self.frame_main.pack(fill=BOTH, expand=True)

        self.colunas_excel = ["(Ignorar / Não Importar)"] + colunas_excel
        self.combos = {}

        if tipo_importacao == "PRODUTO":
            self.campos_sistema = {
                # Mapeamento Híbrido: Vai para produto e produto_empresa
                'proId': 'ID Produto (Se numérico = Fixo, Se letra = Automático)',
                'zzz_proCodigo': 'Referência / Código de Barras',
                'proDescricao': 'Descrição do Produto',
                
                # Campos exclusivos da tabela PRODUTO_EMPRESA
                'proUn': 'Unidade (UN, KG, CX) -> Tab. Empresa',
                'zzz_proCusto': 'Preço de Custo -> Tab. Empresa',
                'zzz_proVenda': 'Preço de Venda -> Tab. Empresa',
                'proEstoqueAtual': 'Estoque Atual -> Tab. Empresa',
                'zzz_proEstoqueMin': 'Estoque Mínimo -> Tab. Empresa',
                'zzz_proCodigoNcm': 'NCM (Classificação Fiscal)'
            }
        elif tipo_importacao == "CLIENTE":
            self.campos_sistema = {
                'cliId': 'ID Cliente',
                'cliNome': 'Nome / Razão Social',
                'cliFantasia': 'Nome Fantasia',
                'cliCpfCgc': 'CPF / CNPJ',
                'cliFone': 'Telefone',
                'cliEmail': 'Email',
                'cliFatEnd': 'Endereço (Rua)',
                'cliFatBairro': 'Bairro',
                'cliFatCidade': 'Cidade',
                'cliFatUf': 'UF'
            }

        # Cabeçalhos
        ttk.Label(self.frame_main, text="DESTINO (BANCO DE DADOS)", font=("Segoe UI", 10, "bold"), bootstyle="inverse-secondary").grid(row=0, column=0, sticky=EW, padx=5, pady=5)
        ttk.Label(self.frame_main, text="ORIGEM (SEU EXCEL)", font=("Segoe UI", 10, "bold"), bootstyle="inverse-success").grid(row=0, column=1, sticky=EW, padx=5, pady=5)

        row = 1
        for campo_db, label_amigavel in self.campos_sistema.items():
            ttk.Label(self.frame_main, text=label_amigavel, font=("Segoe UI", 10)).grid(row=row, column=0, sticky=W, padx=10, pady=8)
            
            cbox = ttk.Combobox(self.frame_main, values=self.colunas_excel, state="readonly", width=35)
            cbox.grid(row=row, column=1, sticky=EW, padx=10, pady=8)
            self.combos[campo_db] = cbox
            
            # Inteligência de Auto-Seleção
            cbox.set("(Ignorar / Não Importar)")
            for col_excel in colunas_excel:
                nm_db = campo_db.lower()
                nm_ex = str(col_excel).lower()
                
                match = False
                if campo_db == 'zzz_proCodigo' and nm_ex in ['referencia', 'ref', 'codigo', 'barras', 'cod']: match = True
                if campo_db == 'proDescricao' and nm_ex in ['nome', 'descricao', 'descrição', 'produto']: match = True
                if 'custo' in nm_db and 'custo' in nm_ex: match = True
                if 'venda' in nm_db and 'venda' in nm_ex: match = True
                if 'un' in nm_db and 'unidade' in nm_ex: match = True
                if 'ncm' in nm_db and 'ncm' in nm_ex: match = True
                if 'estoque' in nm_db and 'estoque' in nm_ex: match = True
                if campo_db == 'proId' and nm_ex in ['id', 'código', 'codigo', 'cod']: match = True

                if match:
                    cbox.set(col_excel)
                    break
            row += 1

        ttk.Separator(self.frame_main).grid(row=row, column=0, columnspan=2, sticky=EW, pady=20)
        
        btn_confirmar = ttk.Button(self.frame_main, text="CONFIRMAR E PROCESSAR", bootstyle="success", command=self.confirmar)
        btn_confirmar.grid(row=row+1, column=0, columnspan=2, sticky=EW, pady=10, ipady=5)

    def confirmar(self):
        mapa = {}
        for campo_db, cbox in self.combos.items():
            valor = cbox.get()
            if valor and valor != "(Ignorar / Não Importar)":
                mapa[campo_db] = valor
        
        if 'proDescricao' in self.campos_sistema and 'proDescricao' not in mapa:
            messagebox.showwarning("Erro", "O campo 'Descrição' é OBRIGATÓRIO.")
            return
            
        self.resultado = mapa
        self.destroy()