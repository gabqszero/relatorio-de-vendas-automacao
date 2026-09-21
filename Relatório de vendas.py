import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import os
from openpyxl import load_workbook
import threading
from datetime import datetime
import re 

# Configuração da aparência (FrontEnzo, Aqui a gente mete o dane-se e a IA faz)
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

#  FUNÇÃO NO ESCOPO GLOBAL 
def extrair_ultimo_ano(valor):
    if pd.isna(valor) or str(valor).strip().lower() == 'nan':
        return None
    
    # Busca todas as ocorrências de anos entre 2020 e 2026 na célula
    anos_encontrados = re.findall(r'(202[0-6])', str(valor))
    
    if anos_encontrados:
        return int(anos_encontrados[-1])  # [-1] garante que pega SEMPRE o último ano da lista
    return None

class RunWorkElite(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("RunWorkElite - CSV ENGINE")
        self.geometry("600x700")
        self.configure(fg_color="#16181d")

        self.path_destino = ""
        self.paths_origem = []

        self.title_bar = ctk.CTkFrame(self, fg_color="#0d0f12", height=35, corner_radius=0)
        self.title_bar.pack(fill="x", side="top")
        ctk.CTkLabel(self.title_bar, text="RUNWORK ELITE PRO", font=("Arial", 11, "bold"), text_color="gray").pack(side="left", padx=15)
        
        ctk.CTkLabel(self, text="SISTEMA DE INJEÇÃO PANDAS", font=("Arial", 22, "bold"), text_color="#00d4ff").pack(pady=30)

        self.btn_dest = ctk.CTkButton(self, text="1. SELECIONAR EXCEL MESTRE", fg_color="#1c1f26", border_width=1, border_color="#ff6600", height=45, command=self.select_destino)
        self.btn_dest.pack(pady=10, padx=50, fill="x")
        self.lbl_dest = ctk.CTkLabel(self, text="Aguardando destino...", font=("Arial", 11), text_color="gray")
        self.lbl_dest.pack()

        self.btn_origem = ctk.CTkButton(self, text="2. SELECIONAR CSVs DE ORIGEM", fg_color="#1c1f26", border_width=1, border_color="#00d4ff", height=45, command=self.select_origem)
        self.btn_origem.pack(pady=10, padx=50, fill="x")
        self.lbl_origem = ctk.CTkLabel(self, text="Aguardando arquivos...", font=("Arial", 11), text_color="gray")
        self.lbl_origem.pack()

        self.txt_log = ctk.CTkTextbox(self, width=500, height=200, fg_color="#0d0f12", text_color="#00ff88", font=("Consolas", 12))
        self.txt_log.pack(pady=20, padx=20)
        self.log("Sistema pronto.")

        self.progress = ctk.CTkProgressBar(self, width=500, progress_color="#ff6600")
        self.progress.set(0)
        self.progress.pack(pady=10)

        self.btn_run = ctk.CTkButton(self, text="EXECUTAR MÁQUINA PANDAS", fg_color="#ff6600", state="disabled", font=("Arial", 14, "bold"), height=55, command=self.start_thread)
        self.btn_run.pack(pady=20, padx=70, fill="x")

    def log(self, msg):
        self.txt_log.insert("end", f"> {msg}\n")
        self.txt_log.see("end")
        self.update_idletasks() # Força o CustomTkinter a atualizar a tela em tempo real

    def select_destino(self):
        self.path_destino = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if self.path_destino:
            self.lbl_dest.configure(text=f"✔ {os.path.basename(self.path_destino)}", text_color="#00ff88")
            self.check_ready()

    def select_origem(self):
        files = filedialog.askopenfilenames(filetypes=[("CSV", "*.csv")])
        if files:
            self.paths_origem = sorted(list(files)) 
            self.lbl_origem.configure(text=f"✔ {len(self.paths_origem)} arquivos ordenados", text_color="#00ff88")
            self.check_ready()

    def check_ready(self):
        if self.path_destino and self.paths_origem:
            self.btn_run.configure(state="normal")

    def start_thread(self):
        self.btn_run.configure(state="disabled", text="EM EXECUÇÃO...")
        threading.Thread(target=self.processar, daemon=True).start()

    def processar(self):
        try:
            self.log("Pandas unificando arquivos em sequência...")
            lista_dfs = [pd.read_csv(p, sep=';', encoding='latin-1', dtype=str) for p in self.paths_origem]
            df_unificado = pd.concat(lista_dfs, ignore_index=True)
            
            # Esse baguiu Ignora as colunas AE e a coluna BJ do baguiu
            colunas_para_deletar = []
            if len(df_unificado.columns) > 30:
                colunas_para_deletar.append(df_unificado.columns[30]) 
            if len(df_unificado.columns) > 61:
                colunas_para_deletar.append(df_unificado.columns[61]) 

            if colunas_para_deletar:
                df_unificado = df_unificado.drop(columns=colunas_para_deletar)
                self.log("Ignorando colunas AE e BJ do relatório.")

            df_vendas = df_unificado

            self.log("Indexando tabela de custos...")
            tabela_custos = pd.read_excel(self.path_destino)

            total_linhas = len(df_vendas)
            self.log(f"Total de {total_linhas} linhas carregadas.")

            self.log("Saneando triagem e convertendo colunas numéricas exigidas...")
            #Removido o índice 30 desta lista para não destruir as strings de data transformando-as em NaN
            indices_para_converter = [0, 8, 14, 17, 28, 29, 33, 34, 35, 37]

            for idx in indices_para_converter:
                if idx < len(df_vendas.columns):
                    col_nome = df_vendas.columns[idx]
                    df_vendas[col_nome] = df_vendas[col_nome].astype(str).str.strip()
                    df_vendas[col_nome] = df_vendas[col_nome].str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
                    df_vendas[col_nome] = pd.to_numeric(df_vendas[col_nome], errors='coerce')
        
            self.log("Banco de custos indexado dinamicamente em memória.")
            
            #Tratamento de Nulos na Coluna AI executado no DataFrame ---
            if df_vendas.shape[1] < 35:
                raise ValueError(f"A planilha selecionada possui apenas {df_vendas.shape[1]} colunas. Uma coluna na posição 'AI' não pôde ser alcançada.")

            # Cabeçalhos baseando-se nas posições físicas do Excel
            coluna_do_codigo = df_vendas.columns[14]
            coluna_com_nulos = df_vendas.columns[34]

            # Identifica as linhas onde a coluna AI está nula/vazia
            filtro_nulos = df_vendas[coluna_com_nulos].isnull()
            codigos_com_nulo = df_vendas.loc[filtro_nulos, coluna_do_codigo].unique()
            
            if len(codigos_com_nulo) > 0:
                print("\n" + "="*60)
                print(f"INVESTIGAÇÃO: Encontrados {len(codigos_com_nulo)} produtos com falha na coluna AI.")
                print("="*60)

            # Converte os valores em número - PROTEÇÃO DA CHAVE DE ACESSO ADICIONADA
            def limpar_e_converter(valor):
                if pd.isna(valor):
                    return 0.0
                v_str = str(valor).strip()
                if v_str == "-" or v_str.lower() == "nan":
                    return 0.0
                try:
                    if "," in v_str:
                        v_str = v_str.replace(".", "")  # Remove pontos de milhar
                        v_str = v_str.replace(",", ".")  # Substitui a vírgula pelo ponto decimal do Python
                    return float(v_str)
                except ValueError:
                    return valor  # Retorna o texto original caso não consiga converter de jeito nenhum

            # Identifica textualmente qual é a coluna da Chave de Acesso (Coluna 5)
            coluna_chave_acesso = df_vendas.columns[4] if len(df_vendas.columns) > 4 else None

            # Varre as colunas aplicando a conversão numérica onde fizer sentido
            for col in df_vendas.columns:
                if col == coluna_do_codigo or col == coluna_chave_acesso:
                    continue  

                # Validação rápida de amostragem para saber se a coluna guarda números
                amostra = df_vendas[col].dropna().head(15).astype(str)
                se_numerico = (
                    amostra.str.replace(",", "", regex=False)
                    .str.replace(".", "", regex=False)
                    .str.replace("-", "", regex=False)
                    .str.isnumeric()
                    .any()
                )

                if se_numerico:
                    df_vendas[col] = df_vendas[col].apply(limpar_e_converter)

            # Automação de cálculo de custo
            mapa_valores_corretos = {}
            for codigo in codigos_com_nulo:
                historico_produto = df_vendas[(df_vendas[coluna_do_codigo] == codigo) & (df_vendas[coluna_com_nulos].notnull())]
                
                if not historico_produto.empty:
                    valor_valido = historico_produto[coluna_com_nulos].iloc[0]
                    mapa_valores_corretos[codigo] = valor_valido
                    print(f" -> Histórico: Código [{codigo}] mapeado com o valor [{valor_valido}]")
                else:
                    print(f" -> Atenção: Código [{codigo}] não possui histórico com valor preenchido.")

            # Substituir os nulos originais usando o mapa construído
            contagem_corrigidos = 0
            for idx in df_vendas[filtro_nulos].index:
                codigo_da_linha = df_vendas.at[idx, coluna_do_codigo]
                
                if codigo_da_linha in mapa_valores_corretos:
                    df_vendas.at[idx, coluna_com_nulos] = mapa_valores_corretos[codigo_da_linha]
                    contagem_corrigidos += 1

            print("\n" + "="*60)
            print(f"CONCLUÍDO: {contagem_corrigidos} linhas vazias foram resolvidas.")
            print("="*60 + "\n")

            self.log("Abrindo Excel Mestre...")
            wb = load_workbook(self.path_destino)
            ws = wb["BASE_ANÁL_OFICIAL"] if "BASE_ANÁL_OFICIAL" in wb.sheetnames else wb.active 
            start_row = ws.max_row + 1
            dados_lista = df_vendas.values.tolist()

            # Mapeamento de posição
            colunas_numericas = [22, 23, 24, 25, 26, 27, 28] 
            #mapear o índice 30 do DataFrame no c_idx 31
            colunas_data = [2, 6, 31] 

            self.log("Injetando dados e corrigindo formatos...")
            
            for r_idx, row in enumerate(dados_lista):
                linha_atual = start_row + r_idx
                
                # descobre a posição exata da coluna DATA_ENTRADA na memória pelo nome (evita erros se colunas mudarem de lugar)
            try:
                idx_data_entrada = df_vendas.columns.get_loc("DATA_ENTRADA")
            except:
                idx_data_entrada = 30 # Fallback caso o nome mude

            self.log("Injetando dados e corrigindo formatos...")
            
            for r_idx, row in enumerate(dados_lista):
                linha_atual = start_row + r_idx
                
                try:
                    qtd_linha = float(str(row[20]).strip().replace('.', '').replace(',', '.'))
                except:
                    qtd_linha = 0.0

                custo_unitario_final = 0.0 
                ano_extraido_linha = None  

                # EXTRAÇÃO: Pega o ano direto da célula de origem antes de misturar com outras colunas
                valor_data_origem = row[idx_data_entrada]
                if pd.notna(valor_data_origem) and str(valor_data_origem).strip().lower() != 'nan':
                    anos_encontrados = re.findall(r'(202[0-6])', str(valor_data_origem))
                    if anos_encontrados:
                        ano_extraido_linha = int(anos_encontrados[-1]) # Pega SEMPRE o último da direita

                # LOOP QUE INJETA AS COLUNAS NO EXCEL MESTRE
                for c_idx, valor in enumerate(row, start=1):
                    if c_idx == 61:
                        continue

                    if c_idx <= 32:
                        col_dest = c_idx
                    else:
                        col_dest = c_idx + 1
                    
                    cell = ws.cell(row=linha_atual, column=col_dest)

                    if pd.isna(valor) or str(valor).strip().lower() == 'nan':
                        cell.value = None
                        if c_idx != 32 and c_idx != 33:
                            continue

                    # 1. Formatação de DATA (Apenas limpa o visual, a extração do ano já foi feita com segurança acima)
                    if c_idx in colunas_data:
                        try:
                            v_clean = str(valor).split()[0].replace('-', '/')
                            if ',' in v_clean:
                                v_clean = v_clean.split(',')[0].strip()
                                
                            data_obj = datetime.strptime(v_clean, "%d/%m/%Y")
                            cell.value = data_obj
                            cell.number_format = 'dd/mm/yyyy'
                        except:
                            try:
                                v_clean = str(valor).split()[0].replace('-', '/')
                                if ',' in v_clean:
                                    v_clean = v_clean.split(',')[0].strip()
                                data_obj = datetime.strptime(v_clean, "%Y/%m/%d")
                                cell.value = data_obj
                                cell.number_format = 'dd/mm/yyyy'
                            except:
                                cell.value = str(valor).strip()

                    # 2. CHAVE DE ACESSO
                    elif c_idx == 5:
                        cell.value = str(valor).strip()
                        cell.number_format = '@'
                    
                    # 3. CONVERSÃO NUMÉRICA PADRÃO
                    elif c_idx in colunas_numericas:
                        try:
                            cell.value = float(valor)
                            cell.number_format = '#,##0.00'
                        except:
                            cell.value = valor
                                         
                    else:
                        if isinstance(valor, (int, float)) and pd.notna(valor):
                            cell.value = valor
                            if col_dest in [22, 23, 24, 25, 26, 27, 28, 32, 33]:
                                cell.number_format = '#,##0.00'
                        else:
                            cell.value = valor

                # GRAVAÇÃO DO ANO NA COLUNA VERMELHA (AG - ENTRA)
                if ano_extraido_linha:
                    cell_entra = ws.cell(row=linha_atual, column=33) # Coluna 33 = AG
                    cell_entra.value = ano_extraido_linha
                    cell_entra.number_format = '0'
            self.log("Salvando arquivo mestre (Aguarde)...")
            wb.save(self.path_destino) 
            self.log("✔ Sincronização concluída com sucesso!")
            
            self.progress.set(1)
            messagebox.showinfo("Sucesso", "Dados injetados, custos e colunas numéricas convertidas com sucesso!")

        except Exception as e:
            self.log(f"ERRO: {str(e)}")
            messagebox.showerror("Erro", f"Falha no processamento: {e}")
        finally:
            self.btn_run.configure(state="normal", text="EXECUTAR MÁQUINA PANDAS")
            self.progress.set(0)

if __name__ == "__main__":
    app = RunWorkElite()
    app.mainloop()
