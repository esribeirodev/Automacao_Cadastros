import os
import sys
import time
import subprocess
import platform
import shutil
import subprocess
import threading
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select

# Configurações de aparência do CustomTkinter
ctk.set_appearance_mode("System")  # Segue o modo do Windows (Dark/Light)
ctk.set_default_color_theme("blue")

class BotExcelSalesforceApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Configurações da Janela
        self.title("Automação de Cadastros - Salesforce (FOCO)")
        self.geometry("850x650")
        self.minsize(800, 600)

        # Variáveis de Controle de Estado
        self.navegador_aberto = False
        self.excel_selecionado = False
        self.caminho_excel = ""
        self.driver = None

        # --- ESTRUTURA LAYOUT (Grid) ---
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # 1. Cabeçalho / Título
        self.lbl_titulo = ctk.CTkLabel(self, text="Painel de Controle do Bot", font=ctk.CTkFont(size=22, weight="bold"))
        self.lbl_titulo.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="w")

        # 2. Área de Instruções e Logs
        self.frame_conteudo = ctk.CTkFrame(self)
        self.frame_conteudo.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")
        self.frame_conteudo.grid_columnconfigure(0, weight=1)
        self.frame_conteudo.grid_columnconfigure(1, weight=1)
        self.frame_conteudo.grid_rowconfigure(0, weight=1)

        self.txt_instrucoes = ctk.CTkTextbox(self.frame_conteudo, wrap="word", font=ctk.CTkFont(size=13))
        self.txt_instrucoes.grid(row=0, column=0, padx=(10, 5), pady=10, sticky="nsew")
        self.carregar_instrucoes()

        self.txt_logs = ctk.CTkTextbox(self.frame_conteudo, wrap="word", font=ctk.CTkFont(family="Courier", size=12))
        self.txt_logs.grid(row=0, column=1, padx=(5, 10), pady=10, sticky="nsew")
        self.log("[SISTEMA] Aguardando inicialização do ambiente...")

        # 3. Frame de Controles
        self.frame_botoes = ctk.CTkFrame(self, height=120)
        self.frame_botoes.grid(row=2, column=0, padx=20, pady=(10, 20), sticky="ew")
        self.frame_botoes.grid_columnconfigure((0, 1, 2), weight=1)

        self.btn_navegador = ctk.CTkButton(self.frame_botoes, text="1. Abrir Navegador", command=self.acao_abrir_navegador, height=45)
        self.btn_navegador.grid(row=0, column=0, padx=10, pady=15, sticky="ew")

        self.btn_selecionar_excel = ctk.CTkButton(self.frame_botoes, text="2. Selecionar Planilha", command=self.acao_selecionar_excel, height=45)
        self.btn_selecionar_excel.grid(row=0, column=1, padx=10, pady=15, sticky="ew")

        self.btn_rodar = ctk.CTkButton(self.frame_botoes, text="3. Rodar Bot", command=self.acao_rodar_bot, state="disabled", fg_color="gray", height=45)
        self.btn_rodar.grid(row=0, column=2, padx=10, pady=15, sticky="ew")

        self.lbl_status_arquivo = ctk.CTkLabel(self.frame_botoes, text="Nenhum arquivo selecionado", font=ctk.CTkFont(size=11), text_color="gray")
        self.lbl_status_arquivo.grid(row=1, column=0, columnspan=3, pady=(0, 10))

    def carregar_instrucoes(self):
        instrucoes = (
            "=== INSTRUÇÕES DO SISTEMA ===\n\n"
            "Do que se trata o Bot?\n"
            "Este automatizador atualiza dados cadastrais (E-mail, Telefone, Canal Preferencial, Endereço e Consentimentos) no Salesforce.\n\n"
            "Colunas obrigatórias no Excel:\n"
            "• CPF : Apenas números ou com pontos e traço.\n"
            "• Telefone : Número do telefone.\n"
            "• Tipo de telefone : (ex: 'Telefone Celular', 'Telefone Comercial').\n"
            "• Email : Endereço de e-mail.\n"
            "• WPP / ARE / ARL / ARM : 'SIM' para autorizar.\n"
            "• Canal preferencial : Opções (ex: 'Email', 'WhatsApp', 'SMS').\n\n"
            "Passo a Passo:\n"
            "1. Clique em '1. Abrir Navegador'.\n"
            "2. Faça seu Login no FOCO na janela que abrir.\n"
            "3. Clique em '2. Selecionar Planilha'.\n"
            "4. Clique em '3. Rodar Bot'."
        )
        self.txt_instrucoes.insert("0.0", instrucoes)
        self.txt_instrucoes.configure(state="disabled")

    def log(self, messaging):
        self.txt_logs.configure(state="normal")
        self.txt_logs.insert("end", f"{messaging}\n")
        self.txt_logs.see("end")
        self.txt_logs.configure(state="disabled")

    def atualizar_estado_botao_rodar(self):
        if self.navegador_aberto and self.excel_selecionado:
            self.btn_rodar.configure(state="normal", fg_color=("#2b719e", "#1f538d"))
            self.log("[SISTEMA] Bot pronto para ser iniciado.")
        else:
            self.btn_rodar.configure(state="disabled", fg_color="gray")

    def acao_abrir_navegador(self):
        threading.Thread(target=self._thread_iniciar_chrome, daemon=True).start()

    def _thread_iniciar_chrome(self):
            self.log("[NAVEGADOR] Abrindo o Google Chrome em modo de Depuração...")
            try:
                sistema = platform.system()

                if sistema == "Windows":
                    chrome_cmd = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
                    user_dir = r"C:\ChromeDevSession"

                elif sistema == "Linux":
                    chrome_cmd = (
                        shutil.which("google-chrome")
                        or shutil.which("google-chrome-stable")
                        or shutil.which("chromium")
                        or shutil.which("chromium-browser")
                    )

                    if not chrome_cmd:
                        raise FileNotFoundError(
                            "Google Chrome/Chromium não foi encontrado."
                        )

                    user_dir = os.path.expanduser("~/ChromeDevSession")

                else:
                    raise RuntimeError(
                        f"Sistema operacional não suportado: {sistema}"
                    )

                os.makedirs(user_dir, exist_ok=True)

                cmd = [
                    chrome_cmd,
                    "--remote-debugging-port=9222",
                    f"--user-data-dir={user_dir}",
                ]

                self.log(f"[NAVEGADOR] Executável: {chrome_cmd}")
                self.log(f"[NAVEGADOR] Perfil: {user_dir}")

                subprocess.Popen(cmd)

                self.navegador_aberto = True
                self.log("[NAVEGADOR] Chrome iniciado na porta 9222.")

                self.atualizar_estado_botao_rodar()

            except Exception as e:
                self.log(f"[ERRO] Não foi possível iniciar o Chrome: {e}")

    def acao_selecionar_excel(self):
        caminho = filedialog.askopenfilename(title="Selecione a base", filetypes=[("Excel", "*.xlsx *.xls")])
        if caminho:
            self.caminho_excel = caminho
            nome_arquivo = os.path.basename(caminho)
            self.lbl_status_arquivo.configure(text=f"Arquivo: {nome_arquivo}", text_color="green")
            self.log(f"[ARQUIVO] Planilha selecionada: {nome_arquivo}")
            self.excel_selecionado = True
            self.atualizar_estado_botao_rodar()

    def acao_rodar_bot(self):
        self.btn_rodar.configure(state="disabled", text="Processando...")
        self.btn_selecionar_excel.configure(state="disabled")
        self.btn_navegador.configure(state="disabled")
        threading.Thread(target=self._thread_loop_principal, daemon=True).start()

    def _thread_loop_principal(self):
        self.log("\n[EXECUÇÃO] Conectando ao navegador...")
        
        try:
            opts = Options()
            opts.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
            self.driver = webdriver.Chrome(options=opts)
            
            if len(self.driver.window_handles) > 0:
                self.driver.switch_to.window(self.driver.window_handles[0])
        except Exception as e:
            self.log(f"[ERRO] Não foi possível conectar ao Chrome: {e}")
            self.finalizar_processamento()
            return

        self.log("[EXECUÇÃO] Lendo a planilha...")
        try:
            df = pd.read_excel(self.caminho_excel, dtype=str)
        except Exception as e:
            self.log(f"[ERRO] Falha ao ler Excel: {e}")
            self.finalizar_processamento()
            return

        resultados = []
        total_linhas = len(df)

        for index, row in df.iterrows():
            self.log(f"--- Registro {index + 1} de {total_linhas} ---")
            status = self.processar_linha_salesforce(self.driver, row)
            
            resultados.append({
                "CNPJ": row.get('CNPJ', 'N/A'), 
                "Telefone": row.get('Telefone', 'N/A'),
                "Tipo de telefone": row.get('Tipo de telefone', 'N/A'),
                "Email": row.get('Email', 'N/A'),
                "WPP": row.get('WPP', 'N/A'),
                "ARE": row.get('ARE', 'N/A'),
                "ARL": row.get('ARL', 'N/A'),
                "ARM": row.get('ARM', 'N/A'),
                "Canal preferencial": row.get('Canal preferencial', 'N/A'),
                "Status": status
            })

        try:
            pd.DataFrame(resultados).to_excel("Log_Final.xlsx", index=False)
            self.log("\n[FIM] Relatório 'Log_Final.xlsx' gerado.")
            messagebox.showinfo("Sucesso", "Processamento finalizado!")
        except Exception as e:
            self.log(f"[ERRO] Falha ao salvar Log: {e}")
            messagebox.showerror("Erro de Permissão", "Não foi possível salvar o Log_Final.xlsx.")
            
        self.finalizar_processamento()

    def finalizar_processamento(self):
        self.btn_selecionar_excel.configure(state="normal")
        self.btn_navegador.configure(state="normal")
        self.btn_rodar.configure(state="normal", text="3. Rodar Bot")

    def processar_linha_salesforce(self, driver, row_data):
            wait = WebDriverWait(driver, 20)
            
            # Extração do CNPJ
            cnpj_bruto = str(row_data.get('CNPJ', '')).strip() if pd.notna(row_data.get('CNPJ')) else ""
            if cnpj_bruto.endswith(".0") and cnpj_bruto.count(".") == 1:
                cnpj_bruto = cnpj_bruto[:-2]

            cnpj_so_numeros = ''.join(filter(str.isdigit, cnpj_bruto))
            cnpj = cnpj_so_numeros.zfill(14) if cnpj_so_numeros else ""

            # Tratamento do Telefone
            telefone_bruto = str(row_data.get('Telefone', '')).strip() if pd.notna(row_data.get('Telefone')) else ""
            telefone_bruto = telefone_bruto[:-2] if telefone_bruto.endswith(".0") else telefone_bruto
            # Deixa apenas os números
            telefone_limpo = ''.join(filter(str.isdigit, telefone_bruto)) 
            
            # Validação do 9 dígito
            if len(telefone_limpo) == 8:
                # Se tem 8 dígitos (sem DDD), coloca o 9 na frente
                telefone_limpo = '9' + telefone_limpo
            elif len(telefone_limpo) == 10:
                # Se tem 10 dígitos (DDD + 8 dígitos), insere o 9 logo após o DDD
                telefone_limpo = telefone_limpo[:2] + '9' + telefone_limpo[2:]

            self.log(f" -> CNPJ Tratado: {cnpj} | Tel Tratado: {telefone_limpo}")

            try:
                def js_click(elm): driver.execute_script("arguments[0].click();", elm)
                def scroll_ate(elm): driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elm); time.sleep(0.5)

                driver.get("https://sebraecrm.lightning.force.com/lightning/o/Account/new?inContextOfRef=1.eyJ0eXBlIjoic3RhbmRhcmRfX29iamVjdFBhZ2UiLCJhdHRyaWJ1dGVzIjp7Im9iamVjdEFwaU5hbWUiOiJBY2NvdW50IiwiYWN0aW9uTmFtZSI6Imxpc3QifSwic3RhdGUiOnsiZmlsdGVyTmFtZSI6Il9fUmVjZW50In19&count=3")
                
                # --- ETAPA: SELECIONAR CNPJ E BUSCAR ---
                xpath_select = "//select[@name='search']"
                select_elemento = wait.until(EC.presence_of_element_located((By.XPATH, xpath_select)))
                dropdown = Select(select_elemento)
                dropdown.select_by_value("CNPJ")
                time.sleep(1) 
                
                xpath_input_cnpj = "//label[text()='CNPJ']/following::div/input[@class='slds-input']"
                input_cnpj = wait.until(EC.presence_of_element_located((By.XPATH, xpath_input_cnpj)))
                scroll_ate(input_cnpj)
                js_click(input_cnpj)
                input_cnpj.send_keys(Keys.CONTROL + "a")
                input_cnpj.send_keys(Keys.BACKSPACE)
                input_cnpj.send_keys(cnpj)
                time.sleep(0.5)

                xpath_btn_buscar_modal = "//button[contains(@class, 'slds-button_brand')][text()='Buscar']"
                js_click(wait.until(EC.presence_of_element_located((By.XPATH, xpath_btn_buscar_modal))))
                time.sleep(3) 

                # --- NOVA ETAPA: VERIFICAR ORIGEM (CPE vs CPE/Salesforce) ---
                origem_texto = ""
                try:
                    # Captura o texto que diz se é CPE ou CPE/Salesforce na tabela de resultados
                    xpath_origem = "//lightning-base-formatted-text[contains(., 'CPE')]"
                    elemento_origem = wait.until(EC.presence_of_element_located((By.XPATH, xpath_origem)))
                    origem_texto = elemento_origem.text.strip()
                    self.log(f" [INFO] Origem identificada: {origem_texto}")
                except:
                    self.log(" [AVISO] Não foi possível ler a origem na tabela.")

                # --- ETAPA: SELECIONAR RESULTADO E CONTINUAR ---
                xpath_radio_resultado = "//input[@type='radio' and contains(@name, 'options')]"
                radio_resultado = wait.until(EC.presence_of_element_located((By.XPATH, xpath_radio_resultado)))
                js_click(radio_resultado)
                time.sleep(1) 
                
                xpath_btn_continuar = "//button[contains(@class, 'slds-button_brand') and contains(., 'Continuar')]"
                btn_continuar = wait.until(EC.presence_of_element_located((By.XPATH, xpath_btn_continuar)))
                js_click(btn_continuar)
                time.sleep(3) 

                # --- NOVA ETAPA: TRATAMENTOS DE FORMULÁRIO ---
                
                if origem_texto == "CPE":
                    self.log(" [INFO] Executando rotina para Origem: CPE")
                    
                    # 1. Tratamento Pessoas Ocupadas
                    xpath_pessoas = "//input[@name='PessoasOcupadas__c']"
                    input_pessoas = wait.until(EC.presence_of_element_located((By.XPATH, xpath_pessoas)))
                    scroll_ate(input_pessoas)
                    
                    # Verifica se o campo está vazio capturando o valor nativo dele
                    valor_atual_pessoas = input_pessoas.get_attribute("value")
                    if not valor_atual_pessoas or str(valor_atual_pessoas).strip() == "":
                        self.log(" [AÇÃO] Campo PessoasOcupadas vazio. Preenchendo com 1.")
                        js_click(input_pessoas)
                        input_pessoas.send_keys("1")
                    else:
                        self.log(f" [AÇÃO] Campo PessoasOcupadas já preenchido com: {valor_atual_pessoas}")
                    
                    time.sleep(0.5)

                    # 2. Tratamento do Telefone
                    if telefone_limpo:
                        xpath_telefone = "//input[@name='Phone']"
                        input_telefone = wait.until(EC.presence_of_element_located((By.XPATH, xpath_telefone)))
                        scroll_ate(input_telefone)
                        js_click(input_telefone)
                        input_telefone.send_keys(Keys.CONTROL + "a")
                        input_telefone.send_keys(Keys.BACKSPACE)
                        input_telefone.send_keys(telefone_limpo)
                        self.log(" [AÇÃO] Telefone preenchido.")
                        time.sleep(1)

                elif origem_texto == "CPE/Salesforce":
                    self.log(" [INFO] Executando rotina para Origem: CPE/Salesforce")
                    # (Aqui entrará o seu tratamento futuro para esta condição)
                    pass

                self.log(" [OK] Teste finalizado. Parando robô sem salvar.")
                return "Teste executado com sucesso"

            except Exception as e:
                self.log(f" [ERRO] Falha na execução: {e}")
                return "Erro - Necessita de análise humana"

if __name__ == "__main__":
    app = BotExcelSalesforceApp()
    app.mainloop()