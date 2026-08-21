import os
import re
import time
import shutil
import platform
import subprocess
import threading
from datetime import datetime

import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

URL_NOVA_CONTA = (
    "https://sebraecrm.lightning.force.com/lightning/o/Account/new"
    "?inContextOfRef=1.eyJ0eXBlIjoic3RhbmRhcmRfX29iamVjdFBhZ2UiLCJhdHRyaWJ1dGVzIjp7"
    "Im9iamVjdEFwaU5hbWUiOiJBY2NvdW50IiwiYWN0aW9uTmFtZSI6Imxpc3QifSwic3RhdGUiOnsiZmls"
    "dGVyTmFtZSI6Il9fUmVjZW50In19&count=3"
)

COLUNAS_CONSENTIMENTO = ["WPP", "ARE", "ARL", "ARM"]
CONSENTIMENTOS = [
    ("Autoriza Receber Whatsapp?", "WhatsApp"),
    ("Autoriza Receber Ligação", "Ligação"),
    ("Autoriza Receber Mensagem", "Mensagem"),
]


class BotExcelSalesforceApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Automação de Cadastros - Salesforce (FOCO)")
        self.geometry("900x680")
        self.minsize(850, 640)

        self.navegador_aberto = False
        self.excel_selecionado = False
        self.caminho_excel = ""
        self.driver = None
        self.ultima_origem = ""
        self.parar_solicitado = threading.Event()
        self.caminho_log = ""

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self.lbl_titulo = ctk.CTkLabel(self, text="Painel de Controle do Bot",
                                       font=ctk.CTkFont(size=22, weight="bold"))
        self.lbl_titulo.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="w")

        self.frame_conteudo = ctk.CTkFrame(self)
        self.frame_conteudo.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")
        self.frame_conteudo.grid_columnconfigure((0, 1), weight=1)
        self.frame_conteudo.grid_rowconfigure(0, weight=1)

        self.txt_instrucoes = ctk.CTkTextbox(self.frame_conteudo, wrap="word",
                                             font=ctk.CTkFont(size=13))
        self.txt_instrucoes.grid(row=0, column=0, padx=(10, 5), pady=10, sticky="nsew")
        self.carregar_instrucoes()

        self.txt_logs = ctk.CTkTextbox(self.frame_conteudo, wrap="word",
                                       font=ctk.CTkFont(family="Courier", size=12))
        self.txt_logs.grid(row=0, column=1, padx=(5, 10), pady=10, sticky="nsew")
        self.txt_logs.configure(state="disabled")
        self.log("[SISTEMA] Aguardando inicialização do ambiente...")

        self.frame_botoes = ctk.CTkFrame(self)
        self.frame_botoes.grid(row=2, column=0, padx=20, pady=(10, 20), sticky="ew")
        self.frame_botoes.grid_columnconfigure((0, 1, 2, 3), weight=1)

        self.btn_navegador = ctk.CTkButton(self.frame_botoes, text="1. Abrir Navegador",
                                           command=self.acao_abrir_navegador, height=45)
        self.btn_navegador.grid(row=0, column=0, padx=10, pady=15, sticky="ew")

        self.btn_selecionar_excel = ctk.CTkButton(self.frame_botoes, text="2. Selecionar Planilha",
                                                  command=self.acao_selecionar_excel, height=45)
        self.btn_selecionar_excel.grid(row=0, column=1, padx=10, pady=15, sticky="ew")

        self.btn_rodar = ctk.CTkButton(self.frame_botoes, text="3. Rodar Bot",
                                       command=self.acao_rodar_bot, state="disabled",
                                       fg_color="gray", height=45)
        self.btn_rodar.grid(row=0, column=2, padx=10, pady=15, sticky="ew")

        self.btn_parar = ctk.CTkButton(self.frame_botoes, text="Parar", command=self.acao_parar,
                                       state="disabled", fg_color="gray",
                                       hover_color="#8B0000", height=45)
        self.btn_parar.grid(row=0, column=3, padx=10, pady=15, sticky="ew")

        self.barra_progresso = ctk.CTkProgressBar(self.frame_botoes)
        self.barra_progresso.set(0)
        self.barra_progresso.grid(row=1, column=0, columnspan=4, padx=10, pady=(0, 5), sticky="ew")

        self.lbl_status_arquivo = ctk.CTkLabel(self.frame_botoes, text="Nenhum arquivo selecionado",
                                               font=ctk.CTkFont(size=11), text_color="gray")
        self.lbl_status_arquivo.grid(row=2, column=0, columnspan=4, pady=(0, 10))

        self.protocol("WM_DELETE_WINDOW", self.ao_fechar)

    # ------------------------------------------------------------------ UI --
    def carregar_instrucoes(self):
        instrucoes = (
            "=== INSTRUÇÕES DO SISTEMA ===\n\n"
            "Do que se trata o Bot?\n"
            "Automatiza a atualização cadastral de empresas (Telefone e Consentimentos) "
            "no Salesforce/FOCO, a partir de uma planilha Excel.\n\n"
            "Colunas da planilha:\n"
            "• CNPJ (obrigatória) : apenas números ou com pontuação.\n"
            "• Telefone (obrigatória) : com ou sem DDD; o 9º dígito é ajustado automaticamente.\n"
            "• WPP / ARE / ARL / ARM (opcionais) : esperado 'SIM'. Divergências são apenas\n"
            "  registradas no log; os consentimentos são sempre marcados como 'Sim'.\n"
            "• Tipo de telefone / Email / Canal preferencial : apenas copiadas para o relatório.\n\n"
            "Regras de origem:\n"
            "• Origem 'CPE' -> rotina específica (Pessoas Ocupadas + Telefone + consentimentos).\n"
            "• Qualquer outra origem -> rotina padrão, passando pela aba Contato.\n\n"
            "Regras de bloqueio:\n"
            "• Situação Cadastral RFB 'Suspensa' ou 'Baixada' -> registro é pulado.\n\n"
            "Passo a Passo:\n"
            "1. Clique em '1. Abrir Navegador'.\n"
            "2. Faça seu login no FOCO na janela que abrir.\n"
            "3. Clique em '2. Selecionar Planilha'.\n"
            "4. Clique em '3. Rodar Bot'.\n\n"
            "Use 'Parar' para encerrar com segurança ao fim do registro atual.\n"
            "O relatório é salvo na mesma pasta da planilha e atualizado durante a execução."
        )
        self.txt_instrucoes.insert("0.0", instrucoes)
        self.txt_instrucoes.configure(state="disabled")

    def log(self, mensagem):
        """Thread-safe: sempre agenda a escrita na thread principal do Tk."""
        self.after(0, self._log_ui, mensagem)

    def _log_ui(self, mensagem):
        carimbo = datetime.now().strftime("%H:%M:%S")
        self.txt_logs.configure(state="normal")
        self.txt_logs.insert("end", f"[{carimbo}] {mensagem}\n")
        self.txt_logs.see("end")
        self.txt_logs.configure(state="disabled")

    def ui(self, funcao, *args):
        """Executa qualquer alteração de widget na thread principal."""
        self.after(0, lambda: funcao(*args))

    def atualizar_progresso(self, atual, total):
        self.ui(self.barra_progresso.set, (atual / total) if total else 0)
        self.after(0, lambda: self.lbl_status_arquivo.configure(
            text=f"Processando {atual} de {total}...", text_color="gray"))

    def atualizar_estado_botao_rodar(self):
        if self.navegador_aberto and self.excel_selecionado:
            self.after(0, lambda: self.btn_rodar.configure(
                state="normal", fg_color=("#2b719e", "#1f538d")))
            self.log("[SISTEMA] Bot pronto para ser iniciado.")
        else:
            self.after(0, lambda: self.btn_rodar.configure(state="disabled", fg_color="gray"))

    # -------------------------------------------------------------- Chrome --
    def acao_abrir_navegador(self):
        threading.Thread(target=self._thread_iniciar_chrome, daemon=True).start()

    def _localizar_chrome(self):
        sistema = platform.system()
        if sistema == "Windows":
            candidatos = [
                r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
                os.path.join(os.environ.get("LOCALAPPDATA", ""),
                             r"Google\Chrome\Application\chrome.exe"),
            ]
            caminho = next((c for c in candidatos if c and os.path.isfile(c)), None)
            if not caminho:
                raise FileNotFoundError("Google Chrome não encontrado nos caminhos padrão do Windows.")
            return caminho, r"C:\ChromeDevSession"

        if sistema == "Linux":
            caminho = (shutil.which("google-chrome") or shutil.which("google-chrome-stable")
                       or shutil.which("chromium") or shutil.which("chromium-browser"))
            if not caminho:
                raise FileNotFoundError("Google Chrome/Chromium não foi encontrado.")
            return caminho, os.path.expanduser("~/ChromeDevSession")

        if sistema == "Darwin":
            caminho = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
            if not os.path.isfile(caminho):
                raise FileNotFoundError("Google Chrome não encontrado em /Applications.")
            return caminho, os.path.expanduser("~/ChromeDevSession")

        raise RuntimeError(f"Sistema operacional não suportado: {sistema}")

    def _thread_iniciar_chrome(self):
        self.log("[NAVEGADOR] Abrindo o Google Chrome em modo de Depuração...")
        try:
            chrome_cmd, user_dir = self._localizar_chrome()
            os.makedirs(user_dir, exist_ok=True)

            self.log(f"[NAVEGADOR] Executável: {chrome_cmd}")
            self.log(f"[NAVEGADOR] Perfil: {user_dir}")

            subprocess.Popen([chrome_cmd, "--remote-debugging-port=9222",
                              f"--user-data-dir={user_dir}"])

            self.navegador_aberto = True
            self.log("[NAVEGADOR] Chrome iniciado na porta 9222. Faça o login no FOCO.")
            self.atualizar_estado_botao_rodar()
        except Exception as e:
            self.log(f"[ERRO] Não foi possível iniciar o Chrome: {type(e).__name__}: {e}")

    # --------------------------------------------------------------- Excel --
    def acao_selecionar_excel(self):
        caminho = filedialog.askopenfilename(title="Selecione a base",
                                             filetypes=[("Excel", "*.xlsx *.xls")])
        if not caminho:
            return
        self.caminho_excel = caminho
        nome = os.path.basename(caminho)
        self.lbl_status_arquivo.configure(text=f"Arquivo: {nome}", text_color="green")
        self.log(f"[ARQUIVO] Planilha selecionada: {nome}")
        self.excel_selecionado = True
        self.atualizar_estado_botao_rodar()

    # ------------------------------------------------------------ Execução --
    def acao_rodar_bot(self):
        self.parar_solicitado.clear()
        self.btn_rodar.configure(state="disabled", text="Processando...")
        self.btn_selecionar_excel.configure(state="disabled")
        self.btn_navegador.configure(state="disabled")
        self.btn_parar.configure(state="normal", fg_color="#a83232")
        self.barra_progresso.set(0)
        threading.Thread(target=self._thread_loop_principal, daemon=True).start()

    def acao_parar(self):
        self.parar_solicitado.set()
        self.btn_parar.configure(state="disabled", text="Parando...")
        self.log("[SISTEMA] Parada solicitada. Encerrando após o registro atual...")

    def ao_fechar(self):
        self.parar_solicitado.set()
        self.destroy()

    def _salvar_log_parcial(self, resultados):
        try:
            pd.DataFrame(resultados).to_excel(self.caminho_log, index=False)
            return True
        except Exception as e:
            self.log(f"[AVISO] Não foi possível gravar o log agora ({type(e).__name__}). "
                     f"Feche o arquivo se ele estiver aberto no Excel.")
            return False

    def _thread_loop_principal(self):
        self.log("[EXECUÇÃO] Conectando ao navegador...")
        try:
            opts = Options()
            opts.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
            self.driver = webdriver.Chrome(options=opts)
            if self.driver.window_handles:
                self.driver.switch_to.window(self.driver.window_handles[0])
        except Exception as e:
            self.log(f"[ERRO] Não foi possível conectar ao Chrome: {type(e).__name__}: {e}")
            self.finalizar_processamento()
            return

        self.log("[EXECUÇÃO] Lendo a planilha...")
        try:
            df = pd.read_excel(self.caminho_excel, dtype=str)
        except Exception as e:
            self.log(f"[ERRO] Falha ao ler Excel: {type(e).__name__}: {e}")
            self.finalizar_processamento()
            return

        if "CNPJ" not in df.columns:
            self.log("[ERRO] A planilha não possui a coluna obrigatória 'CNPJ'.")
            self.after(0, lambda: messagebox.showerror(
                "Coluna ausente", "A planilha precisa conter a coluna 'CNPJ'."))
            self.finalizar_processamento()
            return

        pasta = os.path.dirname(os.path.abspath(self.caminho_excel))
        marca = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.caminho_log = os.path.join(pasta, f"Log_Final_{marca}.xlsx")
        self.log(f"[EXECUÇÃO] Relatório: {self.caminho_log}")

        resultados = []
        total = len(df)
        interrompido = False

        for index, row in df.iterrows():
            if self.parar_solicitado.is_set():
                interrompido = True
                self.log("[SISTEMA] Execução interrompida pelo usuário.")
                break

            posicao = len(resultados) + 1
            self.log(f"--- Registro {posicao} de {total} ---")
            self.atualizar_progresso(posicao, total)

            self.ultima_origem = ""
            status = self.processar_linha_salesforce(self.driver, row)

            resultados.append({
                "CNPJ": row.get("CNPJ", "N/A"),
                "Telefone": row.get("Telefone", "N/A"),
                "Tipo de telefone": row.get("Tipo de telefone", "N/A"),
                "Email": row.get("Email", "N/A"),
                "WPP": row.get("WPP", "N/A"),
                "ARE": row.get("ARE", "N/A"),
                "ARL": row.get("ARL", "N/A"),
                "ARM": row.get("ARM", "N/A"),
                "Canal preferencial": row.get("Canal preferencial", "N/A"),
                "Origem": self.ultima_origem or "Não identificada",
                "Status": status,
                "Processado em": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            })

            # Gravação incremental: nada se perde se o bot travar no meio do lote.
            if posicao % 5 == 0 or posicao == total:
                self._salvar_log_parcial(resultados)

        gravou = self._salvar_log_parcial(resultados) if resultados else False
        total_ok = sum(1 for r in resultados if r["Status"] == "Cadastrado com sucesso")
        total_pulado = sum(1 for r in resultados if str(r["Status"]).startswith("Pulado"))
        total_erro = len(resultados) - total_ok - total_pulado

        resumo = (f"Processados: {len(resultados)} de {total}\n"
                  f"Sucesso: {total_ok} | Pulados: {total_pulado} | Erros: {total_erro}")
        self.log(f"[FIM] {resumo.replace(chr(10), ' | ')}")

        if gravou:
            self.log(f"[FIM] Relatório salvo em: {self.caminho_log}")
            titulo = "Interrompido" if interrompido else "Sucesso"
            self.after(0, lambda: messagebox.showinfo(
                titulo, f"{resumo}\n\nRelatório:\n{self.caminho_log}"))
        elif resultados:
            self.after(0, lambda: messagebox.showerror(
                "Erro ao salvar", f"Não foi possível salvar:\n{self.caminho_log}\n\n"
                                  "Verifique se o arquivo está aberto ou sem permissão."))

        self.finalizar_processamento()

    def finalizar_processamento(self):
        def _restaurar():
            self.btn_selecionar_excel.configure(state="normal")
            self.btn_navegador.configure(state="normal")
            self.btn_rodar.configure(state="normal", text="3. Rodar Bot")
            self.btn_parar.configure(state="disabled", text="Parar", fg_color="gray")
            self.lbl_status_arquivo.configure(
                text=os.path.basename(self.caminho_excel) or "Nenhum arquivo selecionado",
                text_color="green" if self.caminho_excel else "gray")
        self.after(0, _restaurar)

    # ------------------------------------------------------- Tratamento in --
    @staticmethod
    def _limpar_numero(valor):
        if valor is None or (isinstance(valor, float) and pd.isna(valor)):
            return ""
        texto = str(valor).strip()
        if texto.lower() in ("nan", "none", "nat"):
            return ""
        if re.fullmatch(r"\d+\.0", texto):
            texto = texto[:-2]
        return "".join(filter(str.isdigit, texto))

    def _tratar_telefone(self, valor):
        numero = self._limpar_numero(valor)
        if len(numero) == 8:
            return "9" + numero
        if len(numero) == 10:
            return numero[:2] + "9" + numero[2:]
        return numero

    def _conferir_consentimentos(self, row_data):
        """Não altera o fluxo: apenas registra divergências para auditoria."""
        divergentes = [c for c in COLUNAS_CONSENTIMENTO
                       if c in row_data.index
                       and str(row_data.get(c, "")).strip().upper() not in ("SIM", "", "NAN", "NONE")]
        if divergentes:
            valores = ", ".join(f"{c}={row_data.get(c)}" for c in divergentes)
            self.log(f" [AUDITORIA] Consentimento fora do padrão na planilha ({valores}). "
                     f"Regra do processo: marcando 'Sim' mesmo assim.")

    # ------------------------------------------------- Helpers de Selenium --
    @staticmethod
    def _js_click(driver, elemento):
        driver.execute_script("arguments[0].click();", elemento)

    def _scroll_ate(self, driver, elemento):
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elemento)
        WebDriverWait(driver, 5).until(EC.visibility_of(elemento))
        return elemento

    def _digitar(self, driver, elemento, texto):
        self._scroll_ate(driver, elemento)
        self._js_click(driver, elemento)
        elemento.send_keys(Keys.CONTROL + "a")
        elemento.send_keys(Keys.BACKSPACE)
        elemento.send_keys(texto)

    def _definir_picklist_sim(self, driver, wait, aria_label, rotulo):
        """Substitui os 6 blocos duplicados de consentimento."""
        try:
            botao = wait.until(EC.presence_of_element_located(
                (By.XPATH, f"//button[@aria-label='{aria_label}']")))
        except TimeoutException:
            self.log(f" [AVISO] Campo '{rotulo}' não encontrado nesta tela.")
            return False

        self._scroll_ate(driver, botao)
        if botao.get_attribute("data-value") == "Sim":
            self.log(f" [OK] {rotulo} já está como 'Sim'.")
            return True

        self.log(f" [AÇÃO] {rotulo} não está como 'Sim'. Alterando...")
        self._js_click(driver, botao)
        opcoes = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located(
            (By.XPATH, "//lightning-base-combobox-item[@data-value='Sim']")))
        for opcao in opcoes:
            if opcao.is_displayed():
                self._js_click(driver, opcao)
                WebDriverWait(driver, 10).until(
                    lambda d: botao.get_attribute("data-value") == "Sim")
                self.log(f" [OK] {rotulo} marcado como 'Sim'.")
                return True
        self.log(f" [AVISO] Não achei a opção 'Sim' visível para {rotulo}.")
        return False

    def _preencher_telefone(self, driver, wait, telefone, contexto):
        if not telefone:
            return
        campo = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@name='Phone']")))
        self._digitar(driver, campo, telefone)
        self.log(f" [OK] Telefone preenchido ({contexto}).")

    def _verificar_salvamento(self, driver):
        """Confirma o salvamento em vez de presumir sucesso após um sleep fixo."""
        fim = time.time() + 25
        xpath_erro_campo = "//div[contains(@class,'slds-form-element__help')]"
        xpath_erro_topo = ("//div[contains(@class,'slds-notify_toast') and "
                           "contains(@class,'slds-theme_error')] | "
                           "//ul[contains(@class,'errorsList')]//li")
        xpath_sucesso = ("//div[contains(@class,'slds-notify_toast') and "
                         "contains(@class,'slds-theme_success')]")
        xpath_modal = "//div[contains(@class,'slds-modal') and contains(@class,'slds-fade-in-open')]"

        while time.time() < fim:
            for xp in (xpath_erro_topo, xpath_erro_campo):
                erros = [e.text.strip() for e in driver.find_elements(By.XPATH, xp)
                         if e.is_displayed() and e.text.strip()]
                if erros:
                    return False, " | ".join(dict.fromkeys(erros))[:300]
            if driver.find_elements(By.XPATH, xpath_sucesso):
                return True, "toast de sucesso"
            if not driver.find_elements(By.XPATH, xpath_modal):
                return True, "modal fechado"
            time.sleep(0.5)
        return False, "tempo esgotado aguardando confirmação do salvamento"

    # ------------------------------------------------- Processamento linha --
    def processar_linha_salesforce(self, driver, row_data):
        wait = WebDriverWait(driver, 20)

        cnpj_numeros = self._limpar_numero(row_data.get("CNPJ"))
        cnpj = cnpj_numeros.zfill(14) if cnpj_numeros else ""
        telefone = self._tratar_telefone(row_data.get("Telefone"))

        if not cnpj:
            self.log(" [ERRO] CNPJ vazio ou inválido na planilha.")
            return "Erro - CNPJ inválido"

        self.log(f" -> CNPJ tratado: {cnpj} | Telefone tratado: {telefone or '(vazio)'}")
        self._conferir_consentimentos(row_data)

        try:
            driver.get(URL_NOVA_CONTA)

            # --- Buscar pelo CNPJ ---
            select_el = wait.until(EC.presence_of_element_located(
                (By.XPATH, "//select[@name='search']")))
            Select(select_el).select_by_value("CNPJ")

            input_cnpj = wait.until(EC.presence_of_element_located(
                (By.XPATH, "//label[text()='CNPJ']/following::div/input[@class='slds-input']")))
            self._digitar(driver, input_cnpj, cnpj)

            btn_buscar = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(@class,'slds-button_brand')][text()='Buscar']")))
            self._js_click(driver, btn_buscar)

            # --- Resultado: origem e radio lidos da MESMA linha ---
            linha = wait.until(EC.presence_of_element_located(
                (By.XPATH, "//tr[.//input[@type='radio' and contains(@name,'options')]]")))
            self._scroll_ate(driver, linha)

            origem_texto = ""
            try:
                celulas = linha.find_elements(By.XPATH, ".//lightning-base-formatted-text")
                for celula in celulas:
                    texto = " ".join(celula.text.split()).strip()
                    if "CPE" in texto.upper() or "SALESFORCE" in texto.upper():
                        origem_texto = texto
                        break
                if not origem_texto and celulas:
                    origem_texto = " ".join(celulas[-1].text.split()).strip()
            except Exception as e:
                self.log(f" [AVISO] Falha ao ler a origem ({type(e).__name__}).")

            self.ultima_origem = origem_texto
            self.log(f" [INFO] Origem identificada: {origem_texto or 'não identificada'}")

            radio = linha.find_element(
                By.XPATH, ".//input[@type='radio' and contains(@name,'options')]")
            self._js_click(driver, radio)

            btn_continuar = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(@class,'slds-button_brand') and "
                           "contains(.,'Continuar')]")))
            self._js_click(driver, btn_continuar)

            # --- Situação Cadastral RFB ---
            try:
                self.log(" [AÇÃO] Verificando Situação Cadastral RFB...")
                situacao_el = WebDriverWait(driver, 8).until(EC.presence_of_element_located(
                    (By.XPATH, "//span[text()='Situação Cadastral RFB']"
                               "/ancestor::div[contains(@class,'slds-form-element')]"
                               "//force-record-output-picklist")))
                self._scroll_ate(driver, situacao_el)
                situacao = situacao_el.text.strip()
                self.log(f" [INFO] Situação Cadastral RFB: {situacao or '(vazia)'}")

                if situacao in ("Suspensa", "Baixada"):
                    self.log(f" [PULO] CNPJ ignorado por situação RFB: {situacao}")
                    try:
                        btn_cancelar = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                            (By.XPATH, "//button[@name='CancelEdit' or @title='Cancelar']")))
                        self._js_click(driver, btn_cancelar)
                    except TimeoutException:
                        self.log(" [AVISO] Botão Cancelar não encontrado; seguindo assim mesmo.")
                    return f"Pulado - RFB {situacao}"
            except TimeoutException:
                self.log(" [AVISO] Campo Situação Cadastral não visível. Seguindo o fluxo...")

            # --- Rotina por origem: CPE tem tratamento próprio; o resto usa a aba Contato ---
            if origem_texto.strip().upper() == "CPE":
                self.log(" [INFO] Origem CPE -> rotina específica.")

                campo_pessoas = wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//input[@name='PessoasOcupadas__c']")))
                self._scroll_ate(driver, campo_pessoas)
                atual = (campo_pessoas.get_attribute("value") or "").strip()
                if not atual:
                    self.log(" [AÇÃO] PessoasOcupadas vazio. Preenchendo com 1.")
                    self._js_click(driver, campo_pessoas)
                    campo_pessoas.send_keys("1")
                else:
                    self.log(f" [OK] PessoasOcupadas já preenchido: {atual}")

                self._preencher_telefone(driver, wait, telefone, "formulário CPE")
            else:
                self.log(f" [INFO] Origem '{origem_texto or 'não identificada'}' "
                         f"-> rotina padrão (aba Contato).")
                aba = wait.until(EC.element_to_be_clickable(
                    (By.XPATH, "//a[@data-label='Contato' or text()='Contato']")))
                self._js_click(driver, aba)
                self._preencher_telefone(driver, wait, telefone, "aba Contato")

            # Consentimentos: mesmos três campos nas duas rotinas.
            for aria_label, rotulo in CONSENTIMENTOS:
                self._definir_picklist_sim(driver, wait, aria_label, rotulo)

            # --- Salvar e CONFERIR o resultado ---
            self.log(" [AÇÃO] Salvando o registro...")
            btn_salvar = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//button[@name='SaveEdit' or @title='Salvar' or text()='Salvar']")))
            self._js_click(driver, btn_salvar)

            ok, detalhe = self._verificar_salvamento(driver)
            if ok:
                self.log(f" [OK] Cadastro salvo ({detalhe}).")
                return "Cadastrado com sucesso"

            self.log(f" [ERRO] Salvamento não confirmado: {detalhe}")
            return f"Erro ao salvar - {detalhe}"

        except TimeoutException as e:
            alvo = str(e).splitlines()[0][:160] if str(e) else "elemento não encontrado"
            self.log(f" [ERRO] Timeout: {alvo}")
            return "Erro - Timeout, necessita de análise humana"
        except Exception as e:
            self.log(f" [ERRO] Falha na execução: {type(e).__name__}: {str(e)[:200]}")
            return "Erro - Necessita de análise humana"


if __name__ == "__main__":
    app = BotExcelSalesforceApp()
    app.mainloop()