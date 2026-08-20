import os
import sys
import time
import subprocess
import platform
import shutil
import socket
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

from selenium.common.exceptions import (
    TimeoutException,
    StaleElementReferenceException,
    ElementClickInterceptedException,
    WebDriverException
)


# ============================================================
# CONFIGURAÇÕES
# ============================================================

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

CHROME_DEBUG_PORT = 9222

# Wait padrão
WAIT_PADRAO = 10

# Wait curto para coisas opcionais
WAIT_CURTO = 2

# Intervalo de verificação do Selenium
POLL_FREQUENCY = 0.1

URL_NOVO_CADASTRO = (
    "https://sebraecrm.lightning.force.com/lightning/o/Account/new"
    "?inContextOfRef=1.eyJ0eXBlIjoic3RhbmRhcmRfX29iamVjdFBhZ2UiLCJhdHRyaWJ1dGVzIjp7"
    "Im9iamVjdEFwaU5hbWUiOiJBY2NvdW50IiwiYWN0aW9uTmFtZSI6Imxpc3QifSwic3RhdGUi"
    "OnsiZmlsdGVyTmFtZSI6Il9fUmVjZW50In19&count=3"
)


# ============================================================
# FUNÇÃO: AGUARDAR CHROME DEBUGGER
# ============================================================

def esperar_chrome_debugger(
    porta=CHROME_DEBUG_PORT,
    timeout=15
):
    """
    Aguarda o Chrome realmente disponibilizar
    a porta de depuração.
    """

    inicio = time.time()

    while time.time() - inicio < timeout:

        try:

            with socket.create_connection(
                ("127.0.0.1", porta),
                timeout=0.2
            ):
                return True

        except OSError:

            time.sleep(0.1)

    return False


# ============================================================
# APLICAÇÃO
# ============================================================

class BotExcelSalesforceApp(ctk.CTk):

    def __init__(self):
        super().__init__()

        # ----------------------------------------------------
        # JANELA
        # ----------------------------------------------------

        self.title(
            "Automação de Cadastros - Salesforce (FOCO)"
        )

        self.geometry(
            "850x650"
        )

        self.minsize(
            800,
            600
        )

        # ----------------------------------------------------
        # ESTADOS
        # ----------------------------------------------------

        self.navegador_aberto = False
        self.excel_selecionado = False

        self.caminho_excel = ""
        self.driver = None

        # ----------------------------------------------------
        # GRID
        # ----------------------------------------------------

        self.grid_columnconfigure(
            0,
            weight=1
        )

        self.grid_rowconfigure(
            1,
            weight=1
        )

        # ----------------------------------------------------
        # TÍTULO
        # ----------------------------------------------------

        self.lbl_titulo = ctk.CTkLabel(
            self,
            text="Painel de Controle do Bot",
            font=ctk.CTkFont(
                size=22,
                weight="bold"
            )
        )

        self.lbl_titulo.grid(
            row=0,
            column=0,
            padx=20,
            pady=(20, 10),
            sticky="w"
        )

        # ----------------------------------------------------
        # CONTEÚDO
        # ----------------------------------------------------

        self.frame_conteudo = ctk.CTkFrame(
            self
        )

        self.frame_conteudo.grid(
            row=1,
            column=0,
            padx=20,
            pady=10,
            sticky="nsew"
        )

        self.frame_conteudo.grid_columnconfigure(
            0,
            weight=1
        )

        self.frame_conteudo.grid_columnconfigure(
            1,
            weight=1
        )

        self.frame_conteudo.grid_rowconfigure(
            0,
            weight=1
        )

        # ----------------------------------------------------
        # INSTRUÇÕES
        # ----------------------------------------------------

        self.txt_instrucoes = ctk.CTkTextbox(
            self.frame_conteudo,
            wrap="word",
            font=ctk.CTkFont(
                size=13
            )
        )

        self.txt_instrucoes.grid(
            row=0,
            column=0,
            padx=(10, 5),
            pady=10,
            sticky="nsew"
        )

        self.carregar_instrucoes()

        # ----------------------------------------------------
        # LOG
        # ----------------------------------------------------

        self.txt_logs = ctk.CTkTextbox(
            self.frame_conteudo,
            wrap="word",
            font=ctk.CTkFont(
                family="Courier",
                size=12
            )
        )

        self.txt_logs.grid(
            row=0,
            column=1,
            padx=(5, 10),
            pady=10,
            sticky="nsew"
        )

        self.log(
            "[SISTEMA] Aguardando inicialização do ambiente..."
        )

        # ----------------------------------------------------
        # BOTÕES
        # ----------------------------------------------------

        self.frame_botoes = ctk.CTkFrame(
            self,
            height=120
        )

        self.frame_botoes.grid(
            row=2,
            column=0,
            padx=20,
            pady=(10, 20),
            sticky="ew"
        )

        self.frame_botoes.grid_columnconfigure(
            (0, 1, 2),
            weight=1
        )

        # ----------------------------------------------------
        # BOTÃO CHROME
        # ----------------------------------------------------

        self.btn_navegador = ctk.CTkButton(
            self.frame_botoes,
            text="1. Abrir Navegador",
            command=self.acao_abrir_navegador,
            height=45
        )

        self.btn_navegador.grid(
            row=0,
            column=0,
            padx=10,
            pady=15,
            sticky="ew"
        )

        # ----------------------------------------------------
        # BOTÃO EXCEL
        # ----------------------------------------------------

        self.btn_selecionar_excel = ctk.CTkButton(
            self.frame_botoes,
            text="2. Selecionar Planilha",
            command=self.acao_selecionar_excel,
            height=45
        )

        self.btn_selecionar_excel.grid(
            row=0,
            column=1,
            padx=10,
            pady=15,
            sticky="ew"
        )

        # ----------------------------------------------------
        # BOTÃO RODAR
        # ----------------------------------------------------

        self.btn_rodar = ctk.CTkButton(
            self.frame_botoes,
            text="3. Rodar Bot",
            command=self.acao_rodar_bot,
            state="disabled",
            fg_color="gray",
            height=45
        )

        self.btn_rodar.grid(
            row=0,
            column=2,
            padx=10,
            pady=15,
            sticky="ew"
        )

        # ----------------------------------------------------
        # STATUS ARQUIVO
        # ----------------------------------------------------

        self.lbl_status_arquivo = ctk.CTkLabel(
            self.frame_botoes,
            text="Nenhum arquivo selecionado",
            font=ctk.CTkFont(
                size=11
            ),
            text_color="gray"
        )

        self.lbl_status_arquivo.grid(
            row=1,
            column=0,
            columnspan=3,
            pady=(0, 10)
        )

    # ========================================================
    # INSTRUÇÕES
    # ========================================================

    def carregar_instrucoes(self):

        instrucoes = (
            "=== INSTRUÇÕES DO SISTEMA ===\n\n"

            "Do que se trata o Bot?\n"
            "Este automatizador atualiza dados cadastrais "
            "(E-mail, Telefone, Canal Preferencial, Endereço "
            "e Consentimentos) no Salesforce.\n\n"

            "Colunas obrigatórias no Excel:\n"
            "• CNPJ : Apenas números ou com pontos e traço.\n"
            "• Telefone : Número do telefone.\n"
            "• Tipo de telefone : "
            "(ex: 'Telefone Celular', 'Telefone Comercial').\n"
            "• Email : Endereço de e-mail.\n"
            "• WPP / ARE / ARL / ARM : 'SIM' para autorizar.\n"
            "• Canal preferencial : "
            "Opções (ex: 'Email', 'WhatsApp', 'SMS').\n\n"

            "Passo a Passo:\n"
            "1. Clique em '1. Abrir Navegador'.\n"
            "2. Faça seu Login no FOCO na janela que abrir.\n"
            "3. Clique em '2. Selecionar Planilha'.\n"
            "4. Clique em '3. Rodar Bot'."
        )

        self.txt_instrucoes.insert(
            "0.0",
            instrucoes
        )

        self.txt_instrucoes.configure(
            state="disabled"
        )

    # ========================================================
    # LOG
    # ========================================================

    def log(self, mensagem):

        try:

            self.after(
                0,
                self._adicionar_log,
                mensagem
            )

        except Exception:
            pass

    def _adicionar_log(self, mensagem):

        try:

            self.txt_logs.configure(
                state="normal"
            )

            self.txt_logs.insert(
                "end",
                f"{mensagem}\n"
            )

            self.txt_logs.see(
                "end"
            )

            self.txt_logs.configure(
                state="disabled"
            )

        except Exception:
            pass

    # ========================================================
    # ESTADO BOTÃO RODAR
    # ========================================================

    def atualizar_estado_botao_rodar(self):

        if (
            self.navegador_aberto
            and self.excel_selecionado
        ):

            self.btn_rodar.configure(
                state="normal",
                fg_color=(
                    "#2b719e",
                    "#1f538d"
                )
            )

            self.log(
                "[SISTEMA] Bot pronto para ser iniciado."
            )

        else:

            self.btn_rodar.configure(
                state="disabled",
                fg_color="gray"
            )

    # ========================================================
    # ABRIR NAVEGADOR
    # ========================================================

    def acao_abrir_navegador(self):

        self.btn_navegador.configure(
            state="disabled",
            text="Abrindo..."
        )

        threading.Thread(
            target=self._thread_iniciar_chrome,
            daemon=True
        ).start()

    def _thread_iniciar_chrome(self):

        self.log(
            "[NAVEGADOR] Abrindo o Google Chrome "
            "em modo de Depuração..."
        )

        try:

            sistema = platform.system()

            # ------------------------------------------------
            # WINDOWS
            # ------------------------------------------------

            if sistema == "Windows":

                chrome_cmd = (
                    r"C:\Program Files\Google\Chrome\Application\chrome.exe"
                )

                if not os.path.exists(
                    chrome_cmd
                ):

                    chrome_cmd = (
                        r"C:\Program Files (x86)"
                        r"\Google\Chrome\Application"
                        r"\chrome.exe"
                    )

                if not os.path.exists(
                    chrome_cmd
                ):

                    raise FileNotFoundError(
                        "Google Chrome não encontrado."
                    )

                user_dir = (
                    r"C:\ChromeDevSession"
                )

            # ------------------------------------------------
            # LINUX
            # ------------------------------------------------

            elif sistema == "Linux":

                chrome_cmd = (
                    shutil.which("google-chrome")
                    or shutil.which("google-chrome-stable")
                    or shutil.which("chromium")
                    or shutil.which("chromium-browser")
                )

                if not chrome_cmd:

                    raise FileNotFoundError(
                        "Google Chrome/Chromium "
                        "não foi encontrado."
                    )

                user_dir = os.path.expanduser(
                    "~/ChromeDevSession"
                )

            else:

                raise RuntimeError(
                    f"Sistema operacional não suportado: "
                    f"{sistema}"
                )

            # ------------------------------------------------
            # PERFIL
            # ------------------------------------------------

            os.makedirs(
                user_dir,
                exist_ok=True
            )

            cmd = [
                chrome_cmd,
                f"--remote-debugging-port={CHROME_DEBUG_PORT}",
                f"--user-data-dir={user_dir}",
            ]

            self.log(
                f"[NAVEGADOR] Executável: {chrome_cmd}"
            )

            self.log(
                f"[NAVEGADOR] Perfil: {user_dir}"
            )

            subprocess.Popen(
                cmd
            )

            # ------------------------------------------------
            # AGUARDA PORTA
            # ------------------------------------------------

            if not esperar_chrome_debugger():

                raise RuntimeError(
                    "O Chrome não disponibilizou "
                    "a porta de depuração."
                )

            self.navegador_aberto = True

            self.log(
                "[NAVEGADOR] Chrome iniciado e pronto."
            )

            self.after(
                0,
                self.atualizar_estado_botao_rodar
            )

            self.after(
                0,
                lambda: self.btn_navegador.configure(
                    state="normal",
                    text="1. Abrir Navegador"
                )
            )

        except Exception as e:

            self.log(
                f"[ERRO] Não foi possível iniciar "
                f"o Chrome: {e}"
            )

            self.after(
                0,
                lambda: self.btn_navegador.configure(
                    state="normal",
                    text="1. Abrir Navegador"
                )
            )

    # ========================================================
    # SELECIONAR EXCEL
    # ========================================================

    def acao_selecionar_excel(self):

        caminho = filedialog.askopenfilename(
            title="Selecione a base",
            filetypes=[
                (
                    "Excel",
                    "*.xlsx *.xls"
                )
            ]
        )

        if caminho:

            self.caminho_excel = caminho

            nome_arquivo = os.path.basename(
                caminho
            )

            self.lbl_status_arquivo.configure(
                text=f"Arquivo: {nome_arquivo}",
                text_color="green"
            )

            self.log(
                f"[ARQUIVO] Planilha selecionada: "
                f"{nome_arquivo}"
            )

            self.excel_selecionado = True

            self.atualizar_estado_botao_rodar()

    # ========================================================
    # RODAR
    # ========================================================

    def acao_rodar_bot(self):

        self.btn_rodar.configure(
            state="disabled",
            text="Processando..."
        )

        self.btn_selecionar_excel.configure(
            state="disabled"
        )

        self.btn_navegador.configure(
            state="disabled"
        )

        threading.Thread(
            target=self._thread_loop_principal,
            daemon=True
        ).start()

    # ========================================================
    # LOOP PRINCIPAL
    # ========================================================

    def _thread_loop_principal(self):

        self.log(
            "\n[EXECUÇÃO] Conectando ao navegador..."
        )

        # ----------------------------------------------------
        # CONECTAR CHROME
        # ----------------------------------------------------

        try:

            opts = Options()

            opts.add_experimental_option(
                "debuggerAddress",
                f"127.0.0.1:{CHROME_DEBUG_PORT}"
            )

            self.driver = webdriver.Chrome(
                options=opts
            )

            if self.driver.window_handles:

                self.driver.switch_to.window(
                    self.driver.window_handles[0]
                )

        except Exception as e:

            self.log(
                f"[ERRO] Não foi possível conectar "
                f"ao Chrome: {e}"
            )

            self.finalizar_processamento()

            return

        # ----------------------------------------------------
        # LER EXCEL
        # ----------------------------------------------------

        self.log(
            "[EXECUÇÃO] Lendo a planilha..."
        )

        try:

            df = pd.read_excel(
                self.caminho_excel,
                dtype=str
            )

            df = df.fillna("")

        except Exception as e:

            self.log(
                f"[ERRO] Falha ao ler Excel: {e}"
            )

            self.finalizar_processamento()

            return

        # ----------------------------------------------------
        # PROCESSAMENTO
        # ----------------------------------------------------

        resultados = []

        total_linhas = len(df)

        self.log(
            f"[EXECUÇÃO] Total de registros: "
            f"{total_linhas}"
        )

        for index, row in df.iterrows():

            numero = index + 1

            self.log(
                f"\n--- Registro {numero} "
                f"de {total_linhas} ---"
            )

            # ------------------------------------------------
            # PROCESSA
            # ------------------------------------------------

            try:

                status, origem = (
                    self.processar_linha_salesforce(
                        self.driver,
                        row
                    )
                )

            except Exception as e:

                status = (
                    "Erro - Necessita de análise humana"
                )

                origem = "Erro ao identificar"

                self.log(
                    f"[ERRO CRÍTICO] Registro "
                    f"{numero}: {e}"
                )

            # ------------------------------------------------
            # GRAVA IMEDIATAMENTE O RESULTADO NA LISTA
            # ------------------------------------------------

            resultados.append(
                {
                    "CNPJ": row.get(
                        "CNPJ",
                        ""
                    ),

                    "Telefone": row.get(
                        "Telefone",
                        ""
                    ),

                    "Origem": origem,

                    "Status": status
                }
            )

            # ------------------------------------------------
            # LOG DA LINHA
            # ------------------------------------------------

            self.log(
                f"[RESULTADO] Origem: {origem} | "
                f"Status: {status}"
            )

        # ----------------------------------------------------
        # SALVAR LOG FINAL
        # ----------------------------------------------------

        try:

            pasta_excel = os.path.dirname(
                os.path.abspath(
                    self.caminho_excel
                )
            )

            caminho_log = os.path.join(
                pasta_excel,
                "Log_Final.xlsx"
            )

            df_resultado = pd.DataFrame(
                resultados
            )

            df_resultado.to_excel(
                caminho_log,
                index=False
            )

            self.log(
                "\n[FIM] Log_Final.xlsx gerado:"
            )

            self.log(
                caminho_log
            )

            self.after(
                0,
                lambda: messagebox.showinfo(
                    "Sucesso",
                    "Processamento finalizado!\n\n"
                    f"Log salvo em:\n{caminho_log}"
                )
            )

        except Exception as e:

            self.log(
                f"[ERRO] Falha ao salvar Log_Final.xlsx: "
                f"{e}"
            )

            self.after(
                0,
                lambda: messagebox.showerror(
                    "Erro",
                    f"Não foi possível salvar "
                    f"Log_Final.xlsx.\n\n{e}"
                )
            )

        self.finalizar_processamento()

    # ========================================================
    # FINALIZAR
    # ========================================================

    def finalizar_processamento(self):

        self.after(
            0,
            lambda: self.btn_selecionar_excel.configure(
                state="normal"
            )
        )

        self.after(
            0,
            lambda: self.btn_navegador.configure(
                state="normal"
            )
        )

        self.after(
            0,
            lambda: self.btn_rodar.configure(
                state="normal",
                text="3. Rodar Bot"
            )
        )

    # ========================================================
    # WAIT
    # ========================================================

    def criar_wait(
        self,
        driver,
        timeout=WAIT_PADRAO
    ):

        return WebDriverWait(
            driver,
            timeout,
            poll_frequency=POLL_FREQUENCY,
            ignored_exceptions=(
                StaleElementReferenceException,
            )
        )

    # ========================================================
    # ESPERAR ELEMENTO
    # ========================================================

    def esperar_elemento(
        self,
        driver,
        xpath,
        timeout=WAIT_PADRAO,
        visivel=True
    ):

        wait = self.criar_wait(
            driver,
            timeout
        )

        if visivel:

            return wait.until(
                EC.visibility_of_element_located(
                    (By.XPATH, xpath)
                )
            )

        return wait.until(
            EC.presence_of_element_located(
                (By.XPATH, xpath)
            )
        )

    # ========================================================
    # CLICAR
    # ========================================================

    def clicar_xpath(
        self,
        driver,
        xpath,
        timeout=WAIT_PADRAO
    ):

        wait = self.criar_wait(
            driver,
            timeout
        )

        elemento = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, xpath)
            )
        )

        try:

            elemento.click()

        except (
            ElementClickInterceptedException,
            WebDriverException
        ):

            driver.execute_script(
                "arguments[0].click();",
                elemento
            )

        return elemento

    # ========================================================
    # JS CLICK
    # ========================================================

    def js_click(
        self,
        driver,
        elemento
    ):

        driver.execute_script(
            "arguments[0].click();",
            elemento
        )

    # ========================================================
    # SCROLL
    # ========================================================

    def scroll_ate(
        self,
        driver,
        elemento
    ):

        driver.execute_script(
            """
            arguments[0].scrollIntoView({
                block: 'center',
                inline: 'nearest'
            });
            """,
            elemento
        )

    # ========================================================
    # PREENCHER CAMPO
    # ========================================================

    def preencher_campo(
        self,
        elemento,
        valor
    ):

        try:

            elemento.click()

        except Exception:

            pass

        try:

            elemento.clear()

        except Exception:

            try:

                elemento.send_keys(
                    Keys.CONTROL + "a"
                )

                elemento.send_keys(
                    Keys.BACKSPACE
                )

            except Exception:

                pass

        elemento.send_keys(
            str(valor)
        )

    # ========================================================
    # IDENTIFICAR ORIGEM
    # ========================================================

    def identificar_origem(
        self,
        driver
    ):
        """
        Identifica a origem do resultado.

        IMPORTANTE:
        Retorna a origem especificamente daquele registro.
        """

        # ----------------------------------------------------
        # PRIMEIRO: aguarda o resultado aparecer
        # ----------------------------------------------------

        xpath_radio = (
            "//input[@type='radio' "
            "and contains(@name, 'options')]"
        )

        try:

            self.criar_wait(
                driver,
                WAIT_PADRAO
            ).until(
                EC.presence_of_element_located(
                    (By.XPATH, xpath_radio)
                )
            )

        except TimeoutException:

            self.log(
                "[AVISO] Resultado da pesquisa "
                "não apareceu."
            )

            return "Erro ao identificar"

        # ----------------------------------------------------
        # MÉTODO PRINCIPAL - IGUAL AO SEU ORIGINAL
        # ----------------------------------------------------

        xpath_origem = (
            "//lightning-base-formatted-text"
            "[contains(., 'CPE')]"
        )

        try:

            elemento_origem = self.criar_wait(
                driver,
                WAIT_PADRAO
            ).until(
                EC.presence_of_element_located(
                    (By.XPATH, xpath_origem)
                )
            )

            texto = elemento_origem.text.strip()

            if texto:

                # Se o texto encontrado for CPE,
                # normalizamos para exatamente CPE.
                if texto.upper() == "CPE":

                    origem = "CPE"

                else:

                    origem = texto

                self.log(
                    f" [INFO] Origem identificada: "
                    f"{origem}"
                )

                return origem

        except TimeoutException:

            pass

        except Exception as e:

            self.log(
                f" [AVISO] Erro ao capturar "
                f"origem principal: {e}"
            )

        # ----------------------------------------------------
        # FALLBACK
        # ----------------------------------------------------
        # Se a estrutura do Salesforce mudou,
        # procuramos elementos visíveis contendo CPE.
        # ----------------------------------------------------

        try:

            elementos = driver.find_elements(
                By.XPATH,
                "//*[contains("
                "translate(normalize-space(.),"
                "'cpe','CPE'),"
                "'CPE')]"
            )

            for elemento in elementos:

                try:

                    if not elemento.is_displayed():
                        continue

                    texto = elemento.text.strip()

                    if not texto:
                        continue

                    if texto.upper() == "CPE":

                        self.log(
                            " [INFO] Origem identificada "
                            "via fallback: CPE"
                        )

                        return "CPE"

                except (
                    StaleElementReferenceException
                ):

                    continue

        except Exception:
            pass

        # ----------------------------------------------------
        # NÃO IDENTIFICADO
        # ----------------------------------------------------

        self.log(
            " [AVISO] Não foi possível identificar "
            "a origem."
        )

        return "Não Identificado"

    # ========================================================
    # VERIFICAR RFB
    # ========================================================

    def verificar_rfb(
        self,
        driver
    ):

        xpath_situacao = (
            "//span[text()='Situação Cadastral RFB']"
            "/ancestor::div[contains("
            "@class, 'slds-form-element')]"
            "//force-record-output-picklist"
        )

        try:

            elemento = self.criar_wait(
                driver,
                2
            ).until(
                EC.visibility_of_element_located(
                    (By.XPATH, xpath_situacao)
                )
            )

            situacao = elemento.text.strip()

            self.log(
                f" [INFO] Situação Cadastral RFB: "
                f"{situacao}"
            )

            return situacao

        except TimeoutException:

            self.log(
                " [INFO] Campo Situação Cadastral "
                "RFB não encontrado."
            )

            return ""

        except Exception as e:

            self.log(
                f" [AVISO] Erro ao verificar RFB: "
                f"{e}"
            )

            return ""

    # ========================================================
    # FECHAR MODAL ENDEREÇO
    # ========================================================

    def fechar_modal_endereco(
        self,
        driver
    ):

        xpath_fechar = (
            "//lc-verify-address"
            "//button[@title='Close']"
            " | "
            "//section[@role='dialog']"
            "//button[@title='Close']"
        )

        try:

            # Aqui usamos wait curto.
            # O modal é opcional.
            elemento = self.criar_wait(
                driver,
                WAIT_CURTO
            ).until(
                EC.visibility_of_element_located(
                    (By.XPATH, xpath_fechar)
                )
            )

            self.js_click(
                driver,
                elemento
            )

            self.log(
                " [OK] Modal de Atualizar "
                "Endereço fechado."
            )

            return True

        except TimeoutException:

            self.log(
                " [INFO] Modal de Atualizar "
                "Endereço não apareceu."
            )

            return False

        except Exception as e:

            self.log(
                f" [AVISO] Erro ao fechar modal: "
                f"{e}"
            )

            return False

    # ========================================================
    # SELECIONAR SIM
    # ========================================================

    def selecionar_sim(
        self,
        driver,
        wait,
        xpath_botao,
        descricao
    ):

        btn = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, xpath_botao)
            )
        )

        valor_atual = btn.get_attribute(
            "data-value"
        )

        if (
            valor_atual
            and valor_atual.strip().lower()
            == "sim"
        ):

            self.log(
                f" [INFO] {descricao} "
                f"já está como 'Sim'."
            )

            return

        self.log(
            f" [AÇÃO] Alterando {descricao} "
            f"para 'Sim'..."
        )

        self.js_click(
            driver,
            btn
        )

        xpath_opcoes_sim = (
            "//lightning-base-combobox-item"
            "[@data-value='Sim']"
        )

        # ----------------------------------------------------
        # Em vez de esperar presença de TODOS,
        # procuramos um que esteja visível.
        # ----------------------------------------------------

        def encontrar_opcao_visivel(d):

            elementos = d.find_elements(
                By.XPATH,
                xpath_opcoes_sim
            )

            for elemento in elementos:

                try:

                    if elemento.is_displayed():

                        return elemento

                except (
                    StaleElementReferenceException
                ):

                    continue

            return False

        opcao_sim = self.criar_wait(
            driver,
            WAIT_CURTO
        ).until(
            encontrar_opcao_visivel
        )

        self.js_click(
            driver,
            opcao_sim
        )

        self.log(
            f" [OK] {descricao} "
            f"marcada como 'Sim'."
        )

    # ========================================================
    # PROCESSAR LINHA
    # ========================================================

    def processar_linha_salesforce(
        self,
        driver,
        row_data
    ):
        """
        Retorna:

            (status, origem)

        Exemplo:

            ("Cadastrado com sucesso", "CPE")

        """

        wait = self.criar_wait(
            driver,
            WAIT_PADRAO
        )

        # ====================================================
        # TRATAR CNPJ
        # ====================================================

        cnpj_bruto = str(
            row_data.get(
                "CNPJ",
                ""
            )
        ).strip()

        if (
            cnpj_bruto.endswith(".0")
            and cnpj_bruto.count(".") == 1
        ):

            cnpj_bruto = cnpj_bruto[:-2]

        cnpj_numeros = "".join(
            filter(
                str.isdigit,
                cnpj_bruto
            )
        )

        cnpj = (
            cnpj_numeros.zfill(14)
            if cnpj_numeros
            else ""
        )

        # ====================================================
        # TRATAR TELEFONE
        # ====================================================

        telefone_bruto = str(
            row_data.get(
                "Telefone",
                ""
            )
        ).strip()

        if telefone_bruto.endswith(".0"):

            telefone_bruto = (
                telefone_bruto[:-2]
            )

        telefone_limpo = "".join(
            filter(
                str.isdigit,
                telefone_bruto
            )
        )

        # ----------------------------------------------------
        # 8 dígitos -> coloca 9
        # ----------------------------------------------------

        if len(telefone_limpo) == 8:

            telefone_limpo = (
                "9" + telefone_limpo
            )

        # ----------------------------------------------------
        # DDD + 8 dígitos -> coloca 9
        # ----------------------------------------------------

        elif len(telefone_limpo) == 10:

            telefone_limpo = (
                telefone_limpo[:2]
                + "9"
                + telefone_limpo[2:]
            )

        self.log(
            f" -> CNPJ Tratado: {cnpj} | "
            f"Tel Tratado: {telefone_limpo}"
        )

        # ====================================================
        # ORIGEM PADRÃO
        # ====================================================

        origem = "Não Identificado"

        try:

            # =================================================
            # ABRIR PÁGINA
            # =================================================

            driver.get(
                URL_NOVO_CADASTRO
            )

            # =================================================
            # SELECT CNPJ
            # =================================================

            xpath_select = (
                "//select[@name='search']"
            )

            select_elemento = (
                self.esperar_elemento(
                    driver,
                    xpath_select
                )
            )

            dropdown = Select(
                select_elemento
            )

            dropdown.select_by_value(
                "CNPJ"
            )

            # =================================================
            # INPUT CNPJ
            # =================================================

            xpath_input_cnpj = (
                "//label[text()='CNPJ']"
                "/following::div/input"
                "[@class='slds-input']"
            )

            input_cnpj = (
                self.esperar_elemento(
                    driver,
                    xpath_input_cnpj
                )
            )

            self.scroll_ate(
                driver,
                input_cnpj
            )

            self.preencher_campo(
                input_cnpj,
                cnpj
            )

            # =================================================
            # BUSCAR
            # =================================================

            xpath_btn_buscar = (
                "//button[contains("
                "@class, 'slds-button_brand')]"
                "[normalize-space(text())='Buscar']"
            )

            self.log(
                " [AÇÃO] Clicando em Buscar..."
            )

            self.clicar_xpath(
                driver,
                xpath_btn_buscar
            )

            # =================================================
            # AGUARDAR RESULTADO
            # =================================================

            xpath_radio = (
                "//input[@type='radio' "
                "and contains(@name, 'options')]"
            )

            radio_resultado = (
                self.esperar_elemento(
                    driver,
                    xpath_radio
                )
            )

            # =================================================
            # IDENTIFICAR ORIGEM
            # =================================================

            origem = self.identificar_origem(
                driver
            )

            # -------------------------------------------------
            # GARANTIA:
            # se origem for CPE, fica exatamente "CPE".
            # -------------------------------------------------

            if origem.strip().upper() == "CPE":

                origem = "CPE"

            # -------------------------------------------------
            # LOG EXPLÍCITO
            # -------------------------------------------------

            self.log(
                f" [ORIGEM DO REGISTRO] {origem}"
            )

            # =================================================
            # SELECIONAR RESULTADO
            # =================================================

            self.js_click(
                driver,
                radio_resultado
            )

            # =================================================
            # CONTINUAR
            # =================================================

            xpath_continuar = (
                "//button[contains("
                "@class, 'slds-button_brand') "
                "and contains(., 'Continuar')]"
            )

            self.log(
                " [AÇÃO] Clicando em Continuar..."
            )

            self.clicar_xpath(
                driver,
                xpath_continuar
            )

            # =================================================
            # AGUARDAR FORMULÁRIO
            # =================================================

            xpath_indicador_formulario = (
                "//button[@name='SaveEdit']"
                " | "
                "//button[@title='Salvar']"
                " | "
                "//input[@name='Phone']"
                " | "
                "//input[@name="
                "'PessoasOcupadas__c']"
            )

            try:

                self.esperar_elemento(
                    driver,
                    xpath_indicador_formulario,
                    timeout=WAIT_PADRAO
                )

            except TimeoutException:

                self.log(
                    " [AVISO] Formulário não "
                    "detectado pelo indicador."
                )

            # =================================================
            # MODAL ENDEREÇO
            # =================================================

            if origem != "CPE":

                self.log(
                    " [AÇÃO] Origem diferente de CPE. "
                    "Verificando modal..."
                )

                self.fechar_modal_endereco(
                    driver
                )

            else:

                self.log(
                    " [INFO] Origem CPE. "
                    "Pulando modal de endereço."
                )

            # =================================================
            # SITUAÇÃO RFB
            # =================================================

            self.log(
                " [AÇÃO] Verificando "
                "Situação Cadastral RFB..."
            )

            situacao = self.verificar_rfb(
                driver
            )

            if situacao in (
                "Suspensa",
                "Baixada"
            ):

                self.log(
                    f" [PULO] CNPJ ignorado. "
                    f"Situação Cadastral RFB: "
                    f"{situacao}"
                )

                xpath_cancelar = (
                    "//button[@name='CancelEdit' "
                    "or @title='Cancelar']"
                )

                try:

                    self.clicar_xpath(
                        driver,
                        xpath_cancelar,
                        timeout=3
                    )

                except Exception:

                    pass

                return (
                    f"Pulado - RFB {situacao}",
                    origem
                )

            # =================================================
            # ROTINA CPE
            # =================================================

            if origem == "CPE":

                self.log(
                    " [INFO] Executando rotina "
                    "específica para CPE."
                )

                # ---------------------------------------------
                # PESSOAS OCUPADAS
                # ---------------------------------------------

                xpath_pessoas = (
                    "//input[@name="
                    "'PessoasOcupadas__c']"
                )

                try:

                    input_pessoas = (
                        self.esperar_elemento(
                            driver,
                            xpath_pessoas
                        )
                    )

                    valor_atual = (
                        input_pessoas.get_attribute(
                            "value"
                        )
                    )

                    if not valor_atual:

                        self.log(
                            " [AÇÃO] PessoasOcupadas "
                            "vazio. Preenchendo 1."
                        )

                        self.preencher_campo(
                            input_pessoas,
                            "1"
                        )

                    else:

                        self.log(
                            f" [INFO] PessoasOcupadas "
                            f"já preenchido: "
                            f"{valor_atual}"
                        )

                except TimeoutException:

                    self.log(
                        " [AVISO] Campo "
                        "PessoasOcupadas não encontrado."
                    )

                # ---------------------------------------------
                # TELEFONE
                # ---------------------------------------------

                if telefone_limpo:

                    xpath_telefone = (
                        "//input[@name='Phone']"
                    )

                    try:

                        input_telefone = (
                            self.esperar_elemento(
                                driver,
                                xpath_telefone
                            )
                        )

                        self.preencher_campo(
                            input_telefone,
                            telefone_limpo
                        )

                        self.log(
                            " [AÇÃO] Telefone preenchido."
                        )

                    except TimeoutException:

                        self.log(
                            " [AVISO] Campo Telefone "
                            "não encontrado."
                        )

                # ---------------------------------------------
                # WHATSAPP
                # ---------------------------------------------

                self.selecionar_sim(
                    driver,
                    wait,
                    "//button[@aria-label="
                    "'Autoriza Receber Whatsapp?']",
                    "WhatsApp"
                )

                # ---------------------------------------------
                # LIGAÇÃO
                # ---------------------------------------------

                self.selecionar_sim(
                    driver,
                    wait,
                    "//button[@aria-label="
                    "'Autoriza Receber Ligação']",
                    "Ligação"
                )

                # ---------------------------------------------
                # MENSAGEM
                # ---------------------------------------------

                self.selecionar_sim(
                    driver,
                    wait,
                    "//button[@aria-label="
                    "'Autoriza Receber Mensagem']",
                    "Mensagem"
                )

            # =================================================
            # ROTINA ORIGEM DIVERSA
            # =================================================

            else:

                self.log(
                    f" [INFO] Executando rotina "
                    f"para origem: {origem}"
                )

                # ---------------------------------------------
                # ATUALIZAR RFB
                # ---------------------------------------------

                self.log(
                    " [AÇÃO] Clicando em "
                    "'Atualizar dados com RFB'..."
                )

                xpath_btn_rfb = (
                    "//button[@name="
                    "'Account.updateAccountPJCPE']"
                    " | "
                    "//button[normalize-space(text())="
                    "'Atualizar dados com RFB']"
                )

                self.clicar_xpath(
                    driver,
                    xpath_btn_rfb
                )

                # ---------------------------------------------
                # CONFIRMAR
                # ---------------------------------------------

                self.log(
                    " [AÇÃO] Procurando "
                    "botão Confirmar..."
                )

                xpath_confirmar = (
                    "//button[contains("
                    "@class, 'slds-button_brand')]"
                    "[normalize-space(text())="
                    "'Confirmar']"
                )

                try:

                    def achar_confirmar(d):

                        botoes = d.find_elements(
                            By.XPATH,
                            xpath_confirmar
                        )

                        for botao in botoes:

                            try:

                                if (
                                    botao.is_displayed()
                                    and botao.is_enabled()
                                ):

                                    return botao

                            except (
                                StaleElementReferenceException
                            ):

                                continue

                        return False

                    btn_confirmar = (
                        self.criar_wait(
                            driver,
                            WAIT_PADRAO
                        ).until(
                            achar_confirmar
                        )
                    )

                    self.js_click(
                        driver,
                        btn_confirmar
                    )

                    self.log(
                        " [OK] Confirmar clicado."
                    )

                except TimeoutException:

                    self.log(
                        " [AVISO] Confirmar não "
                        "encontrado pelo primeiro método."
                    )

                    # -----------------------------------------
                    # FALLBACK
                    # -----------------------------------------

                    xpath_confirmar_fallback = (
                        "//button[normalize-space("
                        "text())='Confirmar']"
                    )

                    botoes = driver.find_elements(
                        By.XPATH,
                        xpath_confirmar_fallback
                    )

                    clicado = False

                    for botao in botoes:

                        try:

                            if (
                                botao.is_displayed()
                                and botao.is_enabled()
                            ):

                                self.js_click(
                                    driver,
                                    botao
                                )

                                clicado = True

                                self.log(
                                    " [OK] Confirmar "
                                    "clicado via fallback."
                                )

                                break

                        except Exception:

                            continue

                    if not clicado:

                        self.log(
                            " [ERRO] Não foi possível "
                            "clicar em Confirmar."
                        )

                # ---------------------------------------------
                # AGUARDAR SALVAR
                # ---------------------------------------------

                try:

                    self.esperar_elemento(
                        driver,
                        "//button[@name='SaveEdit']"
                        " | "
                        "//button[@title='Salvar']",
                        timeout=WAIT_PADRAO
                    )

                except TimeoutException:

                    pass

            # =================================================
            # SALVAR
            # =================================================

            self.log(
                " [AÇÃO] Salvando o cadastro..."
            )

            xpath_salvar = (
                "//button[@name='SaveEdit']"
                " | "
                "//button[@title='Salvar']"
                " | "
                "//button[normalize-space(text())="
                "'Salvar']"
            )

            btn_salvar = self.clicar_xpath(
                driver,
                xpath_salvar
            )

            # =================================================
            # AGUARDAR RESULTADO DO SAVE
            # =================================================

            xpath_toast_sucesso = (
                "//div[contains("
                "@class,'toastMessage')]"
            )

            try:

                toast = self.criar_wait(
                    driver,
                    8
                ).until(
                    EC.visibility_of_element_located(
                        (By.XPATH, xpath_toast_sucesso)
                    )
                )

                texto_toast = (
                    toast.text.strip()
                )

                self.log(
                    f" [OK] Cadastro salvo. "
                    f"Mensagem: {texto_toast}"
                )

                return (
                    "Cadastrado com sucesso",
                    origem
                )

            except TimeoutException:

                # ------------------------------------------------
                # FALLBACK:
                # verifica se o botão salvar desapareceu.
                # ------------------------------------------------

                try:

                    WebDriverWait(
                        driver,
                        3,
                        poll_frequency=POLL_FREQUENCY
                    ).until(
                        EC.staleness_of(
                            btn_salvar
                        )
                    )

                    self.log(
                        " [OK] Cadastro salvo "
                        "(botão desapareceu)."
                    )

                    return (
                        "Cadastrado com sucesso",
                        origem
                    )

                except Exception:

                    self.log(
                        " [AVISO] Não foi possível "
                        "confirmar o salvamento."
                    )

                    return (
                        "Salvo - confirmação não localizada",
                        origem
                    )

        # ====================================================
        # ERRO
        # ====================================================

        except Exception as e:

            self.log(
                f" [ERRO] Falha na execução: {e}"
            )

            return (
                "Erro - Necessita de análise humana",
                origem
            )


# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":

    app = BotExcelSalesforceApp()

    app.mainloop()
