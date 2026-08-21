import os
import sys
import time
import subprocess
import platform
import shutil
import socket
import json
import urllib.request
import urllib.error
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
    # IDENTIFICAR ORIGEM  (le o texto real da linha do resultado)
    # ========================================================

    def identificar_origem(self, driver):
        """
        Le a origem do registro na tabela de resultados.

        ANTES: o XPath era //lightning-base-formatted-text[contains(.,'CPE')]
        ou seja, so encontrava origem se o texto contivesse "CPE".
        Origens como "Salesforce" caiam no fallback (que tambem so
        procurava CPE) e viravam "Nao Identificado" no Log_Final.

        AGORA: lemos a linha do resultado e devolvemos o texto real,
        qualquer que seja ele.
        """

        xpath_linha = (
            "//tr[.//input[@type='radio' and contains(@name,'options')]]"
        )

        # ----- 1) aguarda a linha do resultado -----

        try:
            linha = self.criar_wait(driver, WAIT_PADRAO).until(
                EC.presence_of_element_located((By.XPATH, xpath_linha))
            )
        except TimeoutException:
            self.log(" [AVISO] Resultado da pesquisa nao apareceu.")
            return "Erro ao identificar"

        # ----- 2) le as celulas da MESMA linha -----

        textos = []

        try:
            celulas = linha.find_elements(
                By.XPATH,
                ".//lightning-base-formatted-text"
                " | .//lightning-formatted-text"
                " | .//td//span"
            )

            for celula in celulas:
                try:
                    texto = " ".join(celula.text.split()).strip()
                    if texto and texto not in textos:
                        textos.append(texto)
                except StaleElementReferenceException:
                    continue

        except Exception as e:
            self.log(f" [AVISO] Falha ao ler celulas da linha: {e}")

        # ----- 3) procura a celula que parece ser a origem -----
        # Nao ha hardcode de valores: aceitamos qualquer texto que
        # contenha CPE ou SALESFORCE, em qualquer combinacao.

        for texto in textos:

            alvo = texto.upper()

            if "CPE" in alvo or "SALESFORCE" in alvo:

                origem = "CPE" if alvo == "CPE" else texto

                self.log(f" [INFO] Origem identificada: {origem}")

                return origem

        # ----- 4) fallback: varre a pagina inteira -----

        try:
            elementos = driver.find_elements(
                By.XPATH,
                "//lightning-base-formatted-text"
                " | //lightning-formatted-text"
            )

            for elemento in elementos:
                try:
                    if not elemento.is_displayed():
                        continue

                    texto = " ".join(elemento.text.split()).strip()
                    alvo = texto.upper()

                    if alvo in ("CPE", "SALESFORCE", "CPE/SALESFORCE"):
                        origem = "CPE" if alvo == "CPE" else texto
                        self.log(
                            f" [INFO] Origem identificada via fallback: {origem}"
                        )
                        return origem

                except StaleElementReferenceException:
                    continue

        except Exception:
            pass

        # ----- 5) ultimo recurso: mostra o que foi lido -----

        if textos:
            amostra = " | ".join(textos[:6])
            self.log(
                f" [AVISO] Origem nao reconhecida. "
                f"Celulas lidas: {amostra}"
            )
        else:
            self.log(" [AVISO] Nenhum texto lido na linha do resultado.")

        return "Nao Identificado"

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
    # ========================================================
    # ATUALIZAR DADOS COM RFB (rotina nao-CPE)
    # ========================================================

    def atualizar_dados_rfb(self, driver):
        """Clica em 'Atualizar dados com RFB' e confirma. Nunca derruba o registro."""

        xpath_btn_rfb = (
            "//button[@name='Account.updateAccountPJCPE']"
            " | //button[normalize-space(text())='Atualizar dados com RFB']"
            " | //button[contains(normalize-space(.),'Atualizar dados com RFB')]"
        )

        self.log(" [ACAO] Clicando em 'Atualizar dados com RFB'...")

        try:
            self.clicar_xpath(driver, xpath_btn_rfb, timeout=WAIT_PADRAO)
        except TimeoutException:
            self.log(" [AVISO] Botao RFB nao encontrado. Seguindo para a aba Contato.")
            return False
        except Exception as e:
            self.log(f" [AVISO] Falha ao clicar em RFB: {e}")
            return False

        self.log(" [ACAO] Procurando botao Confirmar...")

        xpath_confirmar = (
            "//button[normalize-space(text())='Confirmar']"
            " | //button[contains(normalize-space(.),'Confirmar')]"
        )

        def achar_confirmar(d):
            for botao in d.find_elements(By.XPATH, xpath_confirmar):
                try:
                    if botao.is_displayed() and botao.is_enabled():
                        return botao
                except StaleElementReferenceException:
                    continue
            return False

        try:
            btn = self.criar_wait(driver, WAIT_PADRAO).until(achar_confirmar)
            self.js_click(driver, btn)
            self.log(" [OK] Confirmar clicado.")
        except TimeoutException:
            self.log(" [AVISO] Confirmar nao apareceu.")

        try:
            self.esperar_elemento(
                driver,
                "//button[@name='SaveEdit'] | //button[@title='Salvar']",
                timeout=WAIT_PADRAO
            )
            self.log(" [OK] Formulario recarregado apos RFB.")
        except TimeoutException:
            self.log(" [AVISO] Formulario nao confirmado apos RFB.")

        return True

    # ========================================================
    # ABA CONTATO (clique blindado)
    # ========================================================

    def clicar_aba_contato(self, driver):
        """
        Abre a aba Contato.
        Correcao do travamento: so clica em elemento VISIVEL, tenta varios
        seletores, usa 3 estrategias de clique e CONFIRMA se a aba trocou.
        """

        seletores = [
            "//a[@data-label='Contato']",
            "//a[@data-label='Contatos']",
            "//*[@role='tab'][normalize-space(.)='Contato']",
            "//*[@role='tab'][contains(normalize-space(.),'Contato')]",
            "//li[contains(@class,'slds-tabs_default__item')]//a[contains(normalize-space(.),'Contato')]",
            "//lightning-tab-bar//a[contains(normalize-space(.),'Contato')]",
            "//span[normalize-space(text())='Contato']/ancestor::a[1]",
            "//span[normalize-space(text())='Contato']/ancestor::button[1]",
            "//a[contains(normalize-space(.),'Contato')]",
        ]

        self.log(" [ACAO] Abrindo a aba Contato...")

        for xpath in seletores:

            try:
                elementos = driver.find_elements(By.XPATH, xpath)
            except Exception:
                continue

            for elemento in elementos:

                try:
                    if not elemento.is_displayed():
                        continue
                    if elemento.get_attribute("aria-selected") == "true":
                        self.log(" [OK] Aba Contato ja estava ativa.")
                        return True
                except StaleElementReferenceException:
                    continue

                self.scroll_ate(driver, elemento)

                try:
                    elemento.click()
                except Exception:
                    pass
                if self.aba_contato_abriu(driver, elemento):
                    self.log(" [OK] Aba Contato aberta (clique nativo).")
                    return True

                try:
                    self.js_click(driver, elemento)
                except Exception:
                    pass
                if self.aba_contato_abriu(driver, elemento):
                    self.log(" [OK] Aba Contato aberta (JS click).")
                    return True

                try:
                    driver.execute_script(
                        """
                        const el = arguments[0];
                        ['pointerdown','mousedown','pointerup','mouseup','click']
                          .forEach(function (t) {
                            el.dispatchEvent(new MouseEvent(t, {
                                bubbles: true, cancelable: true, view: window
                            }));
                        });
                        """,
                        elemento
                    )
                except Exception:
                    pass
                if self.aba_contato_abriu(driver, elemento):
                    self.log(" [OK] Aba Contato aberta (eventos de mouse).")
                    return True

        self.log(" [ERRO] Nao foi possivel abrir a aba Contato.")
        return False

    def aba_contato_abriu(self, driver, elemento):
        """Confirma a troca: aria-selected='true' OU campo de telefone visivel."""

        try:
            if elemento.get_attribute("aria-selected") == "true":
                return True
        except StaleElementReferenceException:
            return True
        except Exception:
            pass

        try:
            self.criar_wait(driver, WAIT_CURTO).until(
                EC.visibility_of_element_located((By.XPATH, self.xpath_telefone()))
            )
            return True
        except TimeoutException:
            return False

    # ========================================================
    # TELEFONE
    # ========================================================

    # ========================================================
    # BOTAO EDITAR (aba Contato / rotina nao-CPE)
    # ========================================================
    # CONTATO PRINCIPAL x ALTERNATIVO 1  (rotina nao-CPE)
    # ========================================================

    @staticmethod
    def so_digitos(valor):
        return "".join(filter(str.isdigit, str(valor or "")))

    def mesmo_telefone(self, a, b):
        """
        Compara ignorando mascara, DDD e o 9o digito.
        (73)99147-2752  ==  73991472752  ==  991472752
        """

        a = self.so_digitos(a)
        b = self.so_digitos(b)

        if not a or not b:
            return False

        if a == b:
            return True

        # compara os 8 digitos finais (nucleo do numero)
        return a[-8:] == b[-8:]

    def xpath_input_label(self, rotulo):
        """Input de texto localizado pelo rotulo visivel na tela."""

        return (
            f"//lightning-input[.//label[normalize-space(text())='{rotulo}']]//input"
            f" | //label[normalize-space(text())='{rotulo}']/following::input[1]"
            f" | //label[starts-with(normalize-space(.),'{rotulo}')]/following::input[1]"
        )

    def ler_campo_label(self, driver, rotulo, timeout=None):
        """Devolve (elemento, valor_atual) ou (None, '')."""

        try:
            elemento = self.esperar_elemento(
                driver,
                self.xpath_input_label(rotulo),
                timeout=timeout or WAIT_PADRAO
            )
            valor = (elemento.get_attribute("value") or "").strip()
            return elemento, valor

        except TimeoutException:
            self.log(f" [AVISO] Campo '{rotulo}' nao encontrado.")
            return None, ""

        except Exception as e:
            self.log(f" [AVISO] Falha ao ler '{rotulo}': {e}")
            return None, ""

    def tratar_contatos(self, driver, wait, telefone):
        """
        Regra do formulario nao-CPE (Contato Principal / Alternativo 1):

        1. Principal VAZIO          -> grava no Principal   + 3 consentimentos do Principal
        2. Principal IGUAL ao nosso -> nao altera o numero  + 3 consentimentos do Principal
        3. Principal DIFERENTE      -> grava no Alternativo 1 + 3 consentimentos do Alternativo 1

        Devolve o texto que vai para a coluna Status/observacao do log.
        """

        consent_principal = [
            ("//button[@aria-label='Autoriza Receber Whatsapp?']", "WhatsApp"),
            ("//button[@aria-label='Autoriza Receber Ligação']", "Ligacao"),
            ("//button[@aria-label='Autoriza Receber Mensagem']", "Mensagem"),
        ]

        consent_alternativo = [
            ("//button[@aria-label='Autoriza Receber Whatsapp 1?']", "WhatsApp 1"),
            ("//button[@aria-label='Autoriza Receber Ligação 1']", "Ligacao 1"),
            ("//button[@aria-label='Autoriza Receber Mensagem 1']", "Mensagem 1"),
        ]

        campo_principal, valor_principal = self.ler_campo_label(
            driver, "Telefone Principal"
        )

        self.log(
            f" [INFO] Telefone Principal na tela: "
            f"{valor_principal or '(vazio)'} | Planilha: {telefone or '(vazio)'}"
        )

        # ---------- sem telefone na planilha ----------

        if not telefone:
            self.log(" [INFO] Sem telefone na planilha. Apenas consentimentos do Principal.")
            self.aplicar_consentimentos(driver, wait, consent_principal)
            return "Sem telefone - consentimentos aplicados"

        # ---------- CASO 1: principal vazio ----------

        if not valor_principal:

            self.log(" [ACAO] Telefone Principal vazio. Preenchendo no Principal.")

            if campo_principal:
                self.scroll_ate(driver, campo_principal)
                self.preencher_campo(campo_principal, telefone)
                self.log(f" [OK] Telefone Principal preenchido: {telefone}")

            self.aplicar_consentimentos(driver, wait, consent_principal)
            return "Telefone gravado no Principal"

        # ---------- CASO 2: principal igual ----------

        if self.mesmo_telefone(valor_principal, telefone):

            self.log(
                " [INFO] Telefone Principal ja e o mesmo da base. "
                "Nao sera alterado."
            )

            self.aplicar_consentimentos(driver, wait, consent_principal)
            return "Telefone ja cadastrado - consentimentos aplicados"

        # ---------- CASO 3: principal diferente -> alternativo 1 ----------

        self.log(
            f" [ACAO] Telefone Principal ({valor_principal}) e diferente da base "
            f"({telefone}). Gravando como Alternativo 1."
        )

        campo_alt, valor_alt = self.ler_campo_label(
            driver, "Telefone Alternativo 1"
        )

        if campo_alt is None:
            self.log(" [ERRO] Campo 'Telefone Alternativo 1' nao encontrado.")
            self.aplicar_consentimentos(driver, wait, consent_principal)
            return "Erro - Alternativo 1 nao encontrado"

        # Se o alternativo ja tiver o nosso numero, nao reescreve.
        if valor_alt and self.mesmo_telefone(valor_alt, telefone):
            self.log(" [INFO] Alternativo 1 ja possui este numero. Nao sera alterado.")
        else:
            if valor_alt:
                self.log(
                    f" [AVISO] Alternativo 1 ja tinha '{valor_alt}'. "
                    f"Sera substituido pelo numero da base."
                )
            self.scroll_ate(driver, campo_alt)
            self.preencher_campo(campo_alt, telefone)
            self.log(f" [OK] Telefone Alternativo 1 preenchido: {telefone}")

        self.aplicar_consentimentos(driver, wait, consent_alternativo)
        return "Telefone gravado no Alternativo 1"

    def aplicar_consentimentos(self, driver, wait, campos):
        """Marca 'Sim' na lista de picklists recebida. Falha isolada nao derruba o resto."""

        for xpath, descricao in campos:
            try:
                self.selecionar_sim(driver, wait, xpath, descricao)
            except TimeoutException:
                self.log(f" [AVISO] Campo '{descricao}' nao encontrado.")
            except Exception as e:
                self.log(f" [AVISO] Falha em '{descricao}': {e}")

    # ========================================================
    # ENDERECO VIA CEP  (rotina CPE)
    # ========================================================

    UFS = {
        "AC": "Acre", "AL": "Alagoas", "AP": "Amapá", "AM": "Amazonas",
        "BA": "Bahia", "CE": "Ceará", "DF": "Distrito Federal",
        "ES": "Espírito Santo", "GO": "Goiás", "MA": "Maranhão",
        "MT": "Mato Grosso", "MS": "Mato Grosso do Sul", "MG": "Minas Gerais",
        "PA": "Pará", "PB": "Paraíba", "PR": "Paraná", "PE": "Pernambuco",
        "PI": "Piauí", "RJ": "Rio de Janeiro", "RN": "Rio Grande do Norte",
        "RS": "Rio Grande do Sul", "RO": "Rondônia", "RR": "Roraima",
        "SC": "Santa Catarina", "SP": "São Paulo", "SE": "Sergipe",
        "TO": "Tocantins",
    }

    def consultar_cep(self, cep):
        """
        Consulta ViaCEP e, em caso de falha, BrasilAPI.
        Usa cache em memoria: o mesmo CEP nao e consultado duas vezes.
        Retorna dict com logradouro/bairro/cidade/uf ou None.
        """

        cep = self.so_digitos(cep)

        if len(cep) != 8:
            self.log(f" [AVISO] CEP invalido para consulta: '{cep}'")
            return None

        if not hasattr(self, "_cache_cep"):
            self._cache_cep = {}

        if cep in self._cache_cep:
            return self._cache_cep[cep]

        fontes = [
            ("ViaCEP", f"https://viacep.com.br/ws/{cep}/json/"),
            ("BrasilAPI", f"https://brasilapi.com.br/api/cep/v1/{cep}"),
        ]

        for nome, url in fontes:

            try:
                requisicao = urllib.request.Request(
                    url, headers={"User-Agent": "BotSalesforceFOCO/1.0"}
                )

                with urllib.request.urlopen(requisicao, timeout=8) as resposta:
                    dados = json.loads(resposta.read().decode("utf-8"))

                if dados.get("erro"):
                    self.log(f" [AVISO] {nome}: CEP {cep} nao encontrado.")
                    continue

                endereco = {
                    "logradouro": (
                        dados.get("logradouro") or dados.get("street") or ""
                    ).strip(),
                    "bairro": (
                        dados.get("bairro") or dados.get("neighborhood") or ""
                    ).strip(),
                    "cidade": (
                        dados.get("localidade") or dados.get("city") or ""
                    ).strip(),
                    "uf": (
                        dados.get("uf") or dados.get("state") or ""
                    ).strip().upper(),
                }

                if not endereco["cidade"]:
                    continue

                endereco["estado"] = self.UFS.get(
                    endereco["uf"], endereco["uf"]
                )

                self.log(
                    f" [OK] CEP {cep} ({nome}): "
                    f"{endereco['cidade']}/{endereco['uf']}"
                )

                self._cache_cep[cep] = endereco
                return endereco

            except Exception as e:
                self.log(f" [AVISO] Falha na consulta {nome}: {type(e).__name__}")
                continue

        self.log(f" [ERRO] Nao foi possivel consultar o CEP {cep}.")
        self._cache_cep[cep] = None
        return None

    def selecionar_picklist_valor(self, driver, rotulo, valor):
        """
        Seleciona um valor em combobox do Lightning (ex.: Estado, Pais).
        Compara ignorando maiusculas/acentos simples.
        """

        if "autoriza" in rotulo.strip().lower():
            self.log(f" [BLOQUEIO] Campo protegido, nao sera alterado: {rotulo}")
            return False

        xpath_container = (
            f"//label[normalize-space(text())='{rotulo}']"
            f"/ancestor::*[contains(@class,'slds-form-element')][1]"
        )

        botao = None

        try:
            for container in driver.find_elements(By.XPATH, xpath_container):
                if not container.is_displayed():
                    continue
                if "autoriza" in container.text.strip().lower():
                    continue
                for alvo in container.find_elements(
                    By.XPATH,
                    ".//button[@role='combobox']"
                    " | .//input[@role='combobox']"
                    " | .//button[contains(@class,'slds-combobox__input')]"
                    " | .//div[contains(@class,'slds-combobox')]//button"
                ):
                    if alvo.is_displayed():
                        botao = alvo
                        break
                if botao is not None:
                    break
        except Exception:
            botao = None

        if botao is None:
            self.log(f" [AVISO] Caixa seletora nao encontrada: {rotulo}")
            return False

        atual = (
            botao.get_attribute("data-value")
            or botao.get_attribute("value")
            or botao.text
            or ""
        ).strip()

        if atual and atual.lower() not in ("--nenhum--", "--none--", "none", "selecione"):
            self.log(f" [INFO] {rotulo} ja preenchido: {atual}")
            return True

        self.scroll_ate(driver, botao)
        self.js_click(driver, botao)

        alvo = valor.strip().lower()

        def achar_opcao(d):
            xpath_opcoes = (
                "//lightning-base-combobox-item"
                " | //div[@role='listbox']//*[@role='option']"
            )
            for opcao in d.find_elements(By.XPATH, xpath_opcoes):
                try:
                    if not opcao.is_displayed():
                        continue
                    texto = " ".join(opcao.text.split()).strip().lower()
                    dvalue = (opcao.get_attribute("data-value") or "").strip().lower()
                    if alvo in (texto, dvalue):
                        return opcao
                except StaleElementReferenceException:
                    continue
            return False

        try:
            opcao = self.criar_wait(driver, WAIT_PADRAO).until(achar_opcao)
            self.scroll_ate(driver, opcao)
            self.js_click(driver, opcao)
            self.log(f" [OK] {rotulo} definido como '{valor}'.")
            return True

        except TimeoutException:
            self.log(f" [AVISO] Opcao '{valor}' nao encontrada em '{rotulo}'.")
            try:
                botao.send_keys(Keys.ESCAPE)
            except Exception:
                pass
            return False

    def completar_endereco_por_cep(self, driver):
        """
        Etapa CPE: o Salvar trava quando Cidade/Estado/Pais estao vazios.

        Le o CEP da tela, consulta a API e preenche SOMENTE os campos
        que estiverem vazios. Nunca sobrescreve dado ja existente.
        """

        self.log(" [ACAO] Verificando endereco (CEP)...")

        campo_cep, valor_cep = self.ler_campo_label(
            driver, "CEP", timeout=WAIT_CURTO
        )

        if campo_cep is None:
            self.log(" [INFO] Campo CEP nao existe nesta tela. Pulando.")
            return "Endereco nao aplicavel"

        if not self.so_digitos(valor_cep):
            self.log(" [AVISO] CEP vazio na tela. Endereco nao sera completado.")
            return "CEP vazio"

        endereco = self.consultar_cep(valor_cep)

        if not endereco:
            return "CEP nao localizado na API"

        preenchidos = []

        # ---------- Cidade ----------

        campo_cidade, valor_cidade = self.ler_campo_label(
            driver, "Cidade", timeout=WAIT_CURTO
        )

        if campo_cidade is not None and not valor_cidade and endereco["cidade"]:
            self.scroll_ate(driver, campo_cidade)
            self.preencher_campo(campo_cidade, endereco["cidade"])
            self.log(f" [OK] Cidade preenchida: {endereco['cidade']}")
            preenchidos.append("Cidade")
        elif valor_cidade:
            self.log(f" [INFO] Cidade ja preenchida: {valor_cidade}")

        # ---------- Logradouro ----------

        campo_log, valor_log = self.ler_campo_label(
            driver, "Logradouro", timeout=WAIT_CURTO
        )

        if campo_log is not None and not valor_log and endereco["logradouro"]:
            self.scroll_ate(driver, campo_log)
            self.preencher_campo(campo_log, endereco["logradouro"])
            self.log(f" [OK] Logradouro preenchido: {endereco['logradouro']}")
            preenchidos.append("Logradouro")

        # ---------- Bairro ----------

        campo_bairro, valor_bairro = self.ler_campo_label(
            driver, "Bairro", timeout=WAIT_CURTO
        )

        if campo_bairro is not None and not valor_bairro and endereco["bairro"]:
            self.scroll_ate(driver, campo_bairro)
            self.preencher_campo(campo_bairro, endereco["bairro"])
            self.log(f" [OK] Bairro preenchido: {endereco['bairro']}")
            preenchidos.append("Bairro")

        # ---------- Estado (picklist) ----------

        if endereco["estado"]:
            if self.selecionar_picklist_valor(driver, "Estado", endereco["estado"]):
                preenchidos.append("Estado")
            elif endereco["uf"]:
                # algumas orgs listam a sigla em vez do nome
                if self.selecionar_picklist_valor(driver, "Estado", endereco["uf"]):
                    preenchidos.append("Estado")

        # ---------- Pais (picklist) ----------

        if self.selecionar_picklist_valor(driver, "País", "Brasil"):
            preenchidos.append("País")

        if preenchidos:
            resumo = ", ".join(preenchidos)
            self.log(f" [OK] Endereco completado via CEP: {resumo}")
            return f"Endereco completado ({resumo})"

        self.log(" [INFO] Endereco ja estava completo.")
        return "Endereco ja completo"

    # ========================================================

    def clicar_botao_editar(self, driver):
        """
        Clica no botao 'Editar' da aba Contato.

        Sem este clique a aba abre em MODO LEITURA e os campos
        (telefone, consentimentos) nao aceitam alteracao.

        HTML de referencia:
          <li data-target-selection-name="sfdc:StandardButton.Account.Edit">
            <runtime_platform_actions-action-renderer apiname="Edit" title="Editar">
              <button name="Edit" class="slds-button slds-button_neutral">

        Depois do clique aguarda o formulario de edicao (SaveEdit) aparecer.
        """

        seletores = [
            "//li[@data-target-selection-name='sfdc:StandardButton.Account.Edit']//button",
            "//runtime_platform_actions-action-renderer[@apiname='Edit']//button",
            "//runtime_platform_actions-action-renderer[@title='Editar']//button",
            "//button[@name='Edit']",
            "//button[@title='Editar']",
            "//button[normalize-space(text())='Editar']",
            "//a[@title='Editar']",
            "//*[@role='button'][normalize-space(.)='Editar']",
        ]

        self.log(" [ACAO] Clicando no botao Editar...")

        for xpath in seletores:

            try:
                elementos = driver.find_elements(By.XPATH, xpath)
            except Exception:
                continue

            for elemento in elementos:

                try:
                    if not elemento.is_displayed():
                        continue
                    if elemento.get_attribute("aria-disabled") == "true":
                        continue
                    if not elemento.is_enabled():
                        continue
                except StaleElementReferenceException:
                    continue

                self.scroll_ate(driver, elemento)

                # Estrategia 1: clique nativo
                try:
                    elemento.click()
                except Exception:
                    pass
                if self.modo_edicao_ativo(driver):
                    self.log(" [OK] Modo de edicao aberto (clique nativo).")
                    return True

                # Estrategia 2: clique via JavaScript
                try:
                    self.js_click(driver, elemento)
                except Exception:
                    pass
                if self.modo_edicao_ativo(driver):
                    self.log(" [OK] Modo de edicao aberto (JS click).")
                    return True

                # Estrategia 3: eventos de mouse reais
                try:
                    driver.execute_script(
                        """
                        const el = arguments[0];
                        ['pointerdown','mousedown','pointerup','mouseup','click']
                          .forEach(function (t) {
                            el.dispatchEvent(new MouseEvent(t, {
                                bubbles: true, cancelable: true, view: window
                            }));
                        });
                        """,
                        elemento
                    )
                except Exception:
                    pass
                if self.modo_edicao_ativo(driver):
                    self.log(" [OK] Modo de edicao aberto (eventos de mouse).")
                    return True

        self.log(" [ERRO] Botao Editar nao encontrado ou nao abriu o formulario.")
        return False

    def modo_edicao_ativo(self, driver):
        """
        Confirma que o formulario de edicao realmente abriu:
        botao Salvar visivel OU campo de telefone editavel na tela.
        """

        xpath_prova = (
            "//button[@name='SaveEdit']"
            " | //button[@title='Salvar']"
            " | //div[contains(@class,'slds-modal')]//button[@name='SaveEdit']"
            " | //input[@name='Phone'][not(@disabled)][not(@readonly)]"
        )

        try:
            self.criar_wait(driver, WAIT_PADRAO).until(
                EC.visibility_of_element_located((By.XPATH, xpath_prova))
            )
            return True
        except TimeoutException:
            return False

    def xpath_telefone(self):
        return (
            "//input[@name='Phone']"
            " | //input[contains(@name,'Phone')]"
            " | //lightning-input[.//label[contains(.,'Telefone')]]//input"
        )

    def preencher_telefone(self, driver, telefone):

        if not telefone:
            self.log(" [INFO] Sem telefone na planilha. Pulando.")
            return False

        try:
            campo = self.esperar_elemento(driver, self.xpath_telefone(), timeout=WAIT_PADRAO)
            self.scroll_ate(driver, campo)
            self.preencher_campo(campo, telefone)
            self.log(f" [OK] Telefone preenchido: {telefone}")
            return True
        except TimeoutException:
            self.log(" [AVISO] Campo Telefone nao encontrado.")
            return False
        except Exception as e:
            self.log(f" [AVISO] Falha ao preencher telefone: {e}")
            return False

    # ========================================================
    # CONSENTIMENTOS
    # ========================================================

    def marcar_consentimentos(self, driver, wait):
        """WhatsApp / Ligacao / Mensagem como 'Sim'. Falha isolada nao derruba o resto."""

        campos = [
            ("//button[@aria-label='Autoriza Receber Whatsapp?']", "WhatsApp"),
            ("//button[@aria-label='Autoriza Receber Ligação']", "Ligacao"),
            ("//button[@aria-label='Autoriza Receber Mensagem']", "Mensagem"),
        ]

        for xpath, descricao in campos:
            try:
                self.selecionar_sim(driver, wait, xpath, descricao)
            except TimeoutException:
                self.log(f" [AVISO] Campo '{descricao}' nao encontrado.")
            except Exception as e:
                self.log(f" [AVISO] Falha em '{descricao}': {e}")

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

                # ---------------------------------------------
                # ENDERECO VIA CEP (destrava o Salvar)
                # ---------------------------------------------

                obs_endereco = self.completar_endereco_por_cep(driver)
                self.log(f" [INFO] Endereco: {obs_endereco}")

            # =================================================
            # ROTINA ORIGEM DIVERSA
            # =================================================

            else:

                self.log(f" [INFO] Executando rotina para origem: {origem}")

                # 1) ATUALIZAR DADOS COM RFB (exclusivo nao-CPE)
                self.atualizar_dados_rfb(driver)

                # 2) ABA CONTATO (clique blindado)
                if not self.clicar_aba_contato(driver):
                    self.log(" [AVISO] Aba Contato nao abriu. Tentando no formulario atual.")


                # 3) BOTAO EDITAR (sem ele os campos ficam em modo leitura)
                if not self.clicar_botao_editar(driver):
                    self.log(" [AVISO] Editar nao abriu. Os campos podem estar bloqueados.")

                # 4) TELEFONE + CONSENTIMENTOS (Principal x Alternativo 1)
                obs_contato = self.tratar_contatos(driver, wait, telefone_limpo)
                self.log(f" [INFO] Regra de contato aplicada: {obs_contato}")

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