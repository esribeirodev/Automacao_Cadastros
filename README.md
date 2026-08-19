# Automação de Cadastros - Salesforce (FOCO)

Este projeto é um bot com interface gráfica desenvolvido para realizar a atualização massiva de dados cadastrais e de consentimento de clientes no ambiente Salesforce (FOCO). Ele utiliza o Selenium para interagir de forma invisível e cirúrgica com a interface web do CRM e o CustomTkinter para fornecer um painel de controle amigável e à prova de falhas.

## 🚀 Funcionalidades
* **Interface Gráfica Moderna:** Painel limpo com tela de instruções e console de logs em tempo real.
* **Processamento Assíncrono:** Uso nativo de *threads* para evitar que a interface trave ou congele durante a execução do Selenium.
* **Bypass de LWC (Salesforce):** Utiliza cliques via JavaScript (`js_click`) e rolagem focada na tela (`scroll_ate`) para evitar o bloqueio de cliques pela barra de utilitários flutuante do Salesforce (erro *element click intercepted*).
* **Tratamento Dinâmico de CPF:** Aceita CPFs puros (só números) ou formatados com pontuação, corrigindo automaticamente as anomalias numéricas geradas pelo Pandas.
* **Sistema de Salvamento de Dupla Tentativa:** Tenta clicar no botão de salvar duas vezes de forma inteligente caso o sistema apresente lentidão ou levante algum modal de validação de negócio no primeiro clique.

## 📋 Pré-requisitos
Certifique-se de ter o Python 3 instalado na sua máquina. Em seguida, instale as bibliotecas necessárias rodando o comando abaixo no terminal:

```bash
pip install pandas selenium customtkinter openpyxl