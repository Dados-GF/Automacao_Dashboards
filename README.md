# Automação de Dashboards com Chrome e PyAutoGUI versão 1.3

Este script Python automatiza a exibição de múltiplos dashboards em abas do Google Chrome. Ele alterna entre as abas, tira capturas de tela periodicamente e atualiza os dashboards para garantir a visualização dos dados mais recentes. É uma solução leve e eficaz para apresentações contínuas de dados em monitores dedicados, TVs ou em ambientes de escritório.

---

## 🚀 Funcionalidades

* **Abertura Rápida:** Inicia o Google Chrome e abre todos os links dos dashboards em abas separadas.
* **Capturas de Tela Diárias:** Organiza e salva capturas de tela de cada dashboard em pastas com a data do dia, facilitando o acompanhamento visual do histórico.
* **Navegação Cíclica:** Alterna automaticamente entre as abas dos dashboards em um intervalo configurável.
* **Atualização Periódica:** Recarrega os dashboards em intervalos definidos, garantindo que os dados exibidos estejam sempre atualizados.
* **Modo Tela Cheia:** Ativa e mantém o modo de tela cheia (F11) no navegador para uma experiência de visualização imersiva, mesmo após a troca de abas.

---

## 📋 Pré-requisitos

Antes de executar o script, certifique-se de ter o seguinte instalado:

* **Python 3.x:** Baixe e instale a versão mais recente em [python.org](https://www.python.org/).
* **Google Chrome:** O navegador deve estar instalado e o caminho para o executável precisa ser configurado no script.

---

## 🛠️ Instalação

Siga os passos abaixo para configurar o ambiente:

1.  **Baixe o Código:**
    Faça o download ou clone o repositório para uma pasta em seu computador.

2.  **Instale as Dependências:**
    Abra o terminal na pasta do projeto e instale as bibliotecas Python necessárias:

    ```bash
    pip install pyautogui pygetwindow
    ```

---

## ⚙️ Configuração

Antes de executar o script, você precisará ajustar alguns parâmetros. Abra o arquivo Python e edite a seção `--- Configurações ---`:

```python
# --- Configurações ---
caminho_chrome = r"C:\Program Files\Google\Chrome\Application\chrome.exe" # <--- VERIFIQUE E ATUALIZE ESTE CAMINHO
caminho_lista = r"C:\Users\Gerando Falcões\Documents\Python Projects\Automacao_Dashboards\lista_dashboards.txt" # <--- ATUALIZE ESTE CAMINHO
tempo_entre_abas = 60 # Tempo em segundos para alternar entre as abas
intervalo_refresh = 5 # A cada quantas trocas de aba um refresh será feito (ex: 5 significa a cada 5 dashboards visitados, o atual será recarregado)
caminho_base_prints = r"C:\Users\Gerando Falcões\Documents\Python Projects\Automacao_Dashboards\prints" # <--- ATUALIZE ESTE CAMINHO