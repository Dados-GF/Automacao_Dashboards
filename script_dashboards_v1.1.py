import time
import os
import pyautogui
import pygetwindow as gw
from datetime import datetime

# Importações do Selenium
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# --- Configurações ---
caminho_lista = r"C:\Users\Gerando Falcões\Documents\Python Projects\Automacao_Dashboards\lista_dashboards.txt"
tempo_entre_abas = 60
intervalo_refresh = 5 
caminho_base_prints = r"C:\Users\Gerando Falcões\Documents\Python Projects\Automacao_Dashboards\prints"

# Remover as configurações de CHROME_USER_DATA_DIR e CHROME_PROFILE_DIR, não serão mais usadas.
# CHROME_USER_DATA_DIR = r"C:\Users\SeuUsuario\AppData\Local\Google\Chrome\User Data" 
# CHROME_PROFILE_DIR = "Profile 1" 

# Variável global para armazenar o driver do Selenium
driver = None

# Variável global para armazenar o caminho da pasta de prints do dia
caminho_prints_do_dia = ""

# --- Funções ---

def configurar_pasta_prints_diaria():
    """Configura e retorna o caminho da pasta de prints para o dia atual."""
    global caminho_prints_do_dia
    
    hoje = datetime.now().strftime("%d-%m-%Y") 
    caminho_pasta_do_dia = os.path.join(caminho_base_prints, hoje)

    if not os.path.exists(caminho_pasta_do_dia):
        print(f"[INFO] Criando nova pasta de prints para {hoje}: {caminho_pasta_do_dia}")
        os.makedirs(caminho_pasta_do_dia)
    else:
        print(f"[INFO] Usando pasta de prints existente para {hoje}: {caminho_pasta_do_dia}")
    
    caminho_prints_do_dia = caminho_pasta_do_dia

def tirar_print_e_salvar(indice_dashboard):
    """
    Tira um print da tela inteira e o salva em um arquivo com um índice numérico
    dentro da pasta de prints do dia.
    """
    if not caminho_prints_do_dia:
        print("[ERRO] Caminho da pasta de prints diária não configurado.")
        return None

    nome_arquivo = f"dashboard_{indice_dashboard:02d}.png" 
    caminho_completo = os.path.join(caminho_prints_do_dia, nome_arquivo)
    
    print(f"[INFO] Tirando print do dashboard {indice_dashboard} e salvando em: {caminho_completo}")
    try:
        screenshot = pyautogui.screenshot() 
        screenshot.save(caminho_completo)
        print("[OK] Print salvo com sucesso!")
        return caminho_completo
    except Exception as e:
        print(f"[ERRO] Erro ao tirar ou salvar print: {e}")
        return None

def abrir_dashboards_com_selenium(links):
    """
    Inicia uma nova instância do Chrome com Selenium no modo anônimo
    e abre todos os dashboards em abas separadas.
    """
    global driver 

    print("[DEBUG] Iniciando o serviço do ChromeDriver...")
    service = Service(ChromeDriverManager().install())
    
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")       # Inicia maximizado
    options.add_argument("--incognito")             # ABRE NO MODO ANÔNIMO
    options.add_argument("--disable-infobars")      # Desabilita a barra "Chrome está sendo controlado por software de teste"
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)

    try:
        print("[DEBUG] Abrindo Chrome com Selenium no modo anônimo...")
        driver = webdriver.Chrome(service=service, options=options)
        
        primeiro_link = True
        for link in links:
            if primeiro_link:
                driver.get(link) # Abre o primeiro link na aba inicial
                primeiro_link = False
            else:
                # Abre uma nova aba e muda o foco para ela
                driver.execute_script("window.open('');")
                driver.switch_to.window(driver.window_handles[-1])
                driver.get(link) # Abre o link na nova aba
            
            print(f"[INFO] Dashboard aberto: {link}")
            # Espera a página carregar (pode ajustar o tempo ou usar WebDriverWait para elementos específicos)
            WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            time.sleep(5) # Tempo adicional para o Power BI renderizar

        # Retorna para a primeira aba para começar o ciclo
        driver.switch_to.window(driver.window_handles[0]) 
        print("[OK] Todos os dashboards abertos. Foco na primeira aba.")
        return True
    except Exception as e:
        print(f"[ERRO] Erro ao abrir dashboards com Selenium: {e}")
        if driver:
            driver.quit() 
        return False

def entrar_tela_cheia():
    print("[DEBUG] Entrando em tela cheia (F11)...")
    pyautogui.press('f11')
    time.sleep(3) 

def trocar_abas_loop_selenium(links):
    """
    Loop principal para trocar entre as abas do navegador controlado pelo Selenium,
    tirar prints e dar refresh.
    """
    print("[DEBUG] Iniciando loop de troca de abas, prints e refresh...")
    
    num_dashboards = len(links)
    indice_atual_dashboard = 0 
    loop_count = 0 
    
    while True:
        try:
            configurar_pasta_prints_diaria()

            loop_count += 1
            
            # Mudar o foco para a aba correta usando Selenium
            driver.switch_to.window(driver.window_handles[indice_atual_dashboard])
            print(f"[INFO] Foco no dashboard: {links[indice_atual_dashboard]}")
            
            time.sleep(5) 
            
            tirar_print_e_salvar(indice_atual_dashboard + 1) 
            
            if loop_count % intervalo_refresh == 0:
                print("[DEBUG] Dando refresh (F5) na aba atual via Selenium.")
                driver.refresh() 
                time.sleep(5) 

            indice_atual_dashboard = (indice_atual_dashboard + 1) % num_dashboards
            
            print(f"[INFO] Aguardando {tempo_entre_abas} segundos para o próximo ciclo (dashboard {indice_atual_dashboard + 1})...")
            time.sleep(tempo_entre_abas)

        except KeyboardInterrupt:
            print("\n[INFO] Automação interrompida pelo usuário (Ctrl+C). Encerrando.")
            break 
        except Exception as e:
            print(f"[ERRO] Ocorreu um erro inesperado no loop principal: {e}")
            print("[INFO] Tentando continuar, mas verifique o problema.")
            time.sleep(10) 

def main():
    print("=== INICIANDO SCRIPT DASHBOARDS ===")
    
    os.makedirs(caminho_base_prints, exist_ok=True)

    print("[DEBUG] Lendo links dos dashboards...")
    if not os.path.exists(caminho_lista):
        print("[ERRO] Lista de dashboards não encontrada. Encerrando.")
        return
    with open(caminho_lista, 'r', encoding='utf-8') as f:
        conteudo = f.read()
    links = [link.strip() for link in conteudo.split(';') if link.strip()]
    
    if not links:
        print("[FIM] Lista de links vazia ou mal formatada. Encerrando.")
        return

    if not abrir_dashboards_com_selenium(links):
        print("[FIM] Não foi possível abrir os dashboards. Encerrando.")
        return
    
    # Após abrir os dashboards com Selenium, ativa a janela e coloca em tela cheia
    # Pode haver um pequeno atraso para a janela do Selenium aparecer para o pygetwindow
    print("[DEBUG] Esperando a janela do Chrome do Selenium aparecer e ativando...")
    tempo_espera_janela = 0
    janela_encontrada = False
    while tempo_espera_janela < 30: # Tenta por até 30 segundos
        janelas_chrome = gw.getWindowsWithTitle("Chrome")
        if janelas_chrome:
            # Com --incognito, o título pode ser "Nova guia - Google Chrome (Incógnito)" ou similar.
            # É mais provável que a primeira janela do Chrome encontrada seja a do Selenium agora.
            selenium_window = None
            for janela in janelas_chrome:
                if "Chrome" in janela.title: # Apenas verifica por "Chrome" no título
                    selenium_window = janela
                    break
            
            if selenium_window:
                selenium_window.activate()
                time.sleep(2)
                entrar_tela_cheia()
                janela_encontrada = True
                break
        time.sleep(1)
        tempo_espera_janela += 1

    if not janela_encontrada:
        print("[FIM] Janela do Chrome controlada por Selenium não detectada ou ativada. Encerrando.")
        if driver: driver.quit()
        return

    time.sleep(3) # Tempo para a primeira aba carregar em tela cheia
    
    configurar_pasta_prints_diaria() 
    
    trocar_abas_loop_selenium(links)

    if driver:
        print("[INFO] Fechando navegador Chrome.")
        driver.quit()

if __name__ == "__main__":
    main()