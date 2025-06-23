import time
import os
import pyautogui
import pygetwindow as gw
from datetime import datetime 
import subprocess 
import glob 
import shutil 

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
caminho_base_organizacao = r"C:\Users\Gerando Falcões\Documents\Python Projects\Automacao_Dashboards\RelatoriosDiarios" 

REMOTE_DEBUGGING_PORT = 9222 

driver = None
caminho_pasta_do_dia_base = "" 
caminho_download_temp_do_dia = "" 

# --- Funções ---

def entrar_tela_cheia():
    print("[DEBUG] Entrando em tela cheia (F11)...")
    pyautogui.press('f11')
    time.sleep(3) 

def configurar_pastas_do_dia_base():
    global caminho_pasta_do_dia_base, caminho_download_temp_do_dia
    
    hoje = datetime.now().strftime("%d-%m-%Y") 
    caminho_pasta_do_dia_base = os.path.join(caminho_base_organizacao, hoje)
    caminho_download_temp_do_dia = os.path.join(caminho_pasta_do_dia_base, "temp_downloads")

    os.makedirs(caminho_pasta_do_dia_base, exist_ok=True)
    os.makedirs(caminho_download_temp_do_dia, exist_ok=True)
    
    print(f"[INFO] Limpando pasta temporária de downloads: {caminho_download_temp_do_dia}")
    for f in glob.glob(os.path.join(caminho_download_temp_do_dia, '*')):
        try:
            if os.path.isfile(f) or os.path.islink(f):
                os.unlink(f)
            elif os.path.isdir(f):
                shutil.rmtree(f)
        except Exception as e:
            print(f'[ALERTA] Falha ao excluir {f}. Motivo: {e}')

    print(f"[INFO] Pasta base do dia: {caminho_pasta_do_dia_base}")
    print(f"[INFO] Pasta temporária de downloads (limpa): {caminho_download_temp_do_dia}")
    return caminho_download_temp_do_dia

def configurar_pasta_dashboard(dashboard_title):
    nome_pasta_dashboard = "".join(c for c in dashboard_title if c.isalnum() or c in (' ', '.', '_', '-')).strip()
    nome_pasta_dashboard = nome_pasta_dashboard.replace(" ", "_") 

    caminho_dashboard = os.path.join(caminho_pasta_do_dia_base, nome_pasta_dashboard)
    caminho_dados_exportados = os.path.join(caminho_dashboard, "Dados_Exportados") 

    os.makedirs(caminho_dashboard, exist_ok=True)
    os.makedirs(caminho_dados_exportados, exist_ok=True)
    
    print(f"[INFO] Estrutura de pastas para dashboard '{dashboard_title}' criada/verificada.")
    return caminho_dashboard, caminho_dados_exportados

def tirar_print_e_salvar(indice_dashboard, caminho_destino_print):
    nome_arquivo = f"dashboard_{indice_dashboard:02d}.png" 
    caminho_completo = os.path.join(caminho_destino_print, nome_arquivo)
    
    print(f"[INFO] Tirando print do dashboard {indice_dashboard} e salvando em: {caminho_completo}")
    try:
        screenshot = pyautogui.screenshot() 
        screenshot.save(caminho_completo)
        print("[OK] Print salvo com sucesso!")
        return caminho_completo
    except Exception as e:
        print(f"[ERRO] Erro ao tirar ou salvar print: {e}")
        return None

def mover_arquivos_exportados(caminho_origem_temp, caminho_destino_dados):
    print(f"[INFO] Movendo arquivos de '{caminho_origem_temp}' para '{caminho_destino_dados}'...")
    arquivos_movidos = []
    try:
        time.sleep(5) 
        
        for filename in os.listdir(caminho_origem_temp):
            origem = os.path.join(caminho_origem_temp, filename)
            destino = os.path.join(caminho_destino_dados, filename)
            
            if os.path.isfile(origem): 
                base, ext = os.path.splitext(filename)
                counter = 1
                while os.path.exists(destino):
                    destino = os.path.join(caminho_destino_dados, f"{base}_{counter}{ext}")
                    counter += 1
                
                shutil.move(origem, destino)
                arquivos_movidos.append(filename)
                print(f"[DEBUG] Movido: {filename} para {os.path.basename(destino)}")
        print(f"[OK] {len(arquivos_movidos)} arquivo(s) movido(s) para {caminho_destino_dados}.")
        return arquivos_movidos
    except Exception as e:
        print(f"[ERRO] Erro ao mover arquivos exportados: {e}")
        return []

def conectar_ao_chrome_existente(links, download_temp_dir):
    """
    Tenta se conectar a uma instância do Chrome já aberta em modo de depuração.
    Se conseguir, navega e abre todos os dashboards.
    A configuração do diretório de download DEVE ser feita manualmente no perfil do Chrome.
    """
    global driver 

    print(f"[DEBUG] Tentando conectar ao Chrome existente na porta {REMOTE_DEBUGGING_PORT}...")
    options = webdriver.ChromeOptions()
    options.add_experimental_option("debuggerAddress", f"127.0.0.1:{REMOTE_DEBUGGING_PORT}")
    
    # --- Estas linhas foram COMENTADAS/REMOVIDAS para resolver o erro 'unrecognized chrome option: excludeSwitches' ---
    # options.add_experimental_option("excludeSwitches", ["enable-automation"])
    # options.add_experimental_option('useAutomationExtension', False)

    try:
        driver = webdriver.Chrome(options=options)
        print("[OK] Conectado ao Chrome existente.")
        
        print("[DEBUG] Limpando abas existentes e abrindo novas dos dashboards...")
        all_handles = driver.window_handles
        
        for handle in reversed(all_handles[1:]): 
            driver.switch_to.window(handle)
            driver.close()
        
        driver.switch_to.window(all_handles[0])

        primeiro_link = True
        for link in links:
            if primeiro_link:
                driver.get(link) 
                primeiro_link = False
            else:
                driver.execute_script("window.open('');") 
                driver.switch_to.window(driver.window_handles[-1]) 
                driver.get(link) 
            
            print(f"[INFO] Dashboard aberto: {link}")
            WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            time.sleep(10) 

        driver.switch_to.window(driver.window_handles[0]) 
        print("[OK] Todos os dashboards abertos. Foco na primeira aba.")
        
        driver.maximize_window()
        
        return True
    except Exception as e:
        print(f"[ERRO] Erro ao conectar ou abrir dashboards: {e}")
        print("[ALERTA] Certifique-se de que o Chrome está aberto com '--remote-debugging-port=9222' e sem outras instâncias usando o mesmo perfil/porta.")
        return False

def exportar_dados_de_todos_os_visuais(caminho_destino_dados, download_temp_dir):
    print("[INFO] Iniciando exportação de dados de todos os visuais no dashboard atual...")
    
    # --- *** VOCÊ PRECISA FORNECER ESTES SELETORES COM BASE NA SUA INSPEÇÃO MANUAL (F12) *** ---
    # Estes são exemplos comuns. Use os que VOCÊ ENCONTROU para seu Power BI.

    SELETOR_VISUAL_CONTAINER_GENERICO = 'div.visualContainer[role="group"]'
    SELETOR_BOTAO_MAIS_OPCOES = 'button[aria-label="Mais opções"]' 
    SELETOR_OPCAO_EXPORTAR_DADOS = "//button[contains(.,'Exportar dados')]"
    SELETOR_BOTAO_CONFIRMAR_EXPORTACAO = "//button[contains(.,'Exportar')]" 

    # --- FIM DA SUA INTERVENÇÃO ---

    visual_containers = driver.find_elements(By.CSS_SELECTOR, SELETOR_VISUAL_CONTAINER_GENERICO) 
    
    if not visual_containers:
        print("[ALERTA] Nenhum visual encontrado para exportar neste dashboard com o seletor padrão. Verifique 'SELETOR_VISUAL_CONTAINER_GENERICO'.")
        return

    print(f"[INFO] Encontrados {len(visual_containers)} potenciais visuais para tentar exportar.")

    for i, visual_container in enumerate(visual_containers):
        try:
            print(f"[DEBUG] Tentando exportar visual {i+1}...")
            
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", visual_container)
            time.sleep(1) 

            more_options_button = WebDriverWait(visual_container, 5).until( 
                EC.element_to_be_clickable((By.CSS_SELECTOR, SELETOR_BOTAO_MAIS_OPCOES)) 
            )
            more_options_button.click()
            print(f"[DEBUG] Clicou nos 'Mais opções' do visual {i+1}.")
            time.sleep(1) 

            export_data_option = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, SELETOR_OPCAO_EXPORTAR_DADOS)) 
            )
            export_data_option.click()
            print(f"[DEBUG] Clicou em 'Exportar dados' para o visual {i+1}.")
            time.sleep(2) 

            try:
                confirm_export_button = WebDriverWait(driver, 3).until( 
                    EC.element_to_be_clickable((By.XPATH, SELETOR_BOTAO_CONFIRMAR_EXPORTACAO)) 
                )
                confirm_export_button.click()
                print(f"[DEBUG] Clicou no botão de confirmação de exportação para visual {i+1}.")
                time.sleep(3) 
            except:
                print(f"[DEBUG] Nenhum botão de confirmação de exportação para visual {i+1} encontrado ou clicável, presumindo download direto ou não aplicável.")
                time.sleep(5) 

            print(f"[OK] Exportação de dados iniciada para o visual {i+1}.")

        except Exception as e:
            print(f"[ALERTA] Não foi possível exportar dados do visual {i+1} no dashboard atual: {e}")
            pyautogui.press('escape') 
            time.sleep(1)

    print("[INFO] Fim da tentativa de exportação de visuais para o dashboard atual. Movendo arquivos...")
    mover_arquivos_exportados(download_temp_dir, caminho_destino_dados)

def trocar_abas_loop_selenium(links):
    print("[DEBUG] Iniciando loop de troca de abas, organização de pastas, prints e exportação...")
    
    num_dashboards = len(links)
    indice_atual_dashboard = 0 
    loop_count = 0 
    
    while True:
        try:
            download_temp_dir = configurar_pastas_do_dia_base()

            loop_count += 1
            
            driver.switch_to.window(driver.window_handles[indice_atual_dashboard])
            current_url = driver.current_url
            
            dashboard_title = driver.title 
            if " - Microsoft Power BI" in dashboard_title:
                dashboard_title = dashboard_title.replace(" - Microsoft Power BI", "").strip()
            if not dashboard_title or dashboard_title in ["Power BI", "Visual"]: 
                report_id_match = current_url.split("reportId=")[-1].split("&")[0]
                dashboard_title = f"Dashboard_{report_id_match[:8]}" 

            print(f"\n[INFO] Foco no dashboard: '{dashboard_title}' ({current_url})")
            
            caminho_prints_destino, caminho_dados_destino = configurar_pasta_dashboard(dashboard_title)

            time.sleep(5) 
            
            tirar_print_e_salvar(indice_atual_dashboard + 1, caminho_prints_destino) 

            exportar_dados_de_todos_os_visuais(caminho_dados_destino, download_temp_dir)
            
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
    
    os.makedirs(caminho_base_organizacao, exist_ok=True) 

    download_temp_dir_global = configurar_pastas_do_dia_base()

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

    if not conectar_ao_chrome_existente(links, download_temp_dir_global):
        print("[FIM] Não foi possível conectar ao Chrome ou abrir os dashboards. Encerrando.")
        return
    
    print("[DEBUG] Ativando a janela do Chrome e colocando em tela cheia...")
    tempo_espera_janela = 0
    janela_encontrada = False
    while tempo_espera_janela < 30: 
        janelas_chrome = gw.getWindowsWithTitle("Chrome")
        if janelas_chrome:
            selenium_window = janelas_chrome[0] 
            
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
        return

    time.sleep(3) 
    
    trocar_abas_loop_selenium(links)

    print("[INFO] Automação concluída ou interrompida. O navegador Chrome permanecerá aberto.")

if __name__ == "__main__":
    main()