import subprocess
import time
import os
import pyautogui
import pygetwindow as gw
from datetime import datetime # Para trabalhar com datas

# --- Configurações ---
caminho_chrome = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
caminho_lista = r"C:\Users\Gerando Falcões\Documents\Python Projects\Automacao_Dashboards\lista_dashboards.txt"
tempo_entre_abas = 60
intervalo_refresh = 5
caminho_base_prints = r"C:\Users\Gerando Falcões\Documents\Python Projects\Automacao_Dashboards\prints" # Caminho base para os prints

# Variável global para armazenar o caminho da pasta de prints do dia
caminho_prints_do_dia = ""

# --- Funções Existentes ---
def abrir_abas_no_chrome():
    print("[DEBUG] Verificando se lista existe...")
    if not os.path.exists(caminho_lista):
        print("[ERRO] Lista de dashboards não encontrada.")
        return []

    print("[DEBUG] Lendo links...")
    with open(caminho_lista, 'r', encoding='utf-8') as f:
        conteudo = f.read()

    links = [link.strip() for link in conteudo.split(';') if link.strip()]
    if not links:
        print("[ERRO] Lista vazia ou mal formatada.")
        return []

    print(f"[OK] Abrindo Chrome com os links: {links}")
    try:
        subprocess.Popen([caminho_chrome] + links)
    except Exception as e:
        print(f"[ERRO] Erro ao abrir o Chrome: {e}")
    return links

def esperar_janela_carregar():
    print("[DEBUG] Esperando Chrome aparecer...")
    for _ in range(30):
        janelas = gw.getWindowsWithTitle("Chrome")
        if janelas:
            print("[OK] Chrome detectado.")
            return janelas[0]
        time.sleep(3)
    print("[ERRO] Chrome não detectado.")
    return None

def entrar_tela_cheia():
    print("[DEBUG] Entrando em tela cheia...")
    time.sleep(10) # Dê um tempo para as abas carregarem antes de F11
    pyautogui.press('f11')
    time.sleep(3) # Tempo para a tela cheia se ajustar

# --- Nova Função: Gerenciar Pastas Diárias ---
def configurar_pasta_prints_diaria():
    global caminho_prints_do_dia # Usar a variável global
    
    hoje = datetime.now().strftime("%d-%m-%Y") # Formato DD-MM-AAAA
    caminho_pasta_do_dia = os.path.join(caminho_base_prints, hoje)

    if not os.path.exists(caminho_pasta_do_dia):
        print(f"[INFO] Criando nova pasta de prints para {hoje}: {caminho_pasta_do_dia}")
        os.makedirs(caminho_pasta_do_dia)
    else:
        print(f"[INFO] Usando pasta de prints existente para {hoje}: {caminho_pasta_do_dia}")
    
    caminho_prints_do_dia = caminho_pasta_do_dia # Atualiza o caminho global

# --- Função de Tirar Print Atualizada para usar a pasta diária ---
def tirar_print_e_salvar(indice_dashboard):
    """
    Tira um print da tela inteira e o salva em um arquivo com um índice numérico
    dentro da pasta de prints do dia.
    """
    if not caminho_prints_do_dia:
        print("[ERRO] Caminho da pasta de prints diária não configurado.")
        return None

    nome_arquivo = f"dashboard_{indice_dashboard:02d}.png" # Formato com 2 dígitos (01, 02, etc.)
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

# --- Função de Loop Atualizada com verificação diária e tratamento de interrupção ---
def trocar_abas_loop(num_dashboards):
    print("[DEBUG] Iniciando loop de troca de abas + refresh + prints com numeração...")
    
    indice_atual_dashboard = 0 
    loop_count = 0 
    
    while True:
        try:
            # A cada ciclo do loop principal, verificar se a data mudou
            # e reconfigurar a pasta de prints se necessário.
            configurar_pasta_prints_diaria()

            loop_count += 1
            
            # Tempo para a aba carregar e o conteúdo do Power BI aparecer
            time.sleep(5) 
            
            # Tirar print da tela atual com o índice do dashboard
            tirar_print_e_salvar(indice_atual_dashboard + 1) 
            
            # Alterna para a próxima aba
            pyautogui.hotkey('ctrl', 'tab') 
            time.sleep(1) # Pequeno atraso para a troca de aba
            
            # Entra em tela cheia novamente (às vezes a troca de aba pode sair do F11)
            pyautogui.press('f11')
            time.sleep(1) # Pequeno atraso para o F11
            
            # Lógica de refresh a cada 'intervalo_refresh' ciclos
            if loop_count % intervalo_refresh == 0:
                print("[DEBUG] Dando refresh (F5)")
                pyautogui.press('f5')
                time.sleep(5) # Tempo extra para o refresh carregar

            # Atualiza o índice do dashboard para o próximo
            indice_atual_dashboard = (indice_atual_dashboard + 1) % num_dashboards
            
            print(f"[INFO] Aguardando {tempo_entre_abas} segundos para o próximo ciclo (dashboard {indice_atual_dashboard + 1})...")
            time.sleep(tempo_entre_abas)

        except KeyboardInterrupt:
            print("\n[INFO] Automação interrompida pelo usuário (Ctrl+C). Encerrando.")
            break # Sai do loop while True

def main():
    print("=== INICIANDO SCRIPT DASHBOARDS ===")
    
    # Garante que o diretório base para prints exista antes de tudo
    os.makedirs(caminho_base_prints, exist_ok=True)

    links = abrir_abas_no_chrome()
    if not links:
        print("[FIM] Nenhum link válido. Encerrando.")
        return

    num_dashboards = len(links) 

    janela = esperar_janela_carregar()
    if janela:
        janela.activate()
        time.sleep(2)
        entrar_tela_cheia()
        time.sleep(3) # Tempo para a primeira aba carregar em tela cheia
        
        # Configura a pasta de prints do dia na inicialização
        configurar_pasta_prints_diaria() 
        
        trocar_abas_loop(num_dashboards) # Passa o número de dashboards para o loop
    else:
        print("[FIM] Janela do Chrome não encontrada. Encerrando.")

if __name__ == "__main__":
    main()