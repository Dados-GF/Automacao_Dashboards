import subprocess
import time
import os
import pyautogui
import pygetwindow as gw

caminho_chrome = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
caminho_lista = r"C:\Users\Gerando Falcões\Documents\Python Projects\Automacao_Dashboards\lista_dashboards.txt"
tempo_entre_abas = 60
intervalo_refresh = 5

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
        time.sleep(1)
    print("[ERRO] Chrome não detectado.")
    return None

def entrar_tela_cheia():
    print("[DEBUG] Entrando em tela cheia...")
    time.sleep(10)
    pyautogui.press('f11')

def trocar_abas_loop():
    print("[DEBUG] Iniciando loop de troca de abas + refresh...")
    contador = 0
    while True:
        contador += 1
        pyautogui.press('f11')
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'tab')
        time.sleep(1)
        pyautogui.press('f11')

        if contador % intervalo_refresh == 0:
            print("[DEBUG] Dando refresh (F5)")
            pyautogui.press('f5')
            time.sleep(2)

        print(f"[INFO] Aguardando {tempo_entre_abas} segundos...")
        time.sleep(tempo_entre_abas)

def main():
    print("=== INICIANDO SCRIPT DASHBOARDS ===")
    links = abrir_abas_no_chrome()
    if not links:
        print("[FIM] Nenhum link válido. Encerrando.")
        return

    janela = esperar_janela_carregar()
    if janela:
        janela.activate()
        time.sleep(2)
        entrar_tela_cheia()
        time.sleep(3)
        trocar_abas_loop()

if __name__ == "__main__":
    main()
