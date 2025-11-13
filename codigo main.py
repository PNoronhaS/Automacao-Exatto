import pyautogui
import time
import subprocess

# Dados de login
EXATTO_USER = "ljmooca"
EXATTO_PASS = "luc1549"

# Caminho do Chrome — ajuste se necessário
chrome_path = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
exatto_url = "https://portalfranquias.avatim.com.br/"

# Abre o navegador no site do Exatto
subprocess.Popen([chrome_path, exatto_url])
print("⏳ Aguardando o navegador abrir...")
time.sleep(3.5)

# Campo usuário
pyautogui.click(x=582, y=427)
pyautogui.hotkey("ctrl", "a")
pyautogui.press("backspace")
pyautogui.write(EXATTO_USER)

# Campo senha
pyautogui.click(x=587, y=487)
pyautogui.hotkey("ctrl", "a")
pyautogui.press("backspace")
pyautogui.write(EXATTO_PASS)

# Campo loja
pyautogui.click(x=610, y=549)
time.sleep(0.5)
pyautogui.press("m")  # pula para lojas que começam com M
time.sleep(0.5)
for _ in range(9):    # 9 cliques para chegar na loja correta
    pyautogui.press("down")
    time.sleep(0.2)
pyautogui.press("enter")

# Botão OK
pyautogui.click(x=1035, y=549)

print("✅ Login enviado com sucesso!")
