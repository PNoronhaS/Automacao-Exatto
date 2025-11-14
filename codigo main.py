import pyautogui
import time
import subprocess
import os
import pandas as pd
from datetime import datetime, timedelta
from openpyxl import load_workbook
import shutil

# Configuração das lojas
lojas = {
    "Mooca": {
        "login": "ljmooca", "senha": "luc1549", "setinhas": 9, "tecla_inicial": "m"
    },
    "Lorena": {
        "login": "LJALORENA", "senha": "luc1549", "setinhas": 8, "tecla_inicial": "m"
    },
    "Alto de Pinheiros": {
        "login": "LJALTOPINHEIROS", "senha": "AVATIM231", "setinhas": 6, "tecla_inicial": "m"
    },
    "Ibirapuera": {
        "login": "LJIBIRAPUERA", "senha": "AVATIM123", "setinhas": 7, "tecla_inicial": "f"
    },
    "Perdizes": {
        "login": "LJPERDIZES", "senha": "LUC1549", "setinhas": 10, "tecla_inicial": "f"
    },
    "Leopoldina": {
        "login": "LJVLEOPOLDINA", "senha": "AVATIM123", "setinhas": 12, "tecla_inicial": "f"
    },
    "Santana": {
        "login": "LJF2SANTANA", "senha": "AVATIM123", "setinhas": 11, "tecla_inicial": "f"
    }
}

# Caminho do Chrome
chrome_path = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
exatto_url = "https://portalfranquias.avatim.com.br/"

# Caminho da planilha F360
planilha_f360 = r"C:\Users\PERFIL\Downloads\Planilha F360.xlsx"
caminho_downloads = r"C:\Users\PERFIL\Downloads"

# Calcula datas automaticamente
hoje = datetime.today()
ontem = hoje - timedelta(days=1)

if hoje.weekday() == 0:  # segunda-feira
    sexta = hoje - timedelta(days=3)
    domingo = hoje - timedelta(days=1)
    DATA_INICIAL = sexta.strftime("%d/%m/%Y")
    DATA_FINAL = domingo.strftime("%d/%m/%Y")
else:
    DATA_INICIAL = ontem.strftime("%d/%m/%Y")
    DATA_FINAL = ontem.strftime("%d/%m/%Y")

print(f"📅 Período: {DATA_INICIAL} até {DATA_FINAL}")

# -------------------------------
# Função para processar uma loja
# -------------------------------
def processar_loja(nome_loja, login, senha, setinhas, tecla_inicial):
    print(f"\n🔄 Processando loja: {nome_loja}")

    # Abre navegador
    subprocess.Popen([chrome_path, exatto_url])
    time.sleep(3.5)

    # Login
    pyautogui.click(x=582, y=427)
    pyautogui.hotkey("ctrl", "a"); pyautogui.press("backspace")
    pyautogui.write(login)

    pyautogui.click(x=587, y=487)
    pyautogui.hotkey("ctrl", "a"); pyautogui.press("backspace")
    pyautogui.write(senha)

    pyautogui.click(x=610, y=549)
    time.sleep(0.5)
    pyautogui.press(tecla_inicial)   # usa 'm' ou 'f' conforme a loja
    time.sleep(0.5)
    for _ in range(setinhas):
        pyautogui.press("down")
        time.sleep(0.1)
    pyautogui.press("enter")

    pyautogui.click(x=1035, y=549)
    time.sleep(2)

    # Consultas > Controle de Caixa
    pyautogui.click(x=276, y=253)  # Botão Consultas
    time.sleep(1)
    pyautogui.click(x=324, y=504)  # Botão Controle de Caixa
    time.sleep(2)

    # Preenche datas
    pyautogui.click(x=200, y=458)  # Data Inicial
    pyautogui.hotkey("ctrl", "a"); pyautogui.press("backspace")
    pyautogui.write(DATA_INICIAL)

    pyautogui.click(x=198, y=483)  # Data Final
    pyautogui.hotkey("ctrl", "a"); pyautogui.press("backspace")
    pyautogui.write(DATA_FINAL)

    # Pesquisar
    pyautogui.click(x=1000, y=560)
    time.sleep(3)

    # Rola a tela
    pyautogui.moveTo(x=1300, y=600)
    time.sleep(0.5)
    for _ in range(2):
        pyautogui.scroll(-1000)
        time.sleep(1)

    # Exportar
    pyautogui.click(x=1052, y=482)
    time.sleep(5)  # espera o download terminar

    # -------------------------------
    # 📊 Processamento com pandas
    # -------------------------------
    arquivo_exportado = os.path.join(caminho_downloads, "Result.xls")
    tables = pd.read_html(arquivo_exportado, header=0)
    df = tables[0]
    df = df.drop(columns=["ABERTURA", "FECHAMENTO"], errors="ignore")

    df_f360 = df[["DATA", "VALOR", "FORMA DE PAGAMENTO", "CLIENTE/FORNECEDOR", "OPERADOR", "NOTA FISCAL"]]
    df_f360 = df_f360.rename(columns={
        "DATA": "Data Venda",
        "VALOR": "Valor Bruto",
        "FORMA DE PAGAMENTO": "Forma de Pagamento",
        "CLIENTE/FORNECEDOR": "Nome do Cliente",
        "OPERADOR": "Operador",
        "NOTA FISCAL": "Nota Fiscal"
    })

    book = load_workbook(planilha_f360)
    sheet = book["Modelo de Importação de GFF-PDV"]

    # Limpa linhas antigas
    max_row = sheet.max_row
    for row in range(3, max_row+1):
        for col in range(1, sheet.max_column+1):
            sheet.cell(row=row, column=col).value = None

    # Insere novos dados
    start_row = 3
    for i, row in df_f360.iterrows():
        sheet.cell(row=start_row+i, column=1, value=row["Data Venda"])
        sheet.cell(row=start_row+i, column=2, value=row["Valor Bruto"])
        sheet.cell(row=start_row+i, column=3, value=row["Forma de Pagamento"])
        sheet.cell(row=start_row+i, column=12, value=row["Nome do Cliente"])
        sheet.cell(row=start_row+i, column=15, value=row["Operador"])
        sheet.cell(row=start_row+i, column=16, value=row["Nota Fiscal"])

    book.save(planilha_f360)

    # Cria cópia com nome da loja + data
    data_str = datetime.today().strftime("%d-%m-%Y")
    copia_planilha = os.path.join(caminho_downloads, f"{nome_loja}_{data_str}.xlsx")
    shutil.copy(planilha_f360, copia_planilha)

    print(f"✅ Loja {nome_loja} processada. Cópia criada: {copia_planilha}")


# -------------------------------
# Loop para todas as lojas
# -------------------------------
for nome_loja, config in lojas.items():
    processar_loja(
        nome_loja,
        config["login"],
        config["senha"],
        config["setinhas"],
        config["tecla_inicial"]
    )

