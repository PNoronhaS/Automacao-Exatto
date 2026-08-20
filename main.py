import pyautogui
import time
import subprocess
import os
import glob
import shutil
import xlrd
from xlutils.copy import copy
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import unicodedata

# ---------------- Configurações ----------------
chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
sistema_url = "https://portal.exemplo.com.br/"
caminho_downloads = r"C:\Users\USUARIO\Downloads"
pasta_relatorios = r"C:\RelatoriosPDV"
os.makedirs(pasta_relatorios, exist_ok=True)

# ---------------- Identificadores por loja ----------------
IDENTIFICADORES = {
    "Loja A": "00.000.000/0001-00",
    "Loja B": "00.000.000/0002-00",
    "Loja C": "00.000.000/0003-00"
}

# ---------------- Datas ----------------
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

# ---------------- Funções auxiliares ----------------
def aguardar_download_result(timeout=40):
    inicio = time.time()
    while time.time() - inicio < timeout:
        arquivos = glob.glob(os.path.join(caminho_downloads, "Result*.xls"))
        if arquivos:
            arquivos.sort(key=os.path.getmtime, reverse=True)
            return arquivos[0]
        time.sleep(1)
    return None

def limpar_antigos_result():
    for fpath in glob.glob(os.path.join(caminho_downloads, "Result*.xls")):
        try:
            os.remove(fpath)
        except PermissionError:
            print(f"⚠️ Não foi possível remover {os.path.basename(fpath)} (arquivo em uso).")

def salvar_original(loja_nome, origem_path):
    data_str = DATA_FINAL.replace("/", "-")
    destino_path = os.path.join(pasta_relatorios, f"{loja_nome}_{data_str}.xls")
    if os.path.exists(destino_path):
        try:
            os.remove(destino_path)
        except PermissionError:
            print(f"⚠️ Arquivo {destino_path} está aberto. Feche-o e rode novamente.")
            return None
    shutil.copy2(origem_path, destino_path)
    print(f"✅ Exportado original salvo em: {destino_path}")
    return destino_path

def integrar_pdv(loja_nome, arquivo_exportado):
    with open(arquivo_exportado, "r", encoding="latin1") as f:
        html = f.read()

    soup = BeautifulSoup(html, "html.parser")
    table = soup.find("table")
    rows = table.find_all("tr")

    base_path = os.path.join(caminho_downloads, "PDV-Modelo.xls")
    rb = xlrd.open_workbook(base_path, formatting_info=True)
    wb = copy(rb)
    ws = wb.get_sheet(0)

    # Índices de colunas
    IDX_DATA = 1
    IDX_HIST = 2
    IDX_NOTA = 3
    IDX_VALOR = 6
    IDX_FORMA = 7
    IDX_CLIENTE = 9
    IDX_CARTAO = 10

    i = 2
    for tr in rows[1:]:
        cols = [td.get_text(strip=True) for td in tr.find_all("td")]
        if not cols or len(cols) < 11:
            continue

        historico = cols[IDX_HIST]
        forma_original = cols[IDX_FORMA].strip()

        if "REF. VENDA" not in historico.upper():
            continue
        if "TROCA DE MERCADORIA" in forma_original.upper():
            continue

        forma_norm = unicodedata.normalize("NFKD", forma_original).encode("ASCII", "ignore").decode("ASCII").strip().upper()
        if "TRANSFERENCIA/PIX" in forma_norm or "TRANFERENCIA/PIX" in forma_norm:
            forma_final = "PIX"
        else:
            forma_final = forma_original

        data = cols[IDX_DATA]
        valor = cols[IDX_VALOR]
        nota = cols[IDX_NOTA]
        cartao = cols[IDX_CARTAO]
        cliente = cols[IDX_CLIENTE]

        # Escrita nas colunas
        ws.write(i, 0, data)
        ws.write(i, 1, valor)
        ws.write(i, 2, forma_final)
        ws.write(i, 3, nota)
        ws.write(i, 4, cartao)
        ws.write(i, 5, data)
        ws.write(i, 7, IDENTIFICADORES[loja_nome])
        ws.write(i, 10, "Venda")
        ws.write(i, 11, cliente)
        ws.write(i, 12, historico)

        i += 1

    destino = os.path.join(pasta_relatorios, f"PDV_{loja_nome}_{DATA_FINAL.replace('/', '-')}.xls")
    wb.save(destino)
    print(f"✅ Integração concluída. Arquivo salvo em: {destino}")
    return destino

def validar_arquivo(arquivo_path):
    try:
        with open(arquivo_path, "r", encoding="latin1") as f:
            html = f.read()
        soup = BeautifulSoup(html, "html.parser")
        table = soup.find("table")
        if not table:
            return False
        rows = table.find_all("tr")
        return len(rows) > 1
    except Exception as e:
        print(f"⚠️ Erro ao validar arquivo: {e}")
        return False

# ---------------- Processamento por loja ----------------
def processar_loja(loja, max_tentativas=3):
    print(f"\n🔄 Processando loja: {loja['nome']}")

    tentativas = 0
    while tentativas < max_tentativas:
        tentativas += 1
        print(f"🌀 Tentativa {tentativas}...")

        subprocess.Popen([chrome_path, sistema_url])
        time.sleep(10)

        # Login automatizado
        pyautogui.click(x=697, y=423)
        pyautogui.write(loja["login"])
        pyautogui.click(x=473, y=490)
        pyautogui.write(loja["senha"])
        pyautogui.click(x=610, y=549)
        time.sleep(1)

        # Navegação e exportação
        pyautogui.press(loja["tecla_inicial"])
        for _ in range(loja["setinhas"]):
            pyautogui.press("down"); time.sleep(0.2)
        pyautogui.press("enter")
        time.sleep(1)

        limpar_antigos_result()
        pyautogui.click(x=1052, y=482)
        time.sleep(2)

        arquivo_download = aguardar_download_result(timeout=20)

        if arquivo_download and validar_arquivo(arquivo_download):
            print("✅ Arquivo válido baixado.")
            arquivo_exportado = salvar_original(loja['nome'], arquivo_download)
            if arquivo_exportado:
                return integrar_pdv(loja['nome'], arquivo_exportado)
        else:
            print("⚠️ Arquivo vazio ou não encontrado. Reiniciando tentativa...")

    print(f"❌ Falha após {max_tentativas} tentativas para loja {loja['nome']}.")
    return None

# ---------------- Lista de lojas ----------------
lojas = [
    {"nome": "Loja A", "login": "usuarioA", "senha": "senhaA", "setinhas": 9, "tecla_inicial": "m", "controle_caixa": (324, 504)},
    {"nome": "Loja B", "login": "usuarioB", "senha": "senhaB", "setinhas": 8, "tecla_inicial": "m", "controle_caixa": (332, 536)},
    {"nome": "Loja C", "login": "usuarioC", "senha": "senhaC", "setinhas": 7, "tecla_inicial": "m", "controle_caixa": (361, 535)}
]

# ---------------- Upload final ----------------
def upload_sistema(lista_arquivos):
    subprocess.Popen([chrome_path, "https://financas.exemplo.com.br/"])

# ---------------- Execução principal ----------------
if __name__ == "__main__":
    arquivos_gerados = []
    for loja_cfg in lojas:
        copia_planilha = processar_loja(loja_cfg)
        if copia_planilha:
            arquivos_gerados.append(copia_planilha)

    if len(arquivos_gerados) == len(lojas):
        upload_sistema(arquivos_gerados)
    else:
        print("⚠️ Atenção: Nem todas as lojas geraram PDV. Upload cancelado para evitar erro.")
