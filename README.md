# ⚙️ Automação de Integração PDV
Este projeto é um script em Python que automatiza o processo de login em um sistema web, exportação de relatórios em formato .xls, tratamento dos dados e integração com uma planilha modelo de PDV. Ao final, os arquivos gerados podem ser enviados para outro sistema financeiro.

# 🚀 Funcionalidades
Login automático no sistema via navegador.

Exportação de relatórios em formato .xls.

Validação dos arquivos baixados para garantir que não estejam vazios.

Tratamento dos dados:

Normalização de formas de pagamento (ex.: PIX).

Exclusão de registros não relevantes.

Inserção de dados em planilha modelo (PDV-Modelo.xls).

Integração por loja com identificadores únicos.

Upload automático para sistema financeiro (simulado).

# 📂 Estrutura do Projeto
main.py → Script principal com toda a lógica.

PDV-Modelo.xls → Planilha base utilizada para integração.

Pasta de Downloads → Local onde os relatórios são baixados.

Pasta de RelatoriosPDV → Local onde os arquivos tratados são salvos.

# ⚙️ Configurações
No início do código existem variáveis que podem ser ajustadas conforme seu ambiente:

sistema_url → URL da página de login do sistema.

chrome_path → Caminho do executável do navegador Chrome.

caminho_downloads → Diretório padrão de downloads do navegador.

pasta_relatorios → Diretório onde os relatórios tratados serão salvos.

IDENTIFICADORES → Dicionário com identificadores únicos (ex.: CNPJs fictícios) por loja.

# 📑 Estrutura da Lista de Lojas
As lojas são configuradas diretamente no código, em uma lista de dicionários:

python
lojas = [
    {"nome": "Loja A", "login": "usuarioA", "senha": "senhaA", "setinhas": 9, "tecla_inicial": "m", "controle_caixa": (324, 504)},
    {"nome": "Loja B", "login": "usuarioB", "senha": "senhaB", "setinhas": 8, "tecla_inicial": "m", "controle_caixa": (332, 536)},
    {"nome": "Loja C", "login": "usuarioC", "senha": "senhaC", "setinhas": 7, "tecla_inicial": "m", "controle_caixa": (361, 535)}
]
Cada loja possui:

Nome

Login e senha

Tecla inicial para navegação

Quantidade de "setinhas" (pressões de tecla para navegar)

Coordenadas de clique para acessar o controle de caixa

# 🖥️ Dependências
O script utiliza as seguintes bibliotecas:

os, time, glob, shutil, subprocess, datetime, unicodedata → Bibliotecas padrão do Python.

pyautogui → Automação de cliques e teclado.

xlrd, xlutils → Manipulação de arquivos Excel antigos (.xls).

BeautifulSoup (bs4) → Leitura e parsing de HTML.

Instale as dependências com:

bash
pip install pyautogui xlrd xlutils bs4

# ▶️ Como Executar
Ajuste as variáveis de configuração (sistema_url, chrome_path, etc.).

Configure a lista de lojas com login, senha e coordenadas.

Certifique-se de que o arquivo PDV-Modelo.xls está disponível na pasta de downloads.

Execute o script:

bash
python main.py
O programa irá:

Abrir o navegador e fazer login em cada loja.

Exportar relatórios.

Validar e tratar os arquivos.

Gerar planilhas integradas por loja.

(Opcional) Fazer upload para o sistema financeiro.

# ⚠️ Observações
As coordenadas de tela foram configuradas para uma resolução específica. Se sua tela tiver outra resolução, ajuste os valores.

O envio para o sistema financeiro é apenas simulado (upload_sistema), você pode adaptar para o sistema real.

Este código é um template genérico. Substitua os valores fictícios (URLs, identificadores, credenciais) pelos dados reais do seu ambiente.
