from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd

# Imports para atualizar o arquivo .xlsx no Google Drive
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# ================================
# PASSOS 1 a 13: Extração dos dados e salvamento local
# ================================

# Inicia o Chrome e maximiza a janela
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.maximize_window()

# 1) Acessa a URL
driver.get("https://portaltecon.csn.com.br/Home/Index")
time.sleep(2)

# 2) Aceita cookies
driver.find_element(By.XPATH, '//*[@id="cookieConsentContainer"]/table/tbody/tr/th[1]/div/a').click()
time.sleep(1)

# 3) Preenche usuário
driver.find_element(By.XPATH, '//*[@id="login"]').send_keys("T002961")
time.sleep(1)

# 4) Preenche senha
driver.find_element(By.XPATH, '//*[@id="password"]').send_keys("5e?C{8Go")
time.sleep(1)

# 5) Clica em "Entrar"
driver.find_element(By.XPATH, '//*[@id="divX"]/div/button[1]').click()
time.sleep(3)

# 6) Clica em "Agendamento"
driver.find_element(By.XPATH, '//*[@id="Agendamento"]').click()
time.sleep(2)

# 7) Clica em "Retirada"
driver.find_element(By.XPATH, '//*[@id="Retirada"]').click()
time.sleep(2)

# 8) Clica em "Agendar Cheio"
driver.find_element(By.XPATH, '//*[@id="AgendarRetirada"]').click()
time.sleep(2)

# 9) Preenche o campo "Número do Container"
numero_container = driver.find_element(By.XPATH, '//*[@id="NumeroContainer"]')
numero_container.click()
numero_container.clear()
numero_container.send_keys("TCKU2258525")
time.sleep(1)

# 10) Clica no botão de Buscar
driver.find_element(By.XPATH, '//*[@id="main"]/div[2]/div/button[1]').click()
time.sleep(5)

# 11) Clica no botão de Agendar/Editar e aguarda o modal carregar
driver.find_element(By.XPATH, '//*[@id="webGrid"]/tbody/tr/td[15]').click()

# Aguarda que a tabela do modal esteja visível
table = WebDriverWait(driver, 20).until(
    EC.visibility_of_element_located((By.XPATH, '//*[@id="TableDivModal"]/table'))
)

# ================================
# Passo 12: Extrair dados da tabela
# ================================

# Extrai os cabeçalhos da tabela (primeira linha)
headers_elements = driver.find_elements(By.XPATH, '//*[@id="TableDivModal"]/table/tbody/tr[1]/th')
headers = [header.text for header in headers_elements]
print("Cabeçalhos da tabela:")
for i, header in enumerate(headers, start=1):
    print(f"Coluna {i}: {header}")

# Extrai as linhas (a partir da segunda linha)
rows_elements = driver.find_elements(By.XPATH, '//*[@id="TableDivModal"]/table/tbody/tr[position()>1]')
table_data = []
for r_idx, row in enumerate(rows_elements, start=1):
    cells = row.find_elements(By.TAG_NAME, "td")
    row_data = []
    for c_idx, cell in enumerate(cells, start=1):
        if c_idx == 1:
            # Para a primeira coluna ("Hora/Dia"), extrai o texto diretamente
            value = cell.text.strip()
        else:
            # Para as demais colunas, extrai o status com base nas classes
            class_attr = cell.get_attribute("class")
            class_list = class_attr.split() if class_attr else []
            if "full" in class_list:
                value = "Indisponível"
            elif "open" in class_list and "selecionado" in class_list:
                value = "Selecionado"
            elif "open" in class_list:
                value = "Disponível"
            else:
                value = "Desconhecido"
        row_data.append(value)
    table_data.append(row_data)

print("\nTabela de horários (status extraído das classes):")
for row_index, row in enumerate(table_data, start=1):
    print(f"Linha {row_index}:", row)

# ================================
# Passo 13: Salvar os resultados em um arquivo Excel
# ================================

# Se os cabeçalhos baterem com o número de colunas, usa-os como nomes; caso contrário, cria DataFrame sem cabeçalhos
if headers and all(len(row) == len(headers) for row in table_data):
    df = pd.DataFrame(table_data, columns=headers)
else:
    df = pd.DataFrame(table_data)

output_path = r"C:\Users\leonardo.fragoso\Desktop\Projetos\TosInterno\downloads\tecon.xlsx"
df.to_excel(output_path, index=False)
print(f"\nDados salvos localmente em: {output_path}")

# Fecha o navegador (a partir daqui não precisamos mais do Selenium)
driver.quit()

# ================================
# Passo 14: Atualizar o arquivo .xlsx no Google Drive
# ================================

# ID do arquivo no Google Drive (parte da URL do arquivo)
file_id = "12oV8rgR9BisF_F-1Yx-ACt80cBggOJtk"

# Cria as credenciais a partir do arquivo JSON e define o escopo para o Drive
SCOPES = ['https://www.googleapis.com/auth/drive']
credentials = service_account.Credentials.from_service_account_file('gdrive_credentials.json', scopes=SCOPES)

# Constrói o serviço da API do Drive
drive_service = build('drive', 'v3', credentials=credentials)

# Prepara o upload do arquivo .xlsx
media = MediaFileUpload(output_path,
                        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        resumable=True)

# Atualiza o arquivo no Google Drive (sobrescrevendo o arquivo existente)
updated_file = drive_service.files().update(
    fileId=file_id,
    media_body=media
).execute()

print("\nArquivo .xlsx atualizado no Google Drive com sucesso!")
