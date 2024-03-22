import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import datetime
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# *Configurações de Data
hoje = datetime.date.today()
hoje_formatado = hoje.strftime('%d/%m/%Y')

# 1. Abrir o GPM
navegador = webdriver.Chrome()
navegador.get('https://endiconpa.gpm.srv.br/index.php')
navegador.maximize_window()

# 1.1 Inserir credenciais de acesso - LOGIN
usuario = navegador.find_element(By.XPATH,'//*[@id="idLogin"]')
usuario.send_keys('ANTONIO.VINICIUS')
time.sleep(1)

# 1.2 Inserir credenciais de acesso - SENHA
senha = navegador.find_element(By.XPATH,'//*[@id="idSenha"]')
senha.send_keys('Endicon2024.')
time.sleep(1)

# 1.3 Inserir credenciais de acesso - LOGAR
botao_login = navegador.find_element(By.XPATH,'//*[@id="form_login"]/input[5]')
botao_login.click()
time.sleep(1)

#Cadastro de Documento
navegador.get('https://endiconpa.gpm.srv.br/cadastro/geral/documentos_veiculo.php')

#Novo Registro
cadastro = navegador.find_element(By.XPATH,'//*[@id="idBtIncluir"]')
cadastro.click()
time.sleep(2)

#Planilha
workbook = openpyxl.load_workbook('C:\\temp\\crlv_gpm_cadastro.xlsx')
planilha = workbook.active

#  Criar iteração do cadastros
for coluna in planilha.iter_rows(min_row=2, values_only=True):
    veiculo = coluna[0]

#placa
    b_placa = navegador.find_element(By.XPATH,'<span>Selecione o veículo </span>')
    b_placa.click()
    placa = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="veiculos_chosen"]/div/div/input')))
    placa.send_keys(veiculo)
    time.sleep(1)
#tipo Documento
    b_tipo = navegador.find_element(By.XPATH,'//*[@id="tip_doc_chosen"]/a/span')
    b_tipo.click()
    time.sleep(1)
    tipo = navegador.find_element(By.XPATH,'//*[@id="tip_doc_chosen"]/div/div/input')
    tipo.send_keys('CRLV')
    tipo.send_keys(Keys.TAB)
    time.sleep(10)