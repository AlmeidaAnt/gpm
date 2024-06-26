import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import datetime
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pyautogui
from selenium.webdriver.common.alert import Alert

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
senha.send_keys('Endicon06.')
time.sleep(1)

# 1.3 Inserir credenciais de acesso - LOGAR
botao_login = navegador.find_element(By.XPATH,'//*[@id="form_login"]/input[5]')
botao_login.click()
time.sleep(5)

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
    emissao = coluna[1]
    valor = coluna[2]
    exercicio = coluna[3]
    nome_doc = coluna[4]
    nome_arq = coluna[6]

#placa
    b_placa = navegador.find_element(By.XPATH,'//*[@id="veiculos_chosen"]/a/span')
    b_placa.click()
    placa = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="veiculos_chosen"]/div/div/input')))
    placa.send_keys(veiculo)
    placa.send_keys(Keys.TAB)
    time.sleep(1)
#tipo Documento
    b_tipo = navegador.find_element(By.XPATH,'//*[@id="tip_doc_chosen"]/a/span')
    b_tipo.click()
    time.sleep(1)
    tipo = navegador.find_element(By.XPATH,'//*[@id="tip_doc_chosen"]/div/div/input')
    tipo.send_keys('CRLV')
    tipo.send_keys(Keys.TAB)
    time.sleep(10)
#Data Emissão
    data = navegador.find_element(By.XPATH,'//*[@id="dat_doc"]')
    data.clear()
    data.send_keys(emissao)
    time.sleep(5)
#Valor Total
    total = navegador.find_element(By.XPATH,'//*[@id="valor_doc"]')
    total.send_keys(valor)
    time.sleep(5)
#Exercicio
    licencimento = navegador.find_element(By.XPATH,'//*[@id="ano_exe"]')
    licencimento.clear()
    licencimento.send_keys(exercicio)
    time.sleep(5)
#Anexo
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('Enter')
    time.sleep(5)
    pyautogui.typewrite(nome_arq)
    time.sleep(5)
    pyautogui.press('Enter')
    time.sleep(2)
#Salvar
    salvar = navegador.find_element(By.XPATH,'//*[@id="idSubmit"]')
    salvar.click()
    alert = Alert(navegador)
    alert.accept()
    print("Documento cadastrado com Sucesso: ",placa)
    time.sleep(1)
    cadastro = navegador.find_element(By.XPATH, '//*[@id="idBtIncluir"]')
    cadastro.click()
    time.sleep(2)

