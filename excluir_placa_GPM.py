import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import datetime
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

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

#cadastro Veiculos
navegador.get('https://endiconpa.gpm.srv.br/cadastro/geral/veiculo.php')
workbook = openpyxl.load_workbook('C:\\temp\\excluir_Placa_gpm.xlsx')
planilha = workbook.active

# 4. Criar iteração do cadastros
for coluna in planilha.iter_rows(min_row=2, values_only=True):
    veiculo = coluna[0]
    time.sleep(2)

#pesquisar placas
    clk_placa = navegador.find_element(By.XPATH,'//*[@id="form"]/table/tbody/tr[2]/td[2]/div/div[1]')
    clk_placa.click()
    time.sleep(1)
    placa = navegador.find_element(By.XPATH,'//*[@id="form"]/table/tbody/tr[2]/td[2]/div/input')
    placa.send_keys(veiculo)
    time.sleep(1)
    placa.send_keys(Keys.ENTER)
    time.sleep(5)

    psc_placa = navegador.find_element(By.XPATH,'//*[@id="form"]/table/tbody/tr[11]/td/input')
    psc_placa.click()
    time.sleep(1)
#EXCLUIR
    excluir = navegador.find_element(By.XPATH,'//*[@id="tab_resultados"]/tbody/tr/td[1]/a[3]')
    excluir.click()
    time.sleep(2)
#confirmar
    confirmar = navegador.find_element(By.XPATH,'/html/body/div[2]/div/div[6]/button[1]')
    confirmar.click()
    print('Placa Inativada ',veiculo)
