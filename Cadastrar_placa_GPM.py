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

#cadastro Veiculos
navegador.get('https://endiconpa.gpm.srv.br/cadastro/geral/veiculo.php')
workbook = openpyxl.load_workbook('C:\\temp\\cadastar_Placa_gpm.xlsx')
planilha = workbook.active

# 4. Criar iteração do cadastros
for coluna in planilha.iter_rows(min_row=2, values_only=True):
    veiculo = coluna[0]
    chassi = coluna[1]
    renavam = coluna[2]
    polo = coluna[3]
    ano_fab = coluna[4]
    ano_mod = coluna[5]
    descricao = coluna[6]
    produto = coluna[7]
    cc = coluna[8]

# Novo Registro
    novo_registro = navegador.find_element(By.XPATH, '//*[@id="idBtIncluir"]')
    novo_registro.click()
    time.sleep(2)
#placa
    placa = navegador.find_element(By.XPATH,'//*[@id="placa"]')
    placa.send_keys(veiculo)
    time.sleep(1)
#renavam
    renv = navegador.find_element(By.XPATH,'//*[@id="renavam"]')
    renv.send_keys(renavam)
    time.sleep(1)
#chassi
    n_chassi = navegador.find_element(By.XPATH,'//*[@id="chassi"]')
    n_chassi.send_keys(chassi)
    time.sleep(1)
#cidade
    bcidade = navegador.find_element(By.XPATH,'//*[@id="form"]/table[1]/tbody/tr[6]/td[2]/div/div[1]')
    bcidade.click()
    cidade = navegador.find_element(By.XPATH,'//*[@id="form"]/table[1]/tbody/tr[6]/td[2]/div/input')
    cidade.send_keys(polo)
    cidade.send_keys(Keys.ENTER)
    time.sleep(1)
#modelo
    bmodelo = navegador.find_element(By.XPATH,'//*[@id="form"]/table[1]/tbody/tr[8]/td[2]/div/div[1]')
    bmodelo.click()
    modelo = navegador.find_element(By.XPATH,'//*[@id="form"]/table[1]/tbody/tr[8]/td[2]/div/input')
    modelo.send_keys(descricao)
    modelo.send_keys(Keys.ENTER)
    time.sleep(1)
#combustivel
    #bcomb = navegador.find_element(By.XPATH,'//*[@id="form"]/table[1]/tbody/tr[9]/td[2]/div/div[1]')
    #bcomb.click()
    comb = navegador.find_element(By.XPATH,'//*[@id="form"]/table[1]/tbody/tr[9]/td[2]/div/input')
    comb.send_keys(produto)
    comb.send_keys(Keys.ENTER)
    time.sleep(1)
#status
    #bstts = navegador.find_element(By.XPATH,'//*[@id="form"]/table[1]/tbody/tr[10]/td[2]/div/div[1]')
    #bstts.click()
    stts = navegador.find_element(By.XPATH,'//*[@id="form"]/table[1]/tbody/tr[10]/td[2]/div/input')
    stts.send_keys('ATIVO')
    stts.send_keys(Keys.ENTER)
    time.sleep(1)
#cor
    #bcor = navegador.find_element(By.XPATH,'//*[@id="form"]/table[1]/tbody/tr[12]/td[2]/div/div[1]')
    #bcor.click()
    cor = navegador.find_element(By.XPATH,'//*[@id="form"]/table[1]/tbody/tr[12]/td[2]/div/input')
    cor.send_keys('BRANCO')
    cor.send_keys(Keys.ENTER)
    time.sleep(1)
#ano modelo
    a_mod = navegador.find_element(By.XPATH,'//*[@id="ano_mod"]')
    a_mod.send_keys(ano_mod)
    time.sleep(1)
#ano Fabricação
    a_fab = navegador.find_element(By.XPATH,'//*[@id="ano_fab"]')
    a_fab.send_keys(ano_fab)
    time.sleep(1)
#velocidade Maxima transito
    vel_tran = navegador.find_element(By.XPATH,'//*[@id="veloc_max"]')
    vel_tran.send_keys('8000')
    time.sleep(1)
#velocidade Maxima veiculo
    vel_veic = navegador.find_element(By.XPATH,'//*[@id="veloc_max_vei"]')
    vel_veic.send_keys('8000')
    time.sleep(1)
#centro de Custo
    bcc = navegador.find_element(By.XPATH,'//*[@id="form"]/table[1]/tbody/tr[26]/td[2]/div/div[1]')
    bcc.click()
    centro_de_custo = navegador.find_element(By.XPATH,'//*[@id="form"]/table[1]/tbody/tr[26]/td[2]/div/input')
    centro_de_custo.send_keys(cc)
    time.sleep(1)
#Cadastrar
    interceptar = navegador.find_element(By.XPATH,'//*[@id="form"]/table[2]/tbody/tr/td[1]')
    interceptar.click()
    cadastrar = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="form"]/table[2]/tbody/tr/td[1]/input')))
    cadastrar.click()

#placa cadastrada
    print(veiculo,' Cadastrado Com Sucesso!')
    time.sleep(2)

    navegador.get('https://endiconpa.gpm.srv.br/cadastro/geral/veiculo.php')
    time.sleep(2)