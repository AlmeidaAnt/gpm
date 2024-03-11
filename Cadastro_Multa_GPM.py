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

#Planilha
workbook = openpyxl.load_workbook('C:\\temp\\cadastro_multa_GPM.xlsx')
planilha = workbook.active

# 4. Criar iteração do cadastros
for coluna in planilha.iter_rows(min_row=2, values_only=True):
    veiculo = coluna[0]
    centro_de_custo = coluna[1]
    infrator = coluna[2]
    infracao = coluna[3]
    dt_infracao = coluna[4]
    hr_infracao = coluna[5]
    polo = coluna[6]
    endereco = coluna[7]
    numero = coluna[8]
    orgao = coluna[10]
    dt_recebimento = coluna[11]
    prazo = coluna[12]
    vencimento = coluna[13]
    valor = coluna[14]
    status = coluna[15]

#cadastro Multas
    navegador.get('https://endiconpa.gpm.srv.br/cadastro/geral/cad_multa_veiculo.php')
    time.sleep(1)
#novo registro
    novo_registro = navegador.find_element(By.XPATH,'//*[@id="idBtIncluir"]')
    novo_registro.click()
    time.sleep(1)
#placa
    bplaca = navegador.find_element(By.XPATH,'//*[@id="cod_veic_chosen"]/a/div')
    bplaca.click()
    placa = navegador.find_element(By.XPATH,'//*[@id="cod_veic_chosen"]/div/div/input')
    placa.send_keys(veiculo)
    placa.send_keys(Keys.ENTER)
    time.sleep(1)
#centro de busto
    bcc = navegador.find_element(By.XPATH,'//*[@id="centro_custo_chosen"]/a/div')
    bcc.click()
    cc = navegador.find_element(By.XPATH,'//*[@id="centro_custo_chosen"]/div/div/input')
    cc.send_keys(centro_de_custo)
    cc.send_keys(Keys.ENTER)
    time.sleep(1)
#funcionario
    #func = navegador.find_element(By.XPATH,'//*[@id="inputString"]')
    #func.send_keys(infrator)
    #time.sleep(1)
#data infração
    data = navegador.find_element(By.XPATH,'//*[@id="dt_multa"]')
    data.send_keys(dt_infracao)
    data.send_keys(Keys.TAB)
    time.sleep(1)
#hora Infração
    hora = navegador.find_element(By.XPATH,'//*[@id="hr_multa"]')
    hora.send_keys(hr_infracao)
    hora.send_keys(Keys.TAB)
    time.sleep(1)
#infração
    bmulta = navegador.find_element(By.XPATH,'//*[@id="cod_infracao_chosen"]/a/div')
    bmulta.click()
    multa = navegador.find_element(By.XPATH,'//*[@id="cod_infracao_chosen"]/div/div/input')
    multa.send_keys(infracao)
    multa.send_keys(Keys.ENTER)
    time.sleep(1)
#cidade
    bcidade = navegador.find_element(By.XPATH,'//*[@id="cod_municipio_chosen"]/a/div')
    bcidade.click()
    cidade = navegador.find_element(By.XPATH,'//*[@id="cod_municipio_chosen"]/div/div/input')
    cidade.send_keys(polo)
    cidade.send_keys(Keys.ENTER)
    time.sleep(1)
#local
    local = navegador.find_element(By.XPATH,'//*[@id="local"]')
    local.send_keys(endereco)
    time.sleep(1)
#numero Auto Infração
    ait = navegador.find_element(By.XPATH,'//*[@id="auto"]')
    ait.send_keys(numero)
    time.sleep(1)
#base legal
    base = navegador.find_element(By.XPATH,'//*[@id="base"]')
    base.send_keys(polo)
    time.sleep(1)
#orgão
    borg = navegador.find_element(By.XPATH,'//*[@id="cod_orgao_atuante_chosen"]/a/div')
    borg.click()
    org = navegador.find_element(By.XPATH,'//*[@id="cod_orgao_atuante_chosen"]/div/div/input')
    org.send_keys(orgao)
    org.send_keys(Keys.ENTER)
    time.sleep(1)
#data recebimento
    data_recebimento = navegador.find_element(By.XPATH,'//*[@id="dt_receb"]')
    data_recebimento.send_keys(dt_recebimento)
    time.sleep(1)
#prazo para identificar o condutor
    ident = navegador.find_element(By.XPATH,'//*[@id="dt_recorre"]')
    ident.send_keys(prazo)
    time.sleep(1)
#vencimento
    venc = navegador.find_element(By.XPATH,'//*[@id="dt_venc"]')
    venc.send_keys(vencimento)
    venc.send_keys(Keys.TAB)
    time.sleep(1)
#valor
    val = navegador.find_element(By.XPATH,'//*[@id="valor"]')
    val.send_keys(valor)
    time.sleep(1)
#status Pagamento
    bstt = navegador.find_element(By.XPATH,'//*[@id="status_chosen"]/a/div')
    bstt.click()
    stt = navegador.find_element(By.XPATH,'//*[@id="status_chosen"]/div/div/input')
    stt.send_keys(status)
    stt.send_keys(Keys.ENTER)
#salvar
    salvar = navegador.find_element(By.XPATH,'/html/body/form[3]/div/input[1]')
    salvar.click()
    time.sleep(5)
    print(veiculo, 'multa:', numero,' Cadastrada com Sucesso!')
