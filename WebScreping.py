from selenium import webdriver 
from selenium.webdriver.common.keys import Keys
import time 
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
import xlrd

arq = open("Resultado.xls","w")
workbook = xlrd.open_workbook(r'C:\Users\thexe\OneDrive\Python Bot\Projeto_CVM\CNPJ_FIs.xls')
sheet = workbook.sheet_by_name('CNPJ_FI')
rows = sheet.nrows
columns = sheet.ncols

options = webdriver.ChromeOptions()
options.add_argument('--disable-logging')
options.add_argument('--log-level=3')

driver = webdriver.Chrome("C:/Users/thexe\OneDrive/Área de Trabalho/Xeco/Robo/chromedriver.exe", options= options)
driver.get ("https://web.cvm.gov.br/app/fundosweb/#/consultaPublica")

for curr_row in range(0,rows):
    x = sheet.cell_value(curr_row,0)
    time.sleep(10)
    #BuscaFundo = driver.find_element(By.LINK_TEXT,'https://sistemas.cvm.gov.br/port/fundos/consultafundos.asp')
    time.sleep(3)
    Pesquisa = driver.find_element(By.ID,"txtCnpj")
    time.sleep(3)
    Pesquisa.clear()
    time.sleep(4)
    Pesquisa.send_keys(x)
    time.sleep(5)
    #Pesquisa.send_keys(Keys.RETURN)
    time.sleep(3)
    click_pesquisa = driver.find_element(By.XPATH, '//*[@id="panelFiltros"]/div/div/a[1]/span' )
    click_pesquisa.click()
    time.sleep(3)
    visualizar_pesquisa = driver.find_element(By.XPATH,'//*[@id="containerView"]/div[2]/div[3]/table/tbody/tr/td[9]/a[1]')
    visualizar_pesquisa.click()
    time.sleep(3)
    click_participantes = driver.find_element(By.XPATH,'//*[@id="tabFundo"]/ul/li[2]')
    click_participantes.click()
    time.sleep(3)
    Dados_Distribuidor = driver.find_element(By.XPATH,'//*[@id="tabPanelVisualizarParticipante"]/div/div[2]/div[21]/div/fieldset/div[2]/div/div/table[1]/tbody/tr/td[3]')
    
    texto = 'O FI %s %s\n' %(x,Dados_Distribuidor.accessible_name)
    time.sleep(1)
    
    Boton_Return = driver.find_element(By.XPATH,'//*[@id="containerView"]/div[2]/div[2]/div[1]/a/span[2]')
    Boton_Return.click()
    arq.write(texto)
    time.sleep(2)
arq.close()    




