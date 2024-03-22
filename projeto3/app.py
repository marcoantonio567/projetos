from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
import openpyxl


# Configurando as opções do Firefox
options = webdriver.FirefoxOptions()
# Criando o driver do Firefox
driver = webdriver.Firefox(options=options)


#pyautogui
workbook = openpyxl.load_workbook('projeto3/produtos.xlsx')
vendas_sheet = workbook['Produtos']


# Definindo a URL 
url = "http://127.0.0.1:3000/projeto3/sistema.html" # caso esse de errado troque pela url do 'sistema'

# Acessando a URL
driver.get(url)
for linha in vendas_sheet.iter_rows(min_row=2,max_row=101,values_only=True):
    driver.find_element(By.XPATH,'//*[@id="nomeProduto"]').send_keys(linha[0])
    driver.find_element(By.XPATH,'//*[@id="descricao"]').send_keys(linha[1])
    driver.find_element(By.XPATH,'//*[@id="categoria"]').send_keys(linha[2])
    driver.find_element(By.XPATH,'//*[@id="codigoProduto"]').send_keys(linha[3])
    driver.find_element(By.XPATH,'//*[@id="peso"]').send_keys(linha[4])
    driver.find_element(By.XPATH,'//*[@id="dimensoes"]').send_keys(linha[5])
    driver.find_element(By.XPATH,'//*[@id="preco"]').send_keys(linha[6])
    driver.find_element(By.XPATH,'//*[@id="quantidadeEstoque"]').send_keys(linha[7])
    driver.find_element(By.XPATH,'//*[@id="dataValidade"]').send_keys(linha[8])
    driver.find_element(By.XPATH,'//*[@id="cor"]').send_keys(linha[9])
    driver.find_element(By.XPATH,'//*[@id="tamanho"]').send_keys(linha[10])
    driver.find_element(By.XPATH,'//*[@id="material"]').send_keys(linha[11])
    driver.find_element(By.XPATH,'//*[@id="fabricante"]').send_keys(linha[12])
    driver.find_element(By.XPATH,'//*[@id="paisOrigem"]').send_keys(linha[13])
    driver.find_element(By.XPATH,'//*[@id="observacoes"]').send_keys(linha[14])
    driver.find_element(By.XPATH,'//*[@id="codigoBarras"]').send_keys(linha[15])
    driver.find_element(By.XPATH,'//*[@id="localizacaoArmazem"]').send_keys(linha[16])
    driver.find_element(By.XPATH,'//*[@id="formCadastro"]/input[16]').click()
