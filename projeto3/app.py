from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
import openpyxl


# configure firefox options
options = webdriver.FirefoxOptions()
# create firefox driver
driver = webdriver.Firefox(options=options)


#pyautogui
workbook = openpyxl.load_workbook(r'C:\Users\marco\Desktop\update_repositorio\projetos_rpa\projetos\projeto3\produtos.xlsx')
vendas_sheet = workbook['Produtos']


# define url 
url = "http://127.0.0.1:3002/projetos/projeto3/sistema.html?serverWindowId=3e2e5378-1382-42e8-8c28-62c762d39049" 

# access url
driver.get(url)


for row in vendas_sheet.iter_rows(min_row=2,max_row=101,values_only=True):  
    driver.find_element(By.XPATH,'//*[@id="nomeProduto"]').send_keys(row[0])
    driver.find_element(By.XPATH,'//*[@id="descricao"]').send_keys(row[1])
    driver.find_element(By.XPATH,'//*[@id="categoria"]').send_keys(row[2])
    driver.find_element(By.XPATH,'//*[@id="codigoProduto"]').send_keys(row[3])
    driver.find_element(By.XPATH,'//*[@id="peso"]').send_keys(row[4])
    driver.find_element(By.XPATH,'//*[@id="dimensoes"]').send_keys(row[5])
    driver.find_element(By.XPATH,'//*[@id="preco"]').send_keys(row[6])
    driver.find_element(By.XPATH,'//*[@id="quantidadeEstoque"]').send_keys(row[7])
    driver.find_element(By.XPATH,'//*[@id="dataValidade"]').send_keys(row[8])
    driver.find_element(By.XPATH,'//*[@id="cor"]').send_keys(row[9])
    driver.find_element(By.XPATH,'//*[@id="tamanho"]').send_keys(row[10])
    driver.find_element(By.XPATH,'//*[@id="material"]').send_keys(row[11])
    driver.find_element(By.XPATH,'//*[@id="fabricante"]').send_keys(row[12])
    driver.find_element(By.XPATH,'//*[@id="paisOrigem"]').send_keys(row[13])
    driver.find_element(By.XPATH,'//*[@id="observacoes"]').send_keys(row[14])
    driver.find_element(By.XPATH,'//*[@id="codigoBarras"]').send_keys(row[15])
    driver.find_element(By.XPATH,'//*[@id="localizacaoArmazem"]').send_keys(row[16])
    driver.find_element(By.XPATH,'//*[@id="formCadastro"]/input[16]').click()
