from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
import openpyxl


# configure firefox options
options = webdriver.FirefoxOptions()
# create firefox driver
driver = webdriver.Firefox(options=options)


# access site https://www.novaliderinformatica.com.br/computadores-gamers
url = "https://www.novaliderinformatica.com.br/computadores-gamers"

# Acessando a URL
driver.get(url)


# extrair todos os títulos
titulos = driver.find_elements(By.XPATH,"//a[@class='nome-produto']")
# extrair todos os preços
precos = driver.find_elements(By.XPATH,"//strong[@class='preco-promocional']")

# create workbook
workbook = openpyxl.Workbook()
# create sheet 'produtos'
workbook.create_sheet('produtos')
# select sheet 'produtos'
sheet_produtos = workbook['produtos']
sheet_produtos['A1'].value = 'Produto'
sheet_produtos['B1'].value = 'Preço'

# insert titles and prices in workbook sheet 'produtos'
# select sheet 'produtos'
sheet_produtos = workbook['produtos']
# insert titles and prices in workbook sheet 'produtos'
# select sheet 'produtos'
sheet_produtos = workbook['produtos']
# insert titles and prices in workbook sheet 'produtos'
for titulo, preco in zip(titulos, precos):
    sheet_produtos.append([titulo.text,preco.text])
# sheet_produtos.append([titulo.text,preco.text])

workbook.save('projeto4/produtos.xlsx') 
driver.close()