from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
import openpyxl


# Configurando as opções do Firefox
options = webdriver.FirefoxOptions()
# Criando o driver do Firefox
driver = webdriver.Firefox(options=options)



# acessar o site https://www.novaliderinformatica.com.br/computadores-gamers
url = "https://www.novaliderinformatica.com.br/computadores-gamers"

# Acessando a URL
driver.get(url)


# extrair todos os títulos
titulos = driver.find_elements(By.XPATH,"//a[@class='nome-produto']")
# extrair todos os preços
precos = driver.find_elements(By.XPATH,"//strong[@class='preco-promocional']")

# Criando a planilha
workbook = openpyxl.Workbook()
# Criando a página 'produtos'
workbook.create_sheet('produtos')
# Seleciono a página produtos
sheet_produtos = workbook['produtos']
sheet_produtos['A1'].value = 'Produto'
sheet_produtos['B1'].value = 'Preço'

# inserir os títulos e preços na planilha
for titulo, preco in zip(titulos, precos):
    sheet_produtos.append([titulo.text,preco.text])

workbook.save('projeto4/produtos.xlsx') 
driver.close()