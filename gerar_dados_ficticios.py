import openpyxl
from faker import Faker
import phonenumbers

def gerar_telefone_brasileiro():
    fake = Faker()
    ddd_brasileiro = "+55"  # Código do Brasil

    # Lista de códigos de área (DDD) do Brasil
    lista_ddd = ["11", "12", "13", "14", "15", "16", "17", "18", "19", "21", "22", "24", "27", "28", "31", "32", "33", "34", "35", "37", "38", "41", "42", "43", "44", "45", "46", "47", "48", "49", "51", "53", "54", "55", "61", "62", "63", "64", "65", "66", "67", "68", "69", "71", "73", "74", "75", "77", "79", "81", "82", "83", "84", "85", "86", "87", "88", "89", "91", "92", "93", "94", "95", "96", "97", "98", "99"]

    # Escolha aleatória de um código de área
    ddd_escolhido = fake.random_element(lista_ddd)

    # Gera um número de telefone com o DDD escolhido
    numero_telefone = fake.numerify(text=f'{ddd_brasileiro} {ddd_escolhido} #########')

    while not phonenumbers.is_valid_number(phonenumbers.parse(numero_telefone, "BR")):
        ddd_escolhido = fake.random_element(lista_ddd)
        numero_telefone = fake.numerify(text=f'{ddd_brasileiro} {ddd_escolhido} #########')

    return numero_telefone

def gerar_dados_fake(quantidade):
    fake = Faker()
    dados = []

    for _ in range(quantidade):
        nome = fake.name()
        cargo = fake.job()
        email = fake.email()
        telefone = gerar_telefone_brasileiro()
        dados.append((nome, cargo, email, telefone))

    return dados

def salvar_em_excel(dados, nome_arquivo):
    planilha = openpyxl.Workbook()
    planilha_ativa = planilha.active
    planilha_ativa.append(["Nome", "Cargo", "Email", "Telefone"])

    for dado in dados:
        planilha_ativa.append(dado)

    planilha.save(nome_arquivo)
    print(f'Dados salvos com sucesso em "{nome_arquivo}".')

if __name__ == "__main__":
    quantidade_de_dados = 10  # Altere conforme necessário
    nome_do_arquivo = "dados_aleatorios.xlsx"

    dados_gerados = gerar_dados_fake(quantidade_de_dados)
    salvar_em_excel(dados_gerados, nome_do_arquivo)
