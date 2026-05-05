import openpyxl
from faker import Faker
import phonenumbers

def generate_brazilain_phone():
    fake = Faker()
    ddd_brasileiro = "+55"  # Brazil area code

    # Brazil area codes list
    lista_ddd = ["11", "12", "13", "14", "15", "16", "17", "18", "19", "21", "22", "24", "27", "28", "31", "32", "33", "34", "35", "37", "38", "41", "42", "43", "44", "45", "46", "47", "48", "49", "51", "53", "54", "55", "61", "62", "63", "64", "65", "66", "67", "68", "69", "71", "73", "74", "75", "77", "79", "81", "82", "83", "84", "85", "86", "87", "88", "89", "91", "92", "93", "94", "95", "96", "97", "98", "99"]

    # Random area code selection
    ddd_selected = fake.random_element(lista_ddd)

    # Generate phone number with selected area code
    phone_number = fake.numerify(text=f'{ddd_brasileiro} {ddd_selected} #########')

    while not phonenumbers.is_valid_number(phonenumbers.parse(phone_number, "BR")):
        ddd_selected = fake.random_element(lista_ddd)
        phone_number = fake.numerify(text=f'{ddd_brasileiro} {ddd_selected} #########')

    return phone_number

def generate_random_datas(quantity):
    fake = Faker()
    datas = []

    for _ in range(quantity):
        name = fake.name()
        job = fake.job()
        email = fake.email()
        phone = generate_brazilain_phone()
        datas.append((name, job, email, phone))

    return datas

def save_to_excel(datas, file_name):
    planilha = openpyxl.Workbook()
    planilha_ativa = planilha.active
    planilha_ativa.append(["Nome", "Cargo", "Email", "Telefone"])

    for dado in datas:
        planilha_ativa.append(dado)

    planilha.save(file_name)
    print(f'datas salvos com sucesso em "{file_name}".')

if __name__ == "__main__":
    quantidade_de_datas = 10  # Altere conforme necessário
    file_name = "projeto5/random_datas.xlsx"

    datas_gerados = generate_random_datas(quantidade_de_datas)
    save_to_excel(datas_gerados, file_name)
