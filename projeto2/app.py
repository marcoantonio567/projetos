import openpyxl
from PIL import ImageDraw, Image, ImageFont
import os

# Carregar planilha
workbook_alunos = openpyxl.load_workbook('projeto2/planilha_alunos.xlsx')

# Acessar planilha
sheet_alunos = workbook_alunos['Sheet1']

# Definir diretório de salvamento
diretorio_certificados = 'projeto2/certificados'

# Iterar linhas da planilha (ignorando primeira linha)
for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2)):
    # Extrair dados da linha
    nome_curso = linha[0].value
    Nome_Participante = linha[1].value
    Tipo_Participação = linha[2].value
    Data_Início = linha[3].value
    Data_Término = linha[4].value
    Carga_Horária = linha[5].value
    Data_Emissão = linha[6].value

    # Carregar fontes
    font_nome = ImageFont.truetype('tahomabd.ttf', 90)
    font_geral = ImageFont.truetype('tahoma.ttf', 80)
    font_data = ImageFont.truetype('tahoma.ttf', 55)

    # Abrir imagem base
    image = Image.open('projeto2/certificado_padrao.jpg')

    # Criar objeto de desenho
    desenhar = ImageDraw.Draw(image)

    # Inserir texto na imagem
    desenhar.text((1020, 827), Nome_Participante, fill='black', font=font_nome)
    desenhar.text((1060, 950), nome_curso, fill='black', font=font_geral)
    desenhar.text((1435, 1065), Tipo_Participação, fill='black', font=font_geral)
    desenhar.text((1480, 1182), str(Carga_Horária), fill='black', font=font_geral)
    desenhar.text((750, 1770), Data_Início, fill='black', font=font_data)
    desenhar.text((750, 1930), Data_Término, fill='black', font=font_data)
    desenhar.text((2220, 1930), Data_Emissão, fill='black', font=font_data)

    # Criar o diretório de certificados se não existir
    os.makedirs(diretorio_certificados, exist_ok=True)

    # Salvar imagem no diretório de certificados
    image.save(f'{diretorio_certificados}/{indice},{Nome_Participante} certificado.png')

print("Certificados gerados com sucesso!")
