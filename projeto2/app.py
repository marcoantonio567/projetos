import openpyxl
from PIL import ImageDraw, Image, ImageFont
import os

# load workbook
workbook_alunos = openpyxl.load_workbook(r'C:\Users\marco\Desktop\update_repositorio\projetos_rpa\projetos\projeto2\planilha_alunos.xlsx')

# access worksheet
sheet_alunos = workbook_alunos['Sheet1']

# define directory for saving
diretorio_certificados = r'C:\Users\marco\Desktop\update_repositorio\projetos_rpa\projetos\projeto2\certificados'

# iterate rows of worksheet (ignorando primeira linha)
for indice, row in enumerate(sheet_alunos.iter_rows(min_row=2)):
    # extract data from row
    nome_curso = row[0].value
    Nome_Participante = row[1].value
    Tipo_Participação = row[2].value
    Data_Início = row[3].value
    Data_Término = row[4].value
    Carga_Horária = row[5].value
    Data_Emissão = row[6].value

    # load fonts
    font_nome = ImageFont.truetype(r'C:\Users\marco\Desktop\update_repositorio\projetos_rpa\projetos\projeto2\tahomabd.ttf', 90)
    font_geral = ImageFont.truetype(r'C:\Users\marco\Desktop\update_repositorio\projetos_rpa\projetos\projeto2\tahoma.ttf', 80)
    font_data = ImageFont.truetype(r'C:\Users\marco\Desktop\update_repositorio\projetos_rpa\projetos\projeto2\tahoma.ttf', 55)

    # open base image
    image = Image.open(r'C:\Users\marco\Desktop\update_repositorio\projetos_rpa\projetos\projeto2\certificado_padrao.jpg')

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

    # create directory for saving if it doesn't exist
    os.makedirs(diretorio_certificados, exist_ok=True)

    # save image to directory for saving
    image.save(os.path.join(diretorio_certificados, f'{indice},{Nome_Participante} certificado.png'))

print("Certificados gerados com sucesso!")
