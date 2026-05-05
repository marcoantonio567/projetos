import openpyxl
from PIL import Image, ImageDraw, ImageFont

def add_sub_title(imagem_path, sub_title_text, output_path, source_path, font_size, position, color):
    try:
        # load image
        image = Image.open(imagem_path)

        # convert to RGB if it's RGBA mode
        if image.mode == 'RGBA':
            image = image.convert('RGB')

        # initialize ImageDraw object
        draw = ImageDraw.Draw(image)

        # define font and size
        font = ImageFont.truetype(source_path, font_size)

        # Add the legend to the image
        draw.text(position, sub_title_text, font=font, fill=color)

        # save the image with the legend
        image.save(output_path)
        print(f"Legenda adicionada com sucesso: {output_path}")

    except Exception as e:
        print(f"Erro ao adicionar legenda: {e}")

if __name__ == "__main__":
    # configurations
    source_path = "teste/tahoma.ttf"
    font_size = 15
    position = (130, 60)
    color = "blue"
    
    # load the workbook
    workbook = openpyxl.load_workbook('projeto5/random_datas.xlsx')
    worksheet = workbook['Sheet']

    # iterate over the rows of the worksheet
    for indice, linha in enumerate(worksheet.iter_rows(min_row=2)):
        name, job, email, phone = linha[:4]

        sub_title_text = f"{name.value}\n {job.value}\n {email.value}\n {phone.value}"

        # Definir caminho da imagem de entrada e saída
        imagem_path = "projeto5/base_image.png"
        output_path = f"projeto5/result/resultado_{indice + 1}.jpg"

        # Adicionar legenda à imagem
        add_sub_title(imagem_path, sub_title_text, output_path, source_path, font_size, position, color)
