import openpyxl
from PIL import Image, ImageDraw, ImageFont

def adicionar_legenda(imagem_path, legenda_texto, output_path, fonte_path, fonte_tamanho, posicao, cor):
    try:
        # Carregar a imagem
        imagem = Image.open(imagem_path)

        # Converter para o modo 'RGB' se estiver em 'RGBA'
        if imagem.mode == 'RGBA':
            imagem = imagem.convert('RGB')

        # Inicializar o objeto ImageDraw
        desenhador = ImageDraw.Draw(imagem)

        # Definir a fonte e o tamanho
        fonte = ImageFont.truetype(fonte_path, fonte_tamanho)

        # Adicionar a legenda à imagem
        desenhador.text(posicao, legenda_texto, font=fonte, fill=cor)

        # Salvar a imagem resultante
        imagem.save(output_path)
        print(f"Legenda adicionada com sucesso: {output_path}")

    except Exception as e:
        print(f"Erro ao adicionar legenda: {e}")

if __name__ == "__main__":
    # Configurações
    fonte_path = "teste/tahoma.ttf"
    fonte_tamanho = 15
    posicao = (130, 60)
    cor = "blue"

    # Carregar a planilha dos dados fictícios
    workbook = openpyxl.load_workbook('projeto5/dados_aleatorios.xlsx')
    worksheet = workbook['Sheet']

    # Iterar sobre as linhas da planilha
    for indice, linha in enumerate(worksheet.iter_rows(min_row=2,max_row=2)):
        nome, cargo, email, telefone = linha[:4]

        legenda_texto = f"{nome.value}\n {cargo.value}\n {email.value}\n {telefone.value}"

        # Definir caminho da imagem de entrada e saída
        imagem_path = "projeto5/imagem_base.png"
        output_path = f"projeto5/resultado_{indice + 1}.jpg"

        # Adicionar legenda à imagem
        adicionar_legenda(imagem_path, legenda_texto, output_path, fonte_path, fonte_tamanho, posicao, cor)
