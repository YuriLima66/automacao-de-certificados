import openpyxl
from PIL import Image, ImageDraw, ImageFont
import datetime

workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')  # Carregar a planilha
sheet_alunos = workbook_alunos['Sheet1']  # Selecionar a aba da planilha

# Iterar sobre as linhas da planilha
for indice, linha in enumerate(sheet_alunos.iter_rows(max_row=2)):
    # Extrair os dados de cada célula
    nome_curso = linha[0].value
    nome_participante = linha[1].value
    tipo_participacao = linha[2].value
    data_inicio = linha[3].value
    data_final = linha[4].value
    carga_horaria = linha[5].value
    data_emissao = linha[6].value

    # Verificar se as datas são do tipo datetime e formatá-las
    if isinstance(data_inicio, datetime.datetime):
        data_inicio = data_inicio.strftime('%d/%m/%Y')
    if isinstance(data_final, datetime.datetime):
        data_final = data_final.strftime('%d/%m/%Y')
    if isinstance(data_emissao, datetime.datetime):
        data_emissao = data_emissao.strftime('%d/%m/%Y')

    # Definir as fontes a serem usadas
    fonte_nome = ImageFont.truetype('./BERNHC.ttf', 40)
    fonte_geral = ImageFont.truetype('./BERNHC.ttf', 40)
    fonte_data = ImageFont.truetype('./BERNHC.ttf', 40)

    # Abrir a imagem do certificado
    image = Image.open('certificado_padrao.jpeg')
    desenhar = ImageDraw.Draw(image)

    # Adicionar os dados do participante ao certificado
    desenhar.text((480, 380), nome_participante, fill='black', font=fonte_nome)
    desenhar.text((480, 430), nome_curso, fill='black', font=fonte_geral)
    desenhar.text((600, 490), tipo_participacao, fill='black', font=fonte_geral)
    desenhar.text((620, 545), str(carga_horaria), fill='black', font=fonte_geral)
    desenhar.text((260, 825), str(data_inicio), fill='blue', font=fonte_data)
    desenhar.text((260, 900), str(data_final), fill='blue', font=fonte_data)
    desenhar.text((480, 1010), str(data_emissao), fill='blue', font=fonte_data)

    # Salvar o certificado com um nome único
    image.save(f'./{indice}_{nome_participante}_certificado.png')