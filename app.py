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
    fonte_nome = ImageFont.truetype('./BOD_PSTC.ttf', 90)
    fonte_geral = ImageFont.truetype('./BOD_PSTC.ttf', 80)
    fonte_data = ImageFont.truetype('./BOD_PSTC.ttf', 55)

    # Abrir a imagem do certificado
    image = Image.open('./certificado_padrao.jpg')
    desenhar = ImageDraw.Draw(image)

    # Adicionar os dados do participante ao certificado
    desenhar.text((400, 400), nome_participante, fill='black', font=fonte_nome)
    desenhar.text((250, 250), nome_curso, fill='black', font=fonte_geral)
    desenhar.text((300, 300), tipo_participacao, fill='black', font=fonte_geral)
    desenhar.text((500, 500), str(carga_horaria), fill='black', font=fonte_geral)
    desenhar.text((650, 700), str(data_inicio), fill='blue', font=fonte_data)
    desenhar.text((850, 700), str(data_final), fill='blue', font=fonte_data)
    desenhar.text((100, 100), str(data_emissao), fill='blue', font=fonte_data)

    # Salvar o certificado com um nome único
    image.save(f'./{indice}_{nome_participante}_certificado.png')