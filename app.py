# tipo do curso, nome participante, tipo de participação, data de inicio,
#  datra final
# carga horária, data de emissão de certficação e as assinaturas do
# gestor geral, do coordenador.


# pegar os dados da planilha :
# tipo do curso, nome participante, tipo de participação, data de inicio,
# datra final
# carga horária, data de emissão de certficação e as assinaturas
# do gestor geral, do coordenador.


# transferir os dados da planilha para a imagem do certificado
# pegar os dados da planilha
# para importar uma planilha
import openpyxl
# transferir os dados da planilha para a imagem do certificado
from PIL import Image, ImageDraw, ImageFont

workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
Sheet_alunos = workbook_alunos['Sheet1']

# para acessar cada linha da planilha
for indice, linha in enumerate(Sheet_alunos.iter_rows(min_row=2, max_row=10)):
    # acessar célula que contém a info que precisamos
    nome_curso = linha[0].value  # nome do curos
    nome_do_participante = linha[1].value  # nome do participante
    tipo_participacao = linha[2].value  # tipo de participação
    data_inicio = linha[3].value  # data de inicio
    data_final = linha[4].value  # data de final
    carga_horaria = linha[5].value  # carga horaria
    data_emissao = linha[6].value  # data de emissão


# transferir os dados da planilha para a imagem do certificado
    # definindo a fonte a ser usada
    fonte_nome = ImageFont.truetype('./tahomabd.ttf', 90)
    fonte_geral = ImageFont.truetype('./tahoma.ttf', 80)
    fonte_data = ImageFont.truetype('./tahoma.ttf', 55)
    # abrir a imagem e sobrepor um texto sobre ela
    imagem = Image.open('./certificado_padrao.jpg')
    # Para poder sobreescrever na imagem
    desenhar = ImageDraw.Draw(imagem)
    desenhar.text((1020, 827), nome_do_participante,
                  fill='black', font=fonte_nome)
    desenhar.text((1060, 950), nome_curso,
                  fill='black', font=fonte_geral)
    desenhar.text((1435, 1065), tipo_participacao,
                  fill='black', font=fonte_geral)
    desenhar.text((1480, 1182), str(carga_horaria),
                  fill='black', font=fonte_geral)

    desenhar.text((750, 1770), data_inicio,
                  fill='black', font=fonte_data)
    desenhar.text((750, 1930), data_final,
                  fill='black', font=fonte_data)

    desenhar.text((2220, 1930), data_emissao,
                  fill='blue', font=fonte_data)

    imagem.save(f'./{indice}{nome_do_participante} certificado.png')
