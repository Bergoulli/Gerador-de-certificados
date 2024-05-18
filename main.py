#Autor: Mateus Costa
#Data: 18/05/2024

import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

planilha = pd.read_excel(r'Certificados_DiaEnergia_(respostas).xlsx') #carregando planilha

CERTI = "certificado_novo.jpg" #coloque o certificado
ASSINA = "assinatura2.png" #coloque a assinatura

#Edição da planilha
#print(planilha)
nomes = planilha['Nome Completo'].tolist()
mat = planilha['Número de Matrícula'].tolist()
emails = planilha['E-mail'].tolist()

def certificado(name, matricula):
    imagem = Image.open(CERTI) #carregando imagem
    logo_img = Image.open(ASSINA)
    imagem.paste(logo_img, ((imagem.width - logo_img.width) // 2, 10), logo_img)

    #Edição da imagem
    desenho = ImageDraw.Draw(imagem)
    texto = f"Certificamos que {name} registrado"
    texto2 = f"na matrícula de N° {matricula}, participou da palestra online"
    texto3 = "do Dia Internacional da Energia, realizado no dia 29/05/2024"
    texto4 = "pelo PET, totalizando uma carga horária de 01 hora/aula."
    fonte = ImageFont.truetype("arial.ttf", 30)

    # Calcular a altura total de todos os textos combinados
    todos_textos = [texto, texto2, texto3, texto4]
    altura_total_textos = sum(desenho.textbbox((0, 0), t, font=fonte)[3] - desenho.textbbox((0, 0), t, font=fonte)[1] for t in todos_textos)
    posicao_y = (imagem.size[1] - altura_total_textos) // 2 - 50

    # Adicionar os textos à imagem
    for t in todos_textos:
        largura_texto = desenho.textbbox((0, 0), t, font=fonte)[2] - desenho.textbbox((0, 0), t, font=fonte)[0]
        posicao_x = (imagem.size[0] - largura_texto) // 2
        desenho.text((posicao_x, posicao_y), t, font=fonte, fill=(0, 105, 109))
        posicao_y += desenho.textbbox((0, 0), t, font=fonte)[3] - desenho.textbbox((0, 0), t, font=fonte)[1]

    # Salvar a imagem modificada
    imagem.save(f"certificados/certificado_{matricula}.jpg")
    imagem.save(f"certificados/certificado_{matricula}.pdf")
id =0;
for nome in nomes:
    certificado(nomes[id], mat[id])
    id = id+1

