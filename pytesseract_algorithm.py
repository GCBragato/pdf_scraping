# Programa que isola páginas de um PDF com base em um texto
# Se o texto existe na página, a mantém

import PyPDF2
import pdf2image
import os
from PIL import Image
from pytesseract import pytesseract
import numpy as np
import cv2
import xlsxwriter
import time


def isolate_pages():

    # INPUT AQUI
    pdf_in = '_SPT1 ao SPT113.pdf'
    pdf_out = '_SPTs.pdf'
    text_filter = 'CLASSIFICAÇÃO DO MATERIAL'

    reader = PyPDF2.PdfReader(pdf_in,'rb')
    output = PyPDF2.PdfWriter()

    pages = reader.pages
    #print(pages[0].extract_text())
    for i,page in enumerate(pages):
        text = page.extract_text()
        if text_filter in text:
            output.add_page(page)

    with open(pdf_out,'wb') as f:
        output.write(f)
    return


def pdf_to_png():

    # INPUT AQUI
    pdf_in = 'viz__SPTs_test.pdf'

    if not os.path.exists('./png_out'):
        os.makedirs('./png_out')

    poppler = r'C:\Users\gusta\AppData\Roaming\poppler-22.04.0\Library\bin'
    pages = pdf2image.convert_from_path(pdf_in,poppler_path=poppler,dpi=600)
    for i,page in enumerate(pages):
        page.save('./png_out/%s.png' % str(i), 'PNG')
    return


def clean_image_to_ocr(img):

    # convertendo em um array editável de numpy[x, y, CANALS]
    npimagem = np.asarray(img).astype(np.uint8)
    # diminuição dos ruidos antes da binarização
    npimagem[:, :, 0] = 0 # zerando o canal R (RED)
    npimagem[:, :, 2] = 0 # zerando o canal B (BLUE)
    # atribuição em escala de cinza
    im = cv2.cvtColor(npimagem, cv2.COLOR_RGB2GRAY)
    # aplicação da truncagem binária para a intensidade
    # pixels de intensidade de cor abaixo de 127 serão convertidos para 0 (PRETO)
    # pixels de intensidade de cor acima de 127 serão convertidos para 255 (BRANCO)
    # A atrubição do THRESH_OTSU incrementa uma análise inteligente dos nivels de truncagem
    ret, thresh = cv2.threshold(im, 127, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
    # reconvertendo o retorno do threshold em um objeto do tipo PIL.Image
    binimagem = Image.fromarray(thresh)
    #binimagem.save('bin.png')

    return binimagem


def get_data(tessdatapath,xlsxpath):

    # INPUT AQUI
    png_in = './png_out/'

    spt_coords = [3310,1668]
    na_coords = [3866,6345]
    furo_coords = [4226,962]

    spt_size = [222,3987]
    na_size = [320,57]
    furo_size = [376,111]

    # Quantidade de páginas
    n_png = len([name for name in os.listdir(png_in)
        if os.path.isfile(os.path.join(png_in, name))])

    # OCR - Tesseract
    path_to_tesseract = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    pytesseract.tesseract_cmd = path_to_tesseract
    os.environ['TESSDATA_PREFIX'] = tessdatapath


    # Loop de corte e leitura de imagens
    spt_data = []
    for i in range(n_png):
        img = Image.open(png_in + str(i) + '.png')
        # Corte
        img_spt = img.crop((spt_coords[0],spt_coords[1],
            spt_coords[0]+spt_size[0],spt_coords[1]+spt_size[1]))
        img_na = img.crop((na_coords[0],na_coords[1],
            na_coords[0]+na_size[0],na_coords[1]+na_size[1]))
        img_furo = img.crop((furo_coords[0],furo_coords[1],
            furo_coords[0]+furo_size[0],furo_coords[1]+furo_size[1]))
        # Limpeza das imagens
        img_spt = clean_image_to_ocr(img_spt)
        img_na = clean_image_to_ocr(img_na)
        img_furo = clean_image_to_ocr(img_furo)
        # Leitura
        tess_config = '--psm 6 --oem 1 -c tessedit_char_whitelist=.SECOPT0123456789'
        spt = pytesseract.image_to_string(img_spt,config=tess_config).replace(
                '\n\n','\n').splitlines()
        na = pytesseract.image_to_string(img_na,config=tess_config).replace(
                '\n\n','\n').splitlines()
        furo = pytesseract.image_to_string(img_furo,config=tess_config).replace(
                '\n\n','\n').splitlines()
        spt_data.append([spt,na,furo])
        #print([spt,na,furo])

    # Loop de escrita de Excel
    workbook = xlsxwriter.Workbook(xlsxpath)
    worksheet = workbook.add_worksheet()
    for i,data in enumerate(spt_data):
        worksheet.write(i,0,data[2][0])
        for j in range(len(data[0])):
            worksheet.write(i,j+1,data[0][j])
        worksheet.write(i,19,data[1][0])
    workbook.close()

    return

if __name__ == '__main__':
    start = time.perf_counter()
    # tessdatapath = r"C:\Program Files\Tesseract-OCR\tessdata-main"
    # xlsxpath = '_SPTs_normal.xlsx'
    # get_data(tessdatapath,xlsxpath)
    # tessdatapath = r"C:\Program Files\Tesseract-OCR\tessdata_fast-main"
    # xlsxpath = '_SPTs_fast.xlsx'
    # get_data(tessdatapath,xlsxpath)
    # tessdatapath = r"C:\Program Files\Tesseract-OCR\tessdata_best-main"
    # xlsxpath = '_SPTs_best.xlsx'
    # get_data(tessdatapath,xlsxpath)
    pdf_to_png()
    end = time.perf_counter()
    print(f'Finished in {round(end-start, 2)} second(s)') 
    exit(0)
