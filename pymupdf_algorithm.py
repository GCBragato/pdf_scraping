# Programa que isola páginas de um PDF com base em um texto
# Se o texto existe na página, a mantém

import PyPDF2
import pdf2image
import os
import xlsxwriter
import time
import fitz
from fitz import Document, Page, Rect


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


def get_data():

    # INPUT AQUI
    pdf_in = '_SPTs.pdf'

    # Retângulos contendo os textos
    # Obtidos usando Photoshop
    furo_rect = Rect(506, 115, 549, 128)
    spt_rect = Rect(388, 201, 422, 677)
    na_rect = Rect(464, 760, 501, 767)

    # Documento fitz
    doc: Document = fitz.open(pdf_in)

    # Loop de leitura
    spt_data = []
    for i in range(len(doc)):
        page: Page = doc[i]
        furo = page.get_textbox(furo_rect)
        spt = page.get_textbox(spt_rect).splitlines()
        na = page.get_textbox(na_rect)
        spt_data.append([furo, spt, na])
        print(f'Furo: {furo} | SPT: {spt} | NA: {na}')

    # Loop de escrita de Excel com filtragem de valores
    workbook = xlsxwriter.Workbook('_SPTs_pymu.xlsx')
    worksheet = workbook.add_worksheet()
    for i,data in enumerate(spt_data):
        worksheet.write(i, 0, data[0])
        for j in range(len(data[1])):
            try:
                spt_value = float(data[1][j])
            except:
                if data[1][j] == '-':
                    spt_value = 0
                else:
                    spt_value = data[1][j]
            worksheet.write(i,j+1,spt_value)
        try:
            na_value = float(data[2])
        except:
            na_value = data[2]
        worksheet.write(i,19,na_value)
    workbook.close()

    return

if __name__ == '__main__':
    start = time.perf_counter()
    get_data()
    end = time.perf_counter()
    print(f'Finished in {round(end-start, 2)} second(s)') 
    exit(0)
