import pdf2image
import os
from pytesseract import pytesseract
import numpy as np
import cv2
from PIL import Image
import xlsxwriter


def get_data():

    # 1 - Abrir PDF
    pdf = PyPDF2.PdfFileReader('_SPTs_teste.pdf','rb')
    n_pages = pdf.getNumPages()

    # temporário - extrair tamanho da página
    # page = pdf.getPage(0)
    # coor_ll = page.cropBox.getLowerLeft()
    # coor_ur = page.cropBox.getUpperRight()
    # width = coor_ur[0] - coor_ll[0]
    # height = coor_ur[1] - coor_ll[1]
    # print(width,height) # 595.32 841.92

    # 2 - Cortar páginas e escrever em novo PDF

    spt_coords = (399,841-675)
    na_coords = (465,841-765)
    furo_coords = (511,841-128)

    spt_size = (23,473)
    na_size = (35,4)
    furo_size = (35,14)


    writer = PyPDF2.PdfFileWriter()
    for i in range(n_pages):
        page = pdf.getPage(i)
        spt_page = page
        spt_page.mediaBox.setLowerLeft((spt_coords))
        spt_page.mediaBox.setUpperRight((spt_coords[0]+spt_size[0],spt_coords[1]+spt_size[1]))
        spt_page.cropBox.setLowerLeft((spt_coords))
        spt_page.cropBox.setUpperRight((spt_coords[0]+spt_size[0],spt_coords[1]+spt_size[1]))
        spt_page.bleedBox.setLowerLeft((spt_coords))
        spt_page.bleedBox.setUpperRight((spt_coords[0]+spt_size[0],spt_coords[1]+spt_size[1]))
        print(spt_page.extract_text())
        writer.addPage(spt_page)
    with open('_SPTs_crop.pdf','wb') as f:
        writer.write(f)

    return


if __name__ == '__main__':
    get_data()
    print('Done')
    exit(0)