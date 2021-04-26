# -*- coding: utf-8 -*-
"""
Created on Fri Jul 17 10:22:43 2020

@author: Flyin
"""


import os
import zipfile
import io
import tabula

import fitz 
from TXTtoXLSX import txtToXlsxFut
from TXTtoXLSX import txtToXlsxOpt

def convert_pdf_to_txtfitz(pathIn,pathOut):
    doc = fitz.open(pathIn)
    file = open(pathOut, "w")
    i=0
    N=doc.pageCount
    while i < N:
        page=doc.loadPage(i)
        text=page.getText("text")
        file.write(str(text))
        i+=1
    file.close()
    
def extrPdfFromZipTo2TXT(nameSourse = "cmegroup",workFilePDF = "Section61_Energy_Futures_Products.pdf",info="Futures",col=-1):
    #Нахождение местоположение проекта на компьюторе
    path = os.getcwd()   
    #Указание адреса исходной директории
    pathDirIn = "\Source"+"\\"+nameSourse
    #Указание адреса итоговой директории
    pathDirOut = "\Data"+"\\"+nameSourse+"\\"+info
    #Задание абсолютного пути 
    pathIn = path + pathDirIn
    pathOut = path + pathDirOut
    #Список файлов в исходной директории
    fileNamesIn = os.listdir(pathIn)
    #print(fileNamesIn)
    fileNamesOut = os.listdir(pathOut)    
    #Тело цикла
    #Счетчик файлов из директории
    i=0
    if col == -1:
        maxI= len(fileNamesIn)
    else:
        maxI=col
    #Номер символа с которого идет дата
    adressDateFromName=18
    print("Исходный файл:",workFilePDF)
    while i < maxI:
        #Задание имени итогового файла
        fileNameOut=fileNamesIn[i][adressDateFromName:(adressDateFromName+8)]+".xlsx"
        #print(fileNameOut)
        if fileNamesOut.count(fileNameOut)==0:
            #Исходный архив Zip
            nameFileZip = pathIn + "\\" + fileNamesIn[i]
            #print(workFileZip)
            #Проверка на существование и то что это действительно zip архив
            if zipfile.is_zipfile(nameFileZip):
                #Открываем архив для работы
                workFileZip = zipfile.ZipFile(nameFileZip, 'r')
                #Указываем файл для обработки
                #Извлекаем файлы для обработки
                workFileZip.extract(workFilePDF)
                print("Идет создание ",fileNameOut, "исходный архив №", i)
                convert_pdf_to_txtfitz(workFilePDF,"PDFtoTXTpymupdf.txt")
                #pathFileNameOut = path + pathDirOut+"\Futures\\"+fileNameOut
                tabula.convert_into(workFilePDF,output_path = "PDFtoTXTTabula.txt", output_format="csv", pages="all")
                if info=="Futures":                           
                    txtToXlsxFut(pathOut=(pathOut+"\\"+fileNameOut))
                elif info=="Options":
                    txtToXlsxOpt(pathOut=(pathOut+"\\"+fileNameOut))
        i+=1
    print("Переформатирование завершено")

        
#extrPdfFromZipTo2TXT(workFilePDF = "Section63_Energy_Options_Products.pdf",col=1,info="Options")