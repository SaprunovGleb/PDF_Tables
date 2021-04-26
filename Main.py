# -*- coding: utf-8 -*-
"""
Created on Mon Jul 13 19:05:00 2020

@author: Saprunov Gleb
"""
#Включение библиотеки для работы с путями к файлам
import os  
from GetSource import getSource
from ProcessingSourse import  extrPdfFromZipTo2TXT
    
def main():
    print("Введите 0, если хотите выйти из программы")
    print("Введите 1, если хотите докачать данные из источников")
    print("Введите 2, если хотите перевести данные от cmegroup Futures из Zip в Xlsx")
    print("Введите 3, если хотите перевести данные от cmegroup Options из Zip в Xlsx")
    j = ""
    while j != "0":
        j = input()
        if  j == "1":
            getSource()
        elif j == "2":
            extrPdfFromZipTo2TXT(workFilePDF = "Section61_Energy_Futures_Products.pdf",info="Futures")
        elif j == "3":
            extrPdfFromZipTo2TXT(workFilePDF = "Section63_Energy_Options_Products.pdf",info="Options")
        if j!="0":
            print("Жду новой команды")
            
            
            
main()

