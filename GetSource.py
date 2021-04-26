# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
#Включение библиотеки для работы с File Transfer Protocol (FTP)
import ftplib
#Включение библиотеки для работы с путями к файлам
import os
from ProcessingSourse import  extrPdfFromZipTo2TXT

def getSource(pathIn="\PathResource.txt"):
    #Нахождение местоположение проекта на компьюторе
    path = os.getcwd()   
    #Указание адреса исходного файла
    pathTableIn = path + pathIn 
    #открыли исходный файл для чтения
    fileTableIn = open(pathTableIn,'r')
    for lineFromTableIn in fileTableIn:
        #удалили последний символ в строке, а именно перенос строки
        lineFromTableIn = lineFromTableIn[0:-1]
        #разделиили строку на массив слов
        massLineFromTableIn = lineFromTableIn.split(' ')
        #вывели строку в виде массива для проверки. 
        #print(massLineFromTableIn)
        if len(massLineFromTableIn)==3:
            getSourceFromFTP(massLineFromTableIn[0],massLineFromTableIn[1],massLineFromTableIn[2])
        elif len(massLineFromTableIn)==5:
            getSourceFromFTP(massLineFromTableIn[0],massLineFromTableIn[1],massLineFromTableIn[2],
            massLineFromTableIn[3],massLineFromTableIn[4])
                
                
def getSourceFromFTP(dirOut, pathFTP, dirFTP, userName = "Null", userPassword = "Null"):
    print("Проверяю FTP:  "+pathFTP)
    #Нахождение местоположение проекта на компьюторе
    path = os.getcwd()
    #Проверка местоположения проекта на компьюторе
    #print (path) 
    #Создание пути для хранения исходных данных
    pathOut = path + dirOut
    #Подключение к серверу из которого будем брать файлы
    ftp = ftplib.FTP (pathFTP)
    #Подключение к каталогу
    if userName == "Null":
        ftp.login()
    else:
        ftp.login(userName,userPassword)
    # =============================================================================
    # #создание списка каталогов в корне сервера и их вывод
    # data = ftp.retrlines('LIST')
    # print(data)
    # =============================================================================
    # Меняем директорию
    ftp.cwd(dirFTP)
    # =============================================================================
    # #создание списка каталогов в директории сервера и их вывод
    # data = ftp.retrlines('LIST')
    # print(data)
    # =============================================================================
    #создание списка имен файлов и директорий в директории сервера и их вывод
    fileNamesServer = ftp.nlst()
    #print(fileNamesServer)
    #создание списка имен файлов и директорий в конечной директории и их вывод
    fileNamesExist = os.listdir(pathOut) 
    #print(fileNamesExist)
    i = 0
    j = 0
    while i < len(fileNamesServer):
        # проверяем есть ли имя которое мы желаем скачать в списке уже существующих файлов
        if fileNamesExist.count(fileNamesServer[i]) == 0:
            #открываем файл для записи
            print("Скачивается "+str(j+1)+" файл")
            with open(pathOut+"\\"+fileNamesServer[i], 'wb') as f:
                #записываем скачанный файл
                ftp.retrbinary('RETR ' + fileNamesServer[i], f.write)
            #нумерация для показа прогресса скачивания файлов
            #print(str(i) + "+")
            #закрытие файла для записи
            f.close()
            j += 1
    # =============================================================================
    #     else:
    #         print(str(i) + "-")
    # =============================================================================
        i += 1
    if j == 0:
        print("Всё уже было скачано")
    else:
        if j == 1:
            print("Скачан "+str(j)+" файл")
        elif (j > 1) and (j < 5):
            print("Скачано "+str(j)+" файла")
        else:
            print("Скачано "+str(j)+" файлов")
    #input()
    return j
    
#getSource()   
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    