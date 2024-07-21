from mergeExcelBooks import merge
import openpyxl #для работы с файлами Excel
import os
from templateFilling import excelContext
from docxtpl import DocxTemplate #для работы с  документами Word


resultsFileName = 'diplom/workingExcelDiplom/resultsFileName.xlsx' # путь и имя файла Excel, куда сохранятся все гниги и из которого будут браться данные.
excelBooks = 'diplom/excelBooksDiplom' # путь откуда брать рабочие программы преподавателей (в формате xlsx)

merge(resultsFileName, excelBooks) # Собираем результирующую большую книгу
resultingBook = openpyxl.load_workbook(resultsFileName, data_only=True) #Открывается файл

# allSheets = resultingBook.sheetnames # список всех листов

fileNameinField = 'Fname' # имя поля, в котором указывается имя файла для сохранения
wordFileName = 'diplom/TemplateRPVNew.docx' #имя шаблона документа Word
saveFolder = 'diplom/workProgramsDiplom' #папка для сохранения

def buildReports(wordFileName, context, saveFolder, fn ): # Имя шаблона, данные, папка для сохранения, имя вайла для сохранения
    doc = DocxTemplate(wordFileName)
    doc.render(context)
    doc.save(str(saveFolder + '/' + context[fileNameinField])+".docx") # Было!!

    # doc.save(str(saveFolder + '/' + fn)+".docx") # Стало!!

for sheet in resultingBook:
    context = excelContext(resultsFileName, sheet.title) # Получаем данные из excel
    buildReports(wordFileName, context, saveFolder,fileNameinField)
    print(sheet.title)
    print(context)



# создаем папку где будут лежать все рабочие программы
# os.mkdir("Work programs") # путь относительно текущего скрипта
# os.mkdir("c://somedir/Work programs") # абсолютный путь

# имя файла, имя листа, имя поля, в котором указывается имя файла для сохранения, имя шаблона документа Word, папку куда сохранить

