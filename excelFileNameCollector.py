# Модуль получения всех названий файлов рабочих программ. Берутся все назавния файлов с расширение xlx и записываются в массив
import os



def collector(example_dir):
    content = os.listdir(example_dir)
    excelFileName = []
    for file in content:
        if os.path.isfile(os.path.join(example_dir, file)) and file.endswith('.xlsx'):
            excelFileName.append(file)
    return excelFileName
