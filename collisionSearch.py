from deepdiff import DeepDiff


dict1 = {'title': {'jr': 3, 'f': 3}, 'description': '64', 'price': '4'}
dict2 = {'python': 'dede', 'key': '#789', 'title': {'jr': 3, 'f': 3}, 'description': '64', 'price': '3'}


# Найти общие ключи
result = dict1.keys() & dict2.keys()

# Найти где и в каком словаре не хватает ключей 1
print(dict1.keys()-dict2.keys())
print(dict2.keys()-dict1.keys())

# Найти где и в каком словаре не хватает ключей 2
result = set(dict1) - (set(dict2))
if result:
    print(f'В dict2 нет ключей dict1: {result}')
else:
    print('В dict1 все ключи dict2', result)

result = set(dict2) - (set(dict1))
if result:
    print(f'В dict1 нет ключей dict2: {result}')
else:
    print('В dict2 все ключи dict1', result)

# Словарь result будет содержать только те ключи, которые присутствуют как в dict1, так и в dict2, с идентичными значениями.
result = {k: dict1[k] for k in dict1 if k in dict2 and dict1[k] == dict2[k]}
print('result', result)

# глубокая проверка (большая вложеность)
result = DeepDiff(dict1, dict2)
print('DeepDiff', result)



# Тестирую заполнения из бд
from templateFilling import buildReports
from db import databaseQuery


context = databaseQuery()
buildReports('test/TemplateTestDB.docx', context[0], 'test', 'testReportsBD')# Имя шаблона, данные, папка для сохранения, имя вайла для сохранения