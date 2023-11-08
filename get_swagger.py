import requests
import openpyxl
import concurrent.futures

# прочесть csv файл с первой строки
# взять нулевой столбец (id)
# результат массив айдишников
# далее подготовить цикл по массиву и в каждой итерации выполнять метод get/del

standurl = 'https://***apiurl'
token = 'Bearer ********token'
#headers = {'accept': 'text/plain',  'Content-Type': 'application/json', 'Authorization': token}
headers = {'accept': 'text/plain', 'Authorization': token}
#r= requests.get(f'{standurl}/api/mnemo/{id}/configuration', headers=headers)
#r= requests.request('GET', standurl, headers=headers)
#print(r.status_code)
#print(r.json())

#1 входные файлы
xlsfile = 'Книга2.xlsx' #xls

#2 подгружаем exel файл
wb = openpyxl.load_workbook(xlsfile)

# 3 определяем активный лист эксель файла
sheet = wb.active

#функция с запросом к нужному методу сваггера
def del_mnemo(i):
    r = requests.delete(f"{standurl}/api/mnemo/{i.value}", headers=headers)
    print(r.status_code)

# 4 пробегаем по строкам из екселя, колонка A
with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
    for i in sheet['A']:
        executor.submit(del_mnemo, i)


