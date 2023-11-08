import xml.etree.ElementTree as ET
import openpyxl
import fileinput
import sys

#указываем необходимое значение для генерации в теге и входные выходные файлы 
gen = "rand(999,9975)"
xlsfile = 'kgm3.xlsx' #xls
xmlfile = 'conf-old.xml' #xml
xmltarget = 'test_anton.xml' #xml-target

#подгружаем exel файл
wb = openpyxl.load_workbook(xlsfile)
#определяем активный лист эксель файла
sheet = wb.active
#подгружаем файл с деревом xml
tree = ET.parse(xmlfile)
#определяем корневой элемент дерева xml
root = tree.getroot()

#пробегаем по строкам из екселя, колонка A
for i in sheet['A']:
    result = ''
    #ищем внутри тегов атрибут name
    for j in root.findall(".//PSTAlias[@name='%s']" % i.value.replace('.', '_')):
        result = j.get('name')
    #если значения из екселя нет в дереве тегов (по атрибуту name), то создаем новые теги    
    if result == '':
        #задаем атрибуты тега PSTAlias, в конце меняем значения для генерации
        attrib = { 'name': i.value.replace('.', '_'), 'itemPath': "", 'type': "2", 'updateRate': "60000",  'calcEquation': gen, }
        #Создаем тег PSTAlias
        element = root[1][2].makeelement('PSTAlias', attrib)
        root[1][2].append(element)
        #задаем атрибуты тега Scaling
        attrib_scaling = { 'enabled': "0", 'type': "0" }
        #Создаем тег Scaling
        subelement = root[1][2][-1].makeelement('Scaling', attrib_scaling)
        root[1][2][-1].append(subelement)
        #задаем атрибуты тега Events
        attrib_events = { 'enabled': "0", 'source': "Alias", 'severity': "1", 'trigger': "0", 'timestamp': "0" }
        #Создаем тег Scaling
        subelement2 = root[1][2][-1].makeelement('Events', attrib_events)
        root[1][2][-1].append(subelement2)
    else:        
        #если тег уже есть в списке, то меняем ему значение генерации на нужное
        for j in root.findall(".//PSTAlias[@name='%s']" % i.value.replace('.', '_')):
            j.set('calcEquation', gen)
          
#Функция по красивому форматированию 
def _pretty_print(current, parent=None, index=-1, depth=0):
    for i, node in enumerate(current):
        _pretty_print(node, current, i, depth + 1)
    if parent is not None:
        if index == 0:
            parent.text = '\n' + ('\t' * depth)
        else:
            parent[index - 1].tail = '\n' + ('\t' * depth)
        if index == len(parent) - 1:
            current.tail = '\n' + ('\t' * (depth - 1))

#форматируем итоговое дерево
_pretty_print(root)

#записываем итоговое дерево в файл
tree.write(xmltarget)

#Какого-то хрена библиотека ElementTree меняет символы & на &amp; - поправляем заменой
for line in fileinput.input(xmltarget, inplace=True):
    sys.stdout.write(line.replace('&amp;', '&'))