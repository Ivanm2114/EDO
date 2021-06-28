import xml.etree.ElementTree as ET
from openpyxl.styles import Border, Side, Alignment, Font
import openpyxl
import shutil
from num2words import num2words

tree = ET.parse(
    'ЭДО/ON_SCHET___20210209_dbcfe2dd-a95c-4b8f-98bc-a0b9e907adb1.xml')


def get_root(path):
    global tree
    param = ''
    value = ''
    arg = ''
    path = path.split(',')
    root = tree.getroot()
    for step in path[:-1]:
        if '(' in step:
            step, arg = step.split('(')
            arg = arg[:-1]
            param, value = arg.split('=')
        a = [x.tag for x in root]
        if a.count(step) == 1:
            root = root[a.index(step)]
        elif a.count(step) > 1:
            indexes = []
            for i in range(len(a)):
                if a[i] == step:
                    indexes.append(i)
            for i in indexes:
                try:
                    if root[i].attrib[param] == value:
                        root = root[i]
                except IndexError:
                    pass
    return root.attrib[path[-1]]


print(get_root('Документ,Параметр(Имя=ДоговорНаименование),Значение'))
