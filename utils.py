import xml.etree.ElementTree as ET
from copy import copy

import openpyxl
import shutil
from openpyxl.styles import Border, Side


def getContent(file):
    def getChilds(root):
        d = {}
        for child in root:
            flag = False
            if child.tag and child.attrib:
                if child.tag not in d:
                    d[child.tag] = child.attrib
                else:
                    if not flag:
                        d[child.tag] = [d[child.tag]]
                        flag = True
                        index = 1
                        d[child.tag].append(child.attrib)
                    else:
                        d[child.tag].append(child.attrib)
                        index += 1

            try:
                temp = getChilds(child)
                for key in temp.keys():
                    if not flag:
                        d[child.tag][key] = temp[key]
                    else:
                        d[child.tag][index][key] = temp[key]
            except IndexError:
                pass
        return d

    tree = ET.parse(file)
    root = tree.getroot()[0]
    content = {root.tag: root.attrib}
    for child in root:
        if getChilds(child):
            if child.tag not in content:
                content[child.tag] = []
            content[child.tag].append(getChilds(child))
        else:
            if child.tag not in content:
                content[child.tag] = []
            content[child.tag].append(child.attrib)

    return content


def putIntoXLS(data):
    thin = 'thin'
    thick = 'thick'
    shutil.copyfile('ЭДО\\Schet.xlsx', 'ЭДО\\result.xlsx')
    workbook = openpyxl.load_workbook('ЭДО\\Schet.xlsx')
    sheet = workbook['TDSheet']
    sheet['B2'] = data['Поставщик'][0]['БанкРекв']['НаимБанк']
    sheet['AD2'] = data['Поставщик'][0]['БанкРекв']['БИК']
    sheet['AD3'] = data['Поставщик'][0]['БанкРекв']['РСчет']
    sheet['AD5'] = data['Поставщик'][0]['БанкРекв']['КСчет']
    sheet['AD5'] = data['Поставщик'][0]['БанкРекв']['КСчет']
    sheet['B6'] = data['Поставщик'][0]['СвЮЛ']['Название']
    sheet['B10'] = f'{data["Документ"]["Название"]} №{data["Документ"]["Номер"]} от {data["Документ"]["Дата"]}  г.'
    sheet['G14'] = data['Поставщик'][0]['СвЮЛ']['Название'] + ',' + data['Поставщик'][0]['Адрес']['АдрТекст']
    sheet['G17'] = data['Покупатель'][0]['СвЮЛ']['Название'] + ',' + data['Покупатель'][0]['Адрес']['АдрТекст']
    sheet['G20'] = data['Параметр'][0]['Значение']
    row = 23
    i = 1
    for el in data['ТаблДок'][0]['СтрТабл']:
        sheet.merge_cells(f'B{row}:C{row}')
        sheet.merge_cells(f'D{row}:X{row}')
        sheet.merge_cells(f'Y{row}:AB{row}')
        sheet.merge_cells(f'AC{row}:AE{row}')
        sheet.merge_cells(f'AF{row}:AJ{row}')
        sheet.merge_cells(f'AK{row}:AQ{row}')
        sheet.row_dimensions[row].height = 50
        sheet[f'B{row}'] = i
        sheet[f'D{row}'] = el['Название']
        i += 1
        row += 1
    new_cell = sheet['AU17']
    cell = sheet['D23']
    for row in sheet.rows:
            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.border = copy(cell.border)
                new_cell.fill = copy(cell.fill)
                new_cell.number_format = copy(cell.number_format)
                new_cell.protection = copy(cell.protection)
                new_cell.alignment = copy(cell.alignment)

    sheet.row_dimensions[24].height = 50

    workbook.save('ЭДО\\result.xlsx')
