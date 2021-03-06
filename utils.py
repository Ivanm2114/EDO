import xml.etree.ElementTree as ET
import openpyxl
import shutil


def getContent(file):
    def getChilds(root):
        d = {}
        for child in root:
            if child.tag and child.attrib:
                d[child.tag] = child.attrib
            try:
                temp = getChilds(child)
                for key in temp.keys():
                    d[key] = temp[key]
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
    workbook.save('ЭДО\\result.xlsx')
