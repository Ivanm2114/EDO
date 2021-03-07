import xml.etree.ElementTree as ET
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
import openpyxl
import shutil


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
    thick = 'medium'
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
    arr = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W',
           'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO',
           'AP', 'AQ']

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

        for letter in arr:
            if letter == 'B':
                sheet[f'{letter}{row}'].border = Border(top=Side(border_style=thin, color='FF000000'),
                                                        right=Side(border_style=thin, color='FF000000'),
                                                        bottom=Side(border_style=thin, color='FF000000'),
                                                        left=Side(border_style=thick, color='FF000000'))
            elif letter == 'AQ':
                sheet[f'{letter}{row}'].border = Border(top=Side(border_style=thin, color='FF000000'),
                                                        right=Side(border_style=thin, color='FF000000'),
                                                        bottom=Side(border_style=thin, color='FF000000'),
                                                        left=Side(border_style=thick, color='FF000000'))
            else:
                sheet[f'{letter}{row}'].border = Border(top=Side(border_style=thin, color='FF000000'),
                                                        right=Side(border_style=thin, color='FF000000'),
                                                        bottom=Side(border_style=thin, color='FF000000'),
                                                        left=Side(border_style=thin, color='FF000000'))

            sheet[f'{letter}{row}'].alignment = Alignment(horizontal='left',
                                                          vertical='top',
                                                          text_rotation=0,
                                                          wrap_text=True,
                                                          shrink_to_fit=False, indent=0)

        i += 1
        row += 1

    row -= 1
    for letter in arr:
        if letter == 'B':
            sheet[f'{letter}{row}'].border = Border(top=Side(border_style=thin, color='FF000000'),
                                                    right=Side(border_style=thin, color='FF000000'),
                                                    bottom=Side(border_style=thick, color='FF000000'),
                                                    left=Side(border_style=thick, color='FF000000'))
        elif letter == 'AQ':
            sheet[f'{letter}{row}'].border = Border(top=Side(border_style=thin, color='FF000000'),
                                                    right=Side(border_style=thin, color='FF000000'),
                                                    bottom=Side(border_style=thick, color='FF000000'),
                                                    left=Side(border_style=thick, color='FF000000'))
        else:
            sheet[f'{letter}{row}'].border = Border(top=Side(border_style=thin, color='FF000000'),
                                                    right=Side(border_style=thin, color='FF000000'),
                                                    bottom=Side(border_style=thick, color='FF000000'),
                                                    left=Side(border_style=thin, color='FF000000'))

    row += 1

    workbook.save('ЭДО\\result.xlsx')
