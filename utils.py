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
    workbook = openpyxl.load_workbook('ЭДО\\result.xlsx')
    sheet = workbook['TDSheet']
    cell = sheet.cell(row=4, column=2)
    print(sheet['B4'].value)
    for cell in sheet:
        print(cell)
