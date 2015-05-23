# -*- coding: utf8 -*-
__author__ = 'mzh'

try:
    import xml.etree.cElementTree as ET
except ImportError:
    import xml.etree.ElementTree as ET

#读取xml文件，将文件转为一个列表，每一条数据是一个map。
def readXml(filePath):

        tree = ET.ElementTree(file=filePath)
        retList = []
        for singleTemplate in tree.getroot():
            tMap = {}
            for child in singleTemplate:
                if child.tag == 'sheetname':
                    tMap['sheetname'] = child.text

                elif child.tag == 'ascolumnname':
                    tMap['ascolumnname'] = child.text

                elif child.tag == 'count':
                    tMap['count'] = child.text

                elif child.tag == 'sum':
                    strTemp = []
                    for temp in child:
                        strTemp.append(temp.text)
                    tMap['sum'] = strTemp

                elif child.tag == 'groupby':
                    strTemp = []
                    for temp in child:
                        strTemp.append(temp.text)
                    tMap['groupby'] = strTemp

                elif child.tag == 'where':
                    tMap['where'] = child[0].text

                elif child.tag == 'coefficient':
                    tMap['coefficient'] = child.text
                elif child.tag == 'state':
                    tMap['state'] = child.text

            retList.append(tMap)
        return retList

def readCoefficientFromXml(filePath=None):
    tree = ET.ElementTree(file=filePath)
    retMap = {}
    for co in tree.getroot():
        name = co.find('name').text
        value = co.find('value').text
        retMap[name] = value
    return retMap


def writeXml(filePath, billModel, headList):

    insertLines = []
    rowCount = billModel.rowCount()
    columnCount = billModel.columnCount()
    for i in range(rowCount):
        tempModelIndex = billModel.index(i, 7)
        state = tempModelIndex.data().toInt()
        if state[1] and state[0] == 1:
            tMap = {}
            for j in range(columnCount-1):
                tmi = billModel.index(i, j)
                tvalue = tmi.data().toString().__str__()
                tMap[headList[j]] = tvalue
            tMap['state'] = '0'
            insertLines.append(tMap)

     #检查是否有数据需要插入
    needInsertLength = len(insertLines)
    if needInsertLength <=0:
        return

    tree = ET.ElementTree(file=filePath)
    root = tree.getroot()

    for tempMap in insertLines:
        attrMap = {'index': '3'}
        newTemplate = ET.Element('template', attrMap)
        for key in tempMap:
            if key == 'where':
                newItem = ET.Element(key)
                childItem = ET.Element('wh', {'index': '88'})
                childItem.text = tempMap[key]
                newItem.append(childItem)
            elif key == 'groupby':
                newItem = ET.Element(key)
                childItem = ET.Element('gr', {'index': '88'})
                childItem.text = tempMap[key]
                newItem.append(childItem)
            elif key == 'sum':
                newItem = ET.Element(key)
                childItem = ET.Element('su', {'index': '88'})
                childItem.text = tempMap[key]
                newItem.append(childItem)
            else:
                newItem = ET.Element(key)
                newItem.text = tempMap[key]

            newTemplate.append(newItem)
        root.append(newTemplate)

        tree.write(filePath, 'utf-8')