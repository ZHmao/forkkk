# -*- coding: utf-8 -*-

import xlrd
import pandas as pd

def main():

    sheetNameList = [u'保险', u'基金', u'分期付款', u'资产业务', u'积存金', u'理财产品', u'个人电子银行', u'银行卡']

    excelFile = 'c:\\123.xls'
    xlsDf = pd.read_excel(excelFile, u'理财产品')
    #print type(xlsDf)
    #print xlsDf.head(20)

    print '-----------------------------'
    selectedDf = xlsDf.loc[:, [u'营销人员', u'金额']]
    #print selectedDf.head()

    print '============================'
    groupedDf = selectedDf.groupby(u'营销人员').sum()
    #print groupedDf.head(2)

    print '============================'

    all_sheet_df = pd.read_excel(excelFile, sheetNameList)
    for sheetName in sheetNameList:
        tempDf = all_sheet_df[sheetName]
        if sheetName == u'保险':
            selectedDf = tempDf.loc[:, [u'营销人员', u'产品名称', u'金额']]
            groupedDf = selectedDf.groupby([u'营销人员', u'产品名称']).sum()
            print '=================================='
            print sheetName
            print groupedDf.head()
            print '=================================='
        elif sheetName == u'基金':
            selectedDf = tempDf.loc[:, [u'营销人员', u'金额']]
            groupedDf = selectedDf.groupby(u'营销人员').sum()
            print '=================================='
            print sheetName
            print groupedDf.head()
            print '=================================='
        elif sheetName == u'分期付款':
            selectedDf = tempDf.loc[:, [u'营销人员', u'金额']]
            groupedDf = selectedDf.groupby(u'营销人员').sum()
            print '=================================='
            print sheetName
            print groupedDf.head()
            print '=================================='



main()
