# -*- coding: utf-8 -*-

import xlrd
import pandas as pd
import pubtool
import os

'''
author mzh
since 2015-4-26
'''

def getData():
    excel_file = 'c:\\123.xls'
    template_list = pubtool.readXml(os.getcwd()+'\\template.xml')
    coefficient_map = pubtool.readCoefficientFromXml(os.getcwd()+'\\coefficient.xml')
    all_sheet_name = []
    for sheet_map in template_list:
        all_sheet_name.append(sheet_map['sheetname'])
    all_original_df = pd.read_excel(excel_file, all_sheet_name)

    xstj_df = []
    zs_df = []
    for sheet_map in template_list:
        print sheet_map['sheetname']
        select_fields = []
        if sheet_map['sum'] is not None:
            select_fields.extend(sheet_map['sum'])
        if sheet_map['groupby'] is not None:
            select_fields.extend(sheet_map['groupby'])

        init_df = all_original_df[sheet_map['sheetname']]

        """根据where条件过滤"""
        if sheet_map['where'] is not None:
            column_and_value = sheet_map['where'].split('=')
            where_df = init_df[ init_df[column_and_value[0]] == column_and_value[1] ]
            #where_df = init_df.where(init_df[3:3] == u'芯片卡')
        else:
			where_df = init_df

        """筛选所需的列"""
        selected_df = where_df.loc[:, select_fields]

        #data_dict = selected_df.to_dict()
        #print data_dict



        """分组合并"""
        grouped_df = selected_df.groupby(sheet_map['groupby'])
        if sheet_map['count'] == 0:
            grouped_df = grouped_df.sum().reset_index()
        else:
            grouped_df = grouped_df.count().reset_index()

        """用作合计"""
        grouped_df[sheet_map['sheetname']] = 0
		
        temp_xstj_df = None
        temp_zs_df = None
        """state为0表示是销售统计的模板，1表示是中收的模板，2表示是公用的模板。"""
        if sheet_map['state'] == '0':
            temp_xstj_df = grouped_df.copy()
        elif sheet_map['state'] == '1':
            temp_zs_df = grouped_df.copy()
        else:
            temp_xstj_df = grouped_df.copy()
            temp_zs_df = grouped_df.copy()


        """计算值"""
        """中收"""
        if temp_zs_df is not None:
            if len(sheet_map['groupby']) == 1:
                for row_index, row in temp_zs_df.iterrows():
                    temp_zs_df.loc[row_index, sheet_map['sheetname']] = row[sheet_map['sum']][0] * float(coefficient_map.get(sheet_map['sheetname'], 1))
                temp_zs_df = temp_zs_df.loc[:, [sheet_map['groupby'][0], sheet_map['sheetname']]]
            elif len(sheet_map['groupby']) == 2:
                for row_index, row in temp_zs_df.iterrows():
                    temp_zs_df.loc[row_index, sheet_map['sheetname']] = row[sheet_map['sum']][0] * float(coefficient_map.get(row[sheet_map['groupby'][1]], 1))
                temp_zs_df = temp_zs_df.loc[:, [sheet_map['groupby'][0], sheet_map['sheetname']]]
                temp_zs_df = temp_zs_df.groupby(sheet_map['groupby'][0]).sum().reset_index()
            zs_df.append(temp_zs_df)

        """销售统计"""
        if temp_xstj_df is not None:
            if len(sheet_map['groupby']) == 1:
                for row_index, row in temp_xstj_df.iterrows():
                    temp_xstj_df.loc[row_index, sheet_map['sheetname']] = row[sheet_map['sum']][0]
                temp_xstj_df = temp_xstj_df.loc[:, [sheet_map['groupby'][0], sheet_map['sheetname']]]
            elif len(sheet_map['groupby']) == 2:
                for row_index, row in temp_xstj_df.iterrows():
                    # i think it should row[sheet_map['sum']], but actual is row[sheet_map['sum']][0]
                    temp_xstj_df.loc[row_index, sheet_map['sheetname']] = row[sheet_map['sum']][0] * float(coefficient_map.get(row[sheet_map['groupby'][1]], 1))
                    print row
                temp_xstj_df = temp_xstj_df.loc[:, [sheet_map['groupby'][0], sheet_map['sheetname']]]
                temp_xstj_df = temp_xstj_df.groupby(sheet_map['groupby'][0]).sum().reset_index()
            xstj_df.append(temp_xstj_df)


    #use merge.
    xstj_results_df = xstj_df.pop()
    for single_df in xstj_df:
        xstj_results_df = pd.merge(xstj_results_df, single_df, how='outer', on=sheet_map['groupby'][0])

    zs_results_df = zs_df.pop()
    for single_df in zs_df:
        zs_results_df = pd.merge(zs_results_df, single_df, how='outer', on=sheet_map['groupby'][0])

    xstj_dict = xstj_results_df.T.to_dict()
    zs_dict = zs_results_df.T.to_dict()

    ret_map = {}
    ret_map['xstj'] = xstj_dict
    ret_map['zs'] = zs_dict

    return ret_map


getData()