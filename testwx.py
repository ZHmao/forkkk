# -*- coding: utf-8 -*-

import wx
from wx import grid as wxgrid
import getDisplayData

'''
author mzh
since 2015-4-28
'''

class Main_Window(wx.Frame):

    def __init__(self, *args, **kwargs):
        super(Main_Window, self).__init__(*args, **kwargs)
        self.error_msg = ''
        self.sale_data = getDisplayData.getData()
        self.InitUI()

    def InitUI(self):

        h_box = wx.BoxSizer(wx.HORIZONTAL)
        panel = wx.Panel(self, -1)

        self.first_grid = wxgrid.Grid(panel)

        self.put_data_in_grid(self.sale_data)

        h_box.Add(self.first_grid, 1, wx.EXPAND)
        panel.SetSizer(h_box)

        self.Center()
        self.Show()

    def put_data_in_grid(self, data=None):
        data_map = {'mao': { '1': 23.4, '2': 12, '3': 43}, 'dage': {'1': 454, '2': 5.32, '3': 0.23}}
        if data is None:
            self.error_msg += 'oh, there is no data in your specified file'
            return

        col_names = data.keys()
        row_names = data[col_names[0]].keys()
        row_num = len(row_names)
        col_num = len(col_names)

        self.first_grid.CreateGrid(row_num, col_num)

        for current_row in range(row_num):
            for current_col in range(col_num):
                cell_value = data[col_names[current_col]][row_names[current_row]]
                """因为grid只接收String和Unicode，所以对其他类型进行转换 """
                if type(cell_value) is not unicode:
                    cell_value = str(cell_value)
                self.first_grid.SetCellValue(current_row, current_col, cell_value)

def main():
    app = wx.App()
    Main_Window(None, title='main window', size=(1000, 600))
    app.MainLoop()

def test():
    app = wx.App()
    window = wx.Frame(None, -1, 'main window', size=(1000, 600))
    window.Center()
    window.Show()
    app.MainLoop()

main()
