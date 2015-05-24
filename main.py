# -*- coding: utf-8 -*-

import wx
from wx import grid as wxgrid
import exceldata

'''
mzh
2015-4-28
'''

class Main_Window(wx.Frame):

    def __init__(self, *args, **kwargs):
        super(Main_Window, self).__init__(*args, **kwargs)

        self.error_msg = ''
        self.data = exceldata.get_data()
        self.initUI()

    def initUI(self):
        h_box = wx.BoxSizer(wx.HORIZONTAL)
        panel = wx.Panel(self, -1)

        self.sales__grid = wxgrid.Grid(panel)

        self.put_data_in_grid(self.sale_data)

        h_box.Add(self.sales__grid, 1, wx.EXPAND)
        panel.SetSizer(h_box)

        self.Center()
        self.Show()

    def put_data_in_grid(self, grid_data=None, col_name_list=None):
        if grid_data is None or col_name_list is None:
            self.error_msg += 'oh, there is no data in your specified file'
            return
        col_count = len(col_name_list)
        row_count = len(grid_data)

        self.sales__grid.CreateGrid(row_count, col_count)
