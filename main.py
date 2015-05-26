# -*- coding: utf-8 -*-

import wx
from wx import grid as wxgrid
import exceldata

'''
mzh
2015-4-28
'''


class MainWindow(wx.Frame):

    def __init__(self, *args, **kwargs):
        super(MainWindow, self).__init__(*args, **kwargs)

        self.error_msg = ''
        self.data = exceldata.get_data()
        self.initUI()

    def initUI(self):
        h_box = wx.BoxSizer(wx.HORIZONTAL)
        panel = wx.Panel(self, -1)

        '''create notebook'''
        self.notebook = wx.Notebook(panel, style=wx.NB_LEFT, size=(1000, 600))

        '''add a scroll panel'''
        self.sales_page = wx.ScrolledWindow(self.notebook, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.HSCROLL | wx.VSCROLL)
        self.payment_page = wx.ScrolledWindow(self.notebook, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.HSCROLL | wx.VSCROLL)
        self.sales_grid = wxgrid.Grid(self.sales_page)
        self.payment_grid = wxgrid.Grid(self.payment_page)
        self.notebook.AddPage(self.sales_page, u"销售统计", select=True)
        self.notebook.AddPage(self.payment_page, u"中收")
        self.put_data_in_grid(self.sales_grid, self.data.get('sales'), self.data.get('sales_column'))

        h_box.Add(self.sales_grid, 1, wx.EXPAND)
        panel.SetSizer(h_box)

        self.Center()
        self.Show()

    def put_data_in_grid(self, target_grid=None, grid_data=None, col_name_list=None):
        if grid_data is None or col_name_list is None:
            self.error_msg += 'oh, there is no data in your specified file'
            return
        col_count = len(col_name_list)
        row_count = len(grid_data)

        target_grid.CreateGrid(row_count, col_count)
        '''设置列名'''
        for i in xrange(col_count):
            target_grid.SetColLabelValue(i, col_name_list[i])

        '''填充表体'''
        row_index = 0
        for row_dict in grid_data.values():
            for i in xrange(col_count):
                cell_value = row_dict.get(col_name_list[i])
                '''因为grid只接收String和Unicode，所以对其他类型进行转换'''
                if type(cell_value) is not unicode:
                    cell_value = str(cell_value)
                target_grid.SetCellValue(row_index, i, cell_value)
            row_index += 1

if __name__ == '__main__':
    app = wx.App()
    MainWindow(None, title='main window', size=(1000, 600))
    app.MainLoop()