#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Time    : 2018/1/29 
# __author__: caoge

import xlrd

wb = xlrd.open_workbook('test.xls.old')
ws = wb.sheet_by_index(0)
cell = ws.cell(0, 8)
print cell.value, type(cell.value)

class Get_Some(object):
    def __init__(self):
        pass

    def _get_sh_(self):
        print 'ok'

    def get_ni(self):
        print 'ok2'


def _get_ki():
    print 'ok3'





if __name__ == '__main__':
    pass
