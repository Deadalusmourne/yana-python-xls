#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Time    : 2018/1/29 
# __author__: caoge

import xlrd

wb = xlrd.open_workbook('test.xls.old')
ws = wb.sheet_by_index(0)
cell = ws.cell(0, 8)
print cell.value, type(cell.value)