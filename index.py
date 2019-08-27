# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

#! /usr/bin/python3
import openpyxl

excel = openpyxl.load_workbook('Thongkeviphamchitiet.xlsx')
print(excel.sheetnames)
print(excel.active.title)
excel.save("abcdef.xlsx")
