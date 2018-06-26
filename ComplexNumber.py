import xlrd
import panda
import numpy as np


workbook = xlrd.open_workbook('testfile_new.xlsx')
Voltage_SpreadSheet = workbook.sheet_by_index(0)
Current_SpreadSheet = workbook.sheet_by_index(1)
print(Voltage_SpreadSheet.nrows)
print(Current_SpreadSheet.nrows)
print(Voltage_SpreadSheet.cell_value(0,0))
val1= Voltage_SpreadSheet.cell_value(0,0)
print(type(val1))
a = np.array([val1,Voltage_SpreadSheet.cell_value(1,0)])
b = np.loadtxt(CSV_File.csv).view(complex)