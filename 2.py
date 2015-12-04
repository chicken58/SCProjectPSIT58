from openpyxl import load_workbook
from openpyxl import Workbook
wb = load_workbook('sample.xlsx')
sh = wb['Sheet']
d = sh['A1'].value
print(d)
