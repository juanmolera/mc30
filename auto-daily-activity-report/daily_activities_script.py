import pandas as pd
from datetime import date
import xlwings as xw
import time

today = date.today()

wb = xw.Book('actividades_desarrolladas_juan_molera.xlsx')

sheet1 = wb.sheets['Sheet1']

sheet1.range(f'A{today.day+3}').value = today.day
sheet1.range(f'B{today.day+3}').value = 'Trabajo de oficina'
sheet1.range(f'C{today.day+3}').value = 'Oficinas Madrid Calle 30'
sheet1.range(f'D{today.day+3}').value = '8h'
sheet1.range(f'E{today.day+3}').value = ''

time.sleep(1)

wb.close()