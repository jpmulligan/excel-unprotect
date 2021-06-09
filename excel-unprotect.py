# -*- coding: utf-8 -*-
"""
Created on Tue Jun  8 21:12:11 2021

@author: 
"""

from openpyxl import load_workbook
import time

starttime = time.time()

root_folder = '.'

filename = 'example.xlsx'

wb = load_workbook(f'{root_folder}/{filename}')

for sheet in wb:
    ws = sheet

    print(f'\nWorksheet Title: {ws.title}')
    print(f'...Sheet Protection Status = {ws.protection.sheet}')
    if ws.protection.sheet == True:
        ws.protection.disable()  
        print('...protection disabled\n')

wb.save(f'{root_folder}/{filename}')      

wb.close()

endtime = time.time()
elapsed = round(endtime - starttime, 2) #seconds

print(f'\nElapsed run time: {elapsed} seconds')


        





