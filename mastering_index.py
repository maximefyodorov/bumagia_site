# -*- coding: utf-8 -*-

import jinja2
import openpyxl

wb = openpyxl.load_workbook('w:/home/bumagia/www/templates/bumagia_data.xlsx', data_only=True)
templateLoader = jinja2.FileSystemLoader(searchpath='w:/home/bumagia/www/templates/All_templates')
templateEnv = jinja2.Environment(loader=templateLoader)

ws = wb['index']
catalog_link = ws.cell(row = 1, column = 2).value
menu_data = []
for i in range (2, ws.max_column):
    menu_item = {}
    for j in range(2,9):
        menu_item[ws.cell(row = j, column = 1).value] = ws.cell(row = j, column = i).value
    menu_data.append(menu_item)

ws = wb['common_data']
page_data = {}
for i in range(2, ws.max_row+1):
    page_data[ws.cell(row = i, column = 1).value] = ws.cell(row = i, column = 2).value

ws = wb[page_data['section_name']]
product_data = []
for i in range (2, ws.max_row+1):
    row = {}
    for j in range (1, ws.max_column+1):
        row[ws.cell(row = 1, column = j).value] = ws.cell(row = i, column = j).value
    if row['main'] ==1:
        product_data.append(row)

tmplt = templateEnv.get_template('index_template.html')
with open('w:/home/bumagia/www/index.html', 'wb') as dest:
    output = tmplt.render(
        catalog_link = catalog_link,
        page_data = page_data,
        menu_data = menu_data,
        product_data = product_data
    )
    dest.write(output.encode('utf-8'))