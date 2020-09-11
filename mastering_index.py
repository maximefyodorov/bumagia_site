# -*- coding: utf-8 -*-

from jinja2 import Template
import openpyxl

wb = openpyxl.load_workbook('w:/home/bumagia/www/templates/bumagia_data.xlsx', data_only=True)

products_list = wb.sheetnames[1:]
number_of_products = len(products_list)

ws = wb['common_data']

products_data = {}

for i in range (number_of_products):
    temp_dict = {}
    for j in range (2, ws.max_row+1):
        temp_dict[ws.cell(row = j, column = 1).value] = ws.cell(row = j, column = i+2).value
    products_data[products_list[i]] = temp_dict


my_product = 'ecofiller'

ws = wb[my_product]

my_arr_name = []
my_arr_articul = []
my_arr_full_articul = []
my_arr_mame_and_articul = []
my_arr_buylink = []

my_len = 0

for i in range(1, ws.max_row + 1):
    # Управляющая колонка 1 (0 - игнорировать, 1 - выводить в общий список, 2 - выводить на витрину на главной)
    control = ws.cell(row = i, column = 1).value 
    if (control == 2):
        my_len+=1
        # Название (колонка 2)
        my_arr_name.append(ws.cell(row = i, column = 2).value)
        # Артикул без буквенной части - последние четыре цифры штрихкода (колонка 3)
        my_arr_articul.append(str(ws.cell(row = i, column = 3).value)[9:])
        # Полный артикул, индекс и цифры (колонка 4)
        my_arr_full_articul.append(ws.cell(row = i, column = 4).value)
        # Имя и полный артикул вместе (колонка 5)
        my_arr_mame_and_articul.append(ws.cell(row = i, column = 5).value)
        # Ссылка на покупку (колонка 6)
        my_arr_buylink.append(ws.cell(row = i, column = 6).value)

html = open('w:/home/bumagia/www/templates/index_template.html', encoding='utf-8').read()
tmplt = Template(html)
with open('w:/home/bumagia/www/index.html', 'wb') as dest:
    output = tmplt.render(
        p_list = products_list,
        p_data = products_data,
        p_num = number_of_products,
        c_prod = my_product, 
        v_num = my_len,
        articul_list = my_arr_articul,
        name_list = my_arr_name,
        full_articul_list = my_arr_full_articul,
        fullnameandarticul_list = my_arr_mame_and_articul,
        buylink_list = my_arr_buylink
    )
    dest.write(output.encode('utf-8'))
