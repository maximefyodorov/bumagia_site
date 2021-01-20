# -*- coding: utf-8 -*-

import jinja2
import openpyxl

wb = openpyxl.load_workbook('w:/home/bumagia/www/templates/bumagia_data.xlsx', data_only=True)

products_list = wb.sheetnames[1:]
number_of_products = len(products_list)

ws = wb['common_data']

products_data = {}
my_products = []
my_menuitem_name =[]

for i in range (number_of_products):
    if ws.cell(row = 1, column = i+2).value == 1:
        my_products.append(ws.cell(row = 2, column = i+2).value)
        temp_dict = {}
        for j in range (2, ws.max_row+1):
            temp_dict[ws.cell(row = j, column = 1).value] = ws.cell(row = j, column = i+2).value
        products_data[products_list[i]] = temp_dict

for i in range (2, ws.max_row+1):
    my_menuitem_name.append(ws.cell(row = 8, column = i).value)

for my_product in (my_products):

    ws = wb[my_product]


    my_control = []
    my_arr_name = []
    my_arr_articul = []
    my_arr_full_articul = []
    my_arr_mame_and_articul = []
    my_arr_buylink = []
    my_ext = []

    my_len = 0

    for i in range(1, ws.max_row + 1):
        # Управляющая колонка 1 (0 - игнорировать, 1 - выводить в общий список, 2 - выводить на витрину на главной)
        control = ws.cell(row = i, column = 1).value 
        if (control != 0):
            my_len+=1
            # Собираем ненулевые управляющие символы
            my_control.append(ws.cell(row = i, column = 1).value)
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
            # Расширение файла (колонка 7)
            my_ext.append(ws.cell(row = i, column = 7).value)


    # print(my_arr_name)

    templateLoader = jinja2.FileSystemLoader(searchpath='w:/home/bumagia/www/templates/Inner_templates/')
    templateEnv = jinja2.Environment(loader=templateLoader)
    tmplt = templateEnv.get_template('{}_template.html'.format(my_product))
    with open('w:/home/bumagia/www/products/{}.html'.format(my_product), 'wb') as dest:
        output = tmplt.render(
            p_list = products_list,
            p_num = number_of_products,
            menuname_list = my_menuitem_name,
            prefix = '',
            p_data = products_data,
            i_num = my_len,
            control = my_control,
            c_prod = my_product, 
            name_list = my_arr_name,
            articul_list = my_arr_articul,
            full_articul_list = my_arr_full_articul,
            fullnameandarticul_list = my_arr_mame_and_articul,
            buylink_list = my_arr_buylink,
            ext = my_ext
        )
        dest.write(output.encode('utf-8'))
