import pandas as pd
import numpy as np
from openpyxl import Workbook, load_workbook



#######################################################
#Добавление столбца reopens на исходный (активный) лист
#######################################################
def insert_reopens_coll():
    wb = load_workbook('test.xlsx')
    sheet = wb.active
    sheet.insert_cols(2, 1)
    sheet['B1'] = 'reopens'
    sheet['B2'] = 1
    wb.save('result.xlsx')


##################################################################
#Заполнение столбца reopens значениями, формула смотрит на chat_id
##################################################################
def fill_reopens():
    wb = load_workbook('result.xlsx')
    sheet = wb.active
    max_rows = sheet.max_row
    for i in range(3, max_rows):
        result = np.where(
            sheet[f'A{i}'].value == sheet[f'A{i-1}'].value, sheet[f'B{i-1}'].value + 1, 1) #Формула - =ЕСЛИ(А3=А2;B2+1;1)
        sheet[f'B{i}'] = int(result)
    wb.save('result.xlsx')
    print('reopens done <3')


#####################################################################
#Добавление столбца answered(Есть ответ?) на исходный (активный) лист
#####################################################################
def insert_response_coll():
    wb = load_workbook('result.xlsx')
    sheet = wb.active
    sheet.insert_cols(8, 1)
    sheet['H1'] = 'answered'
    wb.save('result.xlsx')


########################################################################################
#Заполнение столбца answered значениями, формула смотрит на столбец response из выгрузки
########################################################################################
def fill_response():
    wb = load_workbook('result.xlsx')
    sheet = wb.active
    max_rows = sheet.max_row
    for i in range(2, max_rows):
        result = np.where(sheet[f'G{i}'].value == None, 'Нет', 'Да')
        sheet[f'H{i}'] = str(result)
    wb.create_sheet('Сводная')
    wb.save('result.xlsx')
    print('response done <3')


#####################################################################
#Попытка создать сводную через пандас, можно закомментить или удалить
#####################################################################
def create_pivot_table():
    xls_file = pd.read_excel('result.xlsx')
    xls_file[['sure_topic', 'reopens', 'chat_id', 'answered']]
    pivot_table = xls_file.pivot_table(index='sure_topic',
                                        columns='reopens',
                                        values='chat_id',
                                        aggfunc='count')
    pivot_table.to_excel('test2.xlsx', sheet_name='Сводная')


if __name__ == '__main__':
    insert_reopens_coll()
    fill_reopens()
    insert_response_coll()
    fill_response()
    #create_pivot_table()