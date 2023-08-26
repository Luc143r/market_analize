import win32com.client as win32
import sys
import os
import itertools
from pathlib import Path
import time


win32c = win32.constants
f_path = Path.cwd()
f_name = 'result.xlsx'

def pivot_table(wb: object, ws1: object, pt_ws: object, ws_name: str, pt_name: str, pt_rows: list, pt_cols: list, pt_filters: list, pt_fields: list):
    """
    wb = workbook1 reference
    ws1 = worksheet1
    pt_ws = pivot table worksheet number
    ws_name = pivot table worksheet name
    pt_name = name given to pivot table
    pt_rows, pt_cols, pt_filters, pt_fields: values selected for filling the pivot tables
    """

    # pivot table location
    pt_loc = len(pt_filters) + 2
    
    # grab the pivot table source data
    pc = wb.PivotCaches().Create(SourceType=win32c.xlDatabase, SourceData=ws1.UsedRange, Version=win32c.xlPivotTableVersion14)
    # create the pivot table object
    pc.CreatePivotTable(TableDestination=f'{ws_name}!R4C1', TableName=pt_name)

    # selecte the pivot table work sheet and location to create the pivot table
    pt_ws.Select()
    pt_ws.Cells(pt_loc, 1).Select()

    # Sets the rows, columns and filters of the pivot table

    for field_list, field_r in ((pt_filters, win32c.xlPageField), (pt_rows, win32c.xlRowField), (pt_cols, win32c.xlColumnField)):
        for i, value in enumerate(field_list):
            pt_ws.PivotTables(pt_name).PivotFields(value).Orientation = field_r
            pt_ws.PivotTables(pt_name).PivotFields(value).Position = i + 1

    # Sets the Values of the pivot table
    for field in pt_fields:
        pt_ws.PivotTables(pt_name).AddDataField(pt_ws.PivotTables(pt_name).PivotFields(field[0]), field[1], field[2]).NumberFormat = field[3]

    # Visiblity True or Valse
    pt_ws.PivotTables(pt_name).ShowValuesRow = True
    pt_ws.PivotTables(pt_name).ColumnGrand = True


def run_excel(f_path: Path, f_name: str):
    
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True
    filename = f_path / f_name
    #path_file = os.path.abspath('result.xlsx')
    wb = excel.Workbooks.Open(filename)
    ws1 = wb.Sheets('test.csv')
    ws2_name = 'pivot'
    wb.Sheets.Add().Name = ws2_name
    ws2 = wb.Sheets('pivot')

    pt_name = 'example'
    pt_rows = ['sure_topic']
    pt_cols = ['iterations_new']
    pt_filters = ['chat_id']
    pt_fields = [['Есть ответ', 'Есть ответ: sum', win32c.xlSum, '$#,##0.00'], 
                ['Нет ответа', 'Нет ответа: sum', win32c.xlSum, '$#,##0.00']]

    pivot_table(wb, ws1, ws2, ws2_name, pt_name, pt_rows, pt_cols, pt_filters, pt_fields)


run_excel(f_path, f_name)
