import openpyxl
from openpyxl import load_workbook
#https://openpyxl.readthedocs.io/en/stable/tutorial.html#loading-from-a-file
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from openpyxl.styles import colors
from openpyxl.styles import Font, Color
from openpyxl import Workbook
from openpyxl.utils  import get_column_letter
from copy import copy

class appender():
    def __init__(self,
                 filename1,
                 filename2,
                 H1,
                 sheet_name,
                 offset = 4,
                 H1_font_size = 15,
                 finding_report_use = False,
                 font_name="Calibri"
                 ):
        if finding_report_use: ##finding report use
            offset = 2
            H1_font_size = 18
        wb1 = load_workbook(filename1)
        wb2 = load_workbook(filename2)
        ws1 = wb1[sheet_name]
        ws2 = wb2[sheet_name]
        No_row = ws1.max_row + offset
        H1_obj = ws1[get_column_letter(1) + str(No_row - 1)]
        H1_obj.value = H1
        if finding_report_use:
            fill_color_header = '92D050'
            H1_obj.font = Font(name = font_name, size=H1_font_size, bold=True)
            fill = PatternFill(start_color=fill_color_header, end_color=fill_color_header, fill_type="solid")
            H1_obj.fill = fill
            ws1.merge_cells(f'{get_column_letter(ws1.min_column)}{str(No_row - 1)}:{get_column_letter(ws1.max_column)}{str(No_row - 1)}')

        else:
            H1_obj.font = Font(name = font_name,size=H1_font_size, bold=True, color='95B3DF')
        for tr in ws2:
            No_col = 1
            for td in tr:
                ws1[get_column_letter(No_col) + str(No_row)].value = td.value
                ws1[get_column_letter(No_col) + str(No_row)].fill = copy(td.fill)
                ws1[get_column_letter(No_col) + str(No_row)].alignment = copy(td.alignment)
                ws1[get_column_letter(No_col) + str(No_row)].font = copy(td.font)
                ws1[get_column_letter(No_col) + str(No_row)].border = copy(td.border)
                ws1[get_column_letter(No_col) + str(No_row)].alignment = copy(td.alignment)

                No_col = No_col + 1
            No_row = No_row + 1
        wb1.save(filename1)

#appender(filename1='temp0.xlsx',filename2 = 'temp2.xlsx',H1='Interal Audit',sheet_name='Issue by Source',finding_report_use = True )