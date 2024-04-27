import datetime
import os
from ctypes import alignment

from dotenv import load_dotenv
from functionalities import *
from functionalities.border_styles import *
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import column_index_from_string

# SECRETS LOADINGS AT VARIABLES
load_dotenv()
input_file_path = os.getenv("INCOME_EXPENSE_RECORD")
monthly_income_expense_report = os.getenv("MONTHLY_REPORT_NAME")
table_start_colum = os.getenv("START_COLUMN")
table_finish_column = os.getenv("FINISH_COLUMN")


def create_monthly_column_sheets(
    month_sheet, title, init_column, text_column, color_title, color_column
):
    # PARSER INIT_COLUMN TO ITERATIONS
    column_letter = init_column[0]
    column_number = column_index_from_string(column_letter)
    row_number = int(init_column[1:])

    # VERTICAL AXIS
    month_sheet[init_column] = f"{title}"
    cell_B2 = month_sheet[init_column]
    cell_B2.alignment = Alignment(horizontal="center")
    cell_B2.fill = PatternFill(
        start_color=color_title, end_color=color_title, fill_type="solid"
    )  # Color morado claro
    cell_B2.font = Font(bold=True, name="Calibri")

    for i in range(5):
        current_row = row_number + i
        cell = month_sheet.cell(row=current_row + 1, column=column_number)
        week_text = f"{text_column} {i + 1}"

        cell.value = week_text
        cell.alignment = Alignment(horizontal="center")
        cell.fill = PatternFill(
            start_color=color_column, end_color=color_column, fill_type="solid"
        )  # Color naranja apagado
        cell.font = Font(bold=True, name="Calibri")

    # HORIZONTAL AXIS
    for day_number, day_name in day_to_name.items():
        current_column = column_number
        cell = month_sheet.cell(row=row_number, column=current_column + 1 + day_number)
        day_text = f"{day_name}"
        cell.value = day_text
        cell.alignment = Alignment(horizontal="center")
        cell.fill = PatternFill(
            start_color="528EF6", end_color="528EF6", fill_type="solid"
        )  # Color azul pastel
        cell.font = Font(bold=True, name="Calibri")

        cell_range = month_sheet[init_column:"I7"]  # Por ejemplo, del A1 al D5
    # Aplicar el borde al rango de celdas seleccionado
    for row in cell_range:
        for cell in row:
            cell.border = thin_border
        # Aplicar el borde grueso a los lados exteriores del rango

    for row_idx, row in enumerate(cell_range, 1):
        for col_idx, cell in enumerate(row, 1):
            if row_idx == 1:  # Borde superior
                cell.border = Border(
                    left=Side(style="thin"),
                    right=Side(style="thin"),
                    top=Side(style="thick"),
                    bottom=Side(style="thin"),
                )
            if row_idx == len(cell_range):  # Borde inferior
                cell.border = Border(
                    left=Side(style="thin"),
                    right=Side(style="thin"),
                    top=Side(style="thin"),
                    bottom=Side(style="thick"),
                )
            if col_idx == 1:  # Borde izquierdo
                cell.border = Border(
                    left=Side(style="thick"),
                    right=Side(style="thin"),
                    top=Side(style="thin"),
                    bottom=Side(style="thin"),
                )
            if col_idx == len(row):  # Borde derecho
                cell.border = Border(
                    left=Side(style="thin"),
                    right=Side(style="thick"),
                    top=Side(style="thin"),
                    bottom=Side(style="thin"),
                )


# VERIFY INPUT FILE
if os.path.exists(input_file_path) and os.path.splitext(input_file_path)[1] in (".xls", ".xlsx"):
    print("File it's correct")
else:
    print("The pathing '" "{}" "' or the file type are incorrect.".format(input_file_path))

# OPEN AND READ THE INPUT EXCEL
expense_income_read = load_workbook(input_file_path)
excel_date = expense_income_read.active[table_start_colum].value
# Verify the date
if isinstance(excel_date, datetime.datetime):
    # Parses the date
    year = excel_date.year
    month = excel_date.month
    day = excel_date.day

# print(excel_date.weekday()) # 0 es Lunes, 6 es domingo --> Para el parseado semanal de las hojas
# print(excel_date.month)


# CREATE OR COMPLETE THE ANNUAL MONTHLY I&E REPORT
file_name = f"{monthly_income_expense_report} {year}.xlsx"
print(file_name)
if os.path.exists(file_name):
    inc_exp_excel = load_workbook(file_name)
    hoja = inc_exp_excel.active

    # ...

else:
    inc_exp_excel = Workbook()
    inc_exp_excel.create_sheet("DASHBOARD")
    inc_exp_excel.create_sheet("CATEGORY")
    # CREATE MONTHLY SHEETS
    for month_number, month_name in month_to_name.items():
        month_sheet = inc_exp_excel.create_sheet(month_name)
        inc_exp_excel.active = month_sheet
        create_monthly_column_sheets(month_sheet, "INCOME", "B10", "Week", "24D124", "D78740")
        create_monthly_column_sheets(month_sheet, "EXPENSE", "B2", "Week", "C06FCA", "D78740")

    inc_exp_excel.remove(inc_exp_excel["Sheet"])
    # ...

inc_exp_excel.save(file_name)
