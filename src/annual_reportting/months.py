from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import column_index_from_string

from utils.excel_styles import thin_border
from utils.functionalities import day_to_name


def create_monthly_column_sheets(
    month_sheet, title, init_column, last_column, text_column, color_title, color_column
):
    # Parse init_column to get the starting cell coordinates
    column_number = column_index_from_string(init_column[0])
    row_number = int(init_column[1:])

    # Lateral Header Creation
    month_sheet[init_column] = f"{title}"
    title_cell = month_sheet[init_column]
    title_cell.alignment = Alignment(horizontal="center")
    title_cell.font = Font(bold=True, name="Calibri")
    title_cell.fill = PatternFill(
        start_color=color_title, end_color=color_title, fill_type="solid"
    )
    # Horizontal Header Creation
    for day_number, day_name in day_to_name.items():
        cell = month_sheet.cell(row=row_number, column=column_number + 1 + day_number)
        cell.value = f"{day_name}"
        cell.alignment = Alignment(horizontal="center")
        cell.fill = PatternFill(start_color="528EF6", end_color="528EF6", fill_type="solid")
        cell.font = Font(bold=True, name="Calibri")

    for i in range(5):
        cell = month_sheet.cell(row=row_number + i + 1, column=column_number)
        cell.value = f"{text_column} {i + 1}"
        cell.alignment = Alignment(horizontal="center")
        cell.fill = PatternFill(
            start_color=color_column, end_color=color_column, fill_type="solid"
        )
        cell.font = Font(bold=True, name="Calibri")

    # Border to the tables
    for row in month_sheet[init_column:last_column]:
        for cell in row:
            cell.border = thin_border
