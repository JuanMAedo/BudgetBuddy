from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string

from utils.excel_styles import thin_border
from utils.functionalities import day_to_name


def create_monthly_column_sheets(
    month_sheet, title, init_column, last_column, text_column, color_title, color_column
):
    # PARSER INIT_COLUMN TO ITERATIONS
    column_letter = init_column[0]
    column_number = column_index_from_string(column_letter)
    row_number = int(init_column[1:])

    # VERTICAL AXIS
    month_sheet[init_column] = f"{title}"
    cell_b2 = month_sheet[init_column]
    cell_b2.alignment = Alignment(horizontal="center")
    cell_b2.fill = PatternFill(
        start_color=color_title, end_color=color_title, fill_type="solid"
    )  # Color morado claro
    cell_b2.font = Font(bold=True, name="Calibri")

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
    
    # Aplicar el borde al rango de celdas seleccionado
    cell_range = month_sheet[init_column:last_column]    
    for row in cell_range:
        for cell in row:
            cell.border = thin_border