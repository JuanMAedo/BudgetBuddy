from openpyxl import load_workbook


def modify_annual_report(file_name):
    inc_exp_excel = load_workbook(file_name)
    # Lógica para modificar la pestaña anual
    inc_exp_excel.save(file_name)
