import datetime
import os

from dotenv import load_dotenv
from openpyxl import load_workbook

from annual_reportting.excel_generator import generate_excel_report

# Load Data from the file .env
load_dotenv()
input_file_path = os.getenv("INCOME_EXPENSE_RECORD")
monthly_income_expense_report = os.getenv("MONTHLY_REPORT_NAME")
table_start_column = os.getenv("START_COLUMN")
table_finish_column = os.getenv("FINISH_COLUMN")

# Load Input Excel and generate report
if os.path.exists(input_file_path) and os.path.splitext(input_file_path)[1] in (".xls", ".xlsx"):
    print("File is correct")

    expense_income_read = load_workbook(input_file_path)
    excel_date = expense_income_read.active[table_start_column].value
    if isinstance(excel_date, datetime.datetime):
        # Parse the date to get the year
        year = excel_date.year
        month = excel_date.month
        day = excel_date.day
        generate_excel_report(expense_income_read, year)
    else:
        print("The date in the Excel file is not valid.")
else:
    print("The pathing '" "{}" "' or the file type are incorrect.".format(input_file_path))
