import datetime

# Diccionary Number Month with Name Month to use it at Excel Sheets
month_to_name = {
    1: "JAN",
    2: "FEB",
    3: "MAR",
    4: "APR",
    5: "MAY",
    6: "JUN",
    7: "JUL",
    8: "AGT",
    9: "SEP",
    10: "OCT",
    11: "NOV",
    12: "DIC",
}
# Diccionary Number Week Day with Name Day to use it at Excel Sheets
day_to_name = {
    0: "Monday",
    1: "Tuesday",
    2: "Wendesday",
    3: "Thursday",
    4: "Friday",
    5: "Saturaday",
    6: "Sunday",
}


def year_from_excel_date(excel_date):
    # Verifica si la fecha es un objeto datetime
    if isinstance(excel_date, datetime.datetime):
        # Retorna el año de la fecha
        return excel_date.year
    else:
        # Si no es un objeto datetime, devuelve None o maneja el error según lo desees
        return None
