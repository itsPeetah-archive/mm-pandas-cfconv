import calendar
from util.calendar import months
from data_handling.core import process_xlsx, render_xlsx

filename = "cashflow.xlsx"
current_year = 2022

df1, df2 = process_xlsx(filename, current_year, months)
render_xlsx(df1, "cashflow_2022_oldf.xlsx")
render_xlsx(df2, "cashflow_2022_newf.xlsx")
