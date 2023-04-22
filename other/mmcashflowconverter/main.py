import pandas as pd
import util

# General settings
current_year = 2022
na_fill = "-"

# Source Excel sheet settings
src_name = "cashflow.xlsx" # name of the excel file
sheet_to_read = 0 # 0-index of the spreadsheet to read from the excel file
src_code_column = "offerte e commesse::codice_commessa"
src_date_column = "data fatturazione prevista"
src_payment_column = "importo acconto fatturato"

# Destination Excel sheet settings
dst_name = f"cashflow_{current_year}.xlsx"
dst_sheet = "cashflow"
dst_prev_year_column = "Anni Precedenti"
dst_next_year_column = "Anni Successivi"

# Other settings (DO NOT CHANGE)
code_col = "code"
date_col = "date"
pay_col = "pay"
prev_col = dst_prev_year_column
next_col = dst_next_year_column
dst_columns = [dst_prev_year_column, "Gennaio","Febbraio","Marzo",
                "Aprile","Maggio","Giugno","Luglio","Agosto", "Settembre", "Ottobre",
                "Novembre", "Dicembre", dst_next_year_column]

# Read data from excel spreadsheet
df = pd.read_excel(src_name, sheet_to_read)
df = df[[src_code_column, src_date_column, src_payment_column]]
df = df.rename(columns={src_code_column : code_col, 
                        src_date_column : date_col, 
                        src_payment_column : pay_col})

codes = df[code_col].unique()
prev_year = df[df[date_col].apply(util.extract_year) < current_year].groupby(code_col).sum().fillna(na_fill)
next_year = df[df[date_col].apply(util.extract_year) > current_year].groupby(code_col).sum().fillna(na_fill)
curr_year = df[df[date_col].apply(util.extract_year) == current_year]
curr_year[date_col] = curr_year[date_col].apply(util.extract_month).apply(util.int_month_to_str)
curr_year =curr_year.groupby([code_col, date_col]).sum()

final = pd.DataFrame(index = codes, columns=dst_columns)
final[prev_col] = prev_year[pay_col]
final[next_col] = next_year[pay_col]

for cy_index, cy_row in curr_year.iterrows():
    code = cy_index[0]
    month = cy_index[1]
    pay = cy_row["pay"]
    final.at [code, month] = pay

final.to_excel(dst_name, dst_sheet)
