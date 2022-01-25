import pandas as pd
from data_handling.functions import *

def process_xlsx(filename:str,current_year:int,month_labels:dict[int:str])->tuple[pd.DataFrame,pd.DataFrame]:
    src = load_src_xlsx(filename)
    df = split_month_year(src)
    df = add_next_prev_year_cols(df,current_year)
    df = add_monthly_cols(df,month_labels,current_year)
    agg = sum_by_code(df)
    final1 = to_old_format(agg)
    aggNew = adapt_monthly_cols_to_new_format(agg,month_labels)
    final2 = to_new_format(aggNew)
    return final1, final2

def render_xlsx(df:pd.DataFrame, filename:str):
    df.to_excel(filename)