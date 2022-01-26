import pandas as pd
import numpy as np

def rename_cols(df:pd.DataFrame, colsToRename:dict[str:str]) -> pd.DataFrame:
    df = df.copy()
    df = df.rename(columns=colsToRename)
    return df

def reorder_cols(df:pd.DataFrame, newColOrder:list[int]) -> pd.DataFrame:
    df = df.copy()
    df = df.iloc[:,np.array(newColOrder)]
    return df

def load_src_xlsx(filename:str) -> pd.DataFrame:
    src = pd.read_excel(filename, parse_dates=[1], engine="openpyxl", names=["code","date","amount"])
    return src

def split_month_year(df:pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["m"], df["y"] = df["date"].dt.month, df["date"].dt.year
    return df

def add_next_prev_year_cols(df:pd.DataFrame, curr_year:int) -> pd.DataFrame:
    df = df.copy()
    df["prev"] = np.where(df["y"] < curr_year, df["amount"], 0)
    df["next"] = np.where(df["y"] > curr_year, df["amount"], 0)
    return df

def add_monthly_cols(df:pd.DataFrame, month_labels:dict[int:str], curr_year:int) -> pd.DataFrame:
    df = df.copy()
    for m_int, m_label in month_labels.items():
        df[m_label] = np.where((df["m"] == m_int) & (df["y"] == curr_year), df["amount"], 0)
    return df

def sum_by_code(df:pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df = df.groupby(df["code"]).sum().sort_index()
    df = df.drop("y", axis=1)
    df = df.drop("m", axis=1)
    df = rename_cols(df, {"prev":"Anni Precedenti", "next":"Anni Successivi", "amount":"Totale Fatturato"})
    return df

def adapt_monthly_cols_to_new_format(df:pd.DataFrame, month_labels:dict[int:str], oLabel:str="Fatturato", nLabel:str="Incasso") -> pd.DataFrame:
    df = df.copy()
    newMonthCols = {}
    for _, m in month_labels.items():
        newMonthCols[m] = f"{m} {oLabel}"
        df[f"{m} {nLabel}"] = 0
    df[f"Totale {nLabel}"] = 0
    df = rename_cols(df, newMonthCols)
    return df

def to_old_format(df:pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    newColOrder = [1] + [m + 3 for m in range(0, 12)] + [2, 0]
    df.index.rename("Codice Commessa", inplace=True)
    df = reorder_cols(df, newColOrder)
    return df

def to_new_format(df:pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    newColOrder = [1] + np.array([(m + 3, m + 3 + 12) for m in range(12)]).flatten().tolist() + [2, 0, -1]
    df.index.rename("Codice Commessa", inplace=True)
    df = reorder_cols(df, newColOrder)
    return df