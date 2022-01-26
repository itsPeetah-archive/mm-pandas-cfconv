import sys
from datetime import datetime
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
import pandas as pd
import numpy as np
from pandas._libs.tslibs.timestamps import Timestamp

months = {
 1:"Gennaio",
 2:"Febbraio",
 3:"Marzo",
 4:"Aprile",
 5:"Maggio",
 6:"Giugno",
 7:"Luglio",
 8:"Agosto",
 9:"Settembre",
 10:"Ottobre",
 11:"Novembre",
 12:"Dicembre"   
}

def extract_month(ts:Timestamp) -> int:
    m =ts.month
    return m
def extract_year(ts:Timestamp) -> int:
    y = ts.year
    return y
def int_month_to_str(month:int) -> str:
    return months[month]

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

current_year = datetime.now().year
filename_in = ""
filename_out_old = ""
filaname_out_new = ""

processed_final_old, processed_final_new = None, None

def openFileNameDialog(window, label=None):
    global filename_in
    options = QFileDialog.Options()
    options |= QFileDialog.DontUseNativeDialog
    fileName, _ = QFileDialog.getOpenFileName(window,"QFileDialog.getOpenFileName()", "","Excel Files (*.xlsx)", options=options)
    if fileName:
        filename_in = fileName
        if label:
            label.setText(filename_in)

        print("Chose input file:", filename_in)

def saveFileDialog(window):
    options = QFileDialog.Options()
    # options |= QFileDialog.DontUseNativeDialog
    fileName, _ = QFileDialog.getSaveFileName(parent=window,caption="QFileDialog.getSaveFileName()",directory="",filter="Excel Files (*.xlsx)", options=options)
    if fileName:
        print("Chosen file:", fileName)
        if fileName[-5:] != ".xlsx": fileName += ".xlsx"
        return fileName
    return None

def load_and_process_files(label:QLabel):
    global processed_final_old, processed_final_new
    global filename_in

    print("Processing file...")
    try:
        df1, df2 = process_xlsx(filename_in, current_year, months)
    except:
        print("Something went wrong...")
        label.setStyleSheet("background-color: maroon")
        label.setText("Oh no! Qualcosa è andato storto nell'elaborazione!\nControlla di aver scelto il file giusto.")
    else:
        print("Done!")
        processed_final_old, processed_final_new = df1, df2
        label.setStyleSheet("background-color: limeGreen")
        label.setText("Elaborazione avvenuta con successo.")

def save_file(format, window):

    fn = saveFileDialog(window)
    if not fn:
        print("Operation was cancelled...")
        return
    if processed_final_old is None or processed_final_new is None:
        print("Something went wrong...")
        return
    
    df = processed_final_old if format == 0 else processed_final_new
    try:
        render_xlsx(df, fn)
    except:
        print("Couldn't render xlsx file...")

def main():
    app = QApplication([])
    window = QWidget()
    layout = QVBoxLayout()
    layout.setAlignment(Qt.AlignmentFlag.AlignTop)

    # Input
    inputRow = QHBoxLayout()
    inFilenameLabel = QLabel("Nessun file selezionato")
    inFileButton = QPushButton("Scegli file")
    inFileButton.setMaximumWidth(100)
    inFileButton.clicked.connect(lambda: openFileNameDialog(window, inFilenameLabel))

    inputRow.addWidget(inFileButton)
    inputRow.addWidget(inFilenameLabel)

    # Process
    processRow = QVBoxLayout()
    processButton = QPushButton("Elabora file")
    processButton.setFixedHeight(60)
    processButton.setStyleSheet("margin-left: 60px; margin-right:60px")
    processButton.clicked.connect(lambda: load_and_process_files(msgLabel))
    msgLabel = QLabel("")
    
    processRow.addWidget(processButton)
    processRow.addWidget(msgLabel)

    # Output
    outputRow = QHBoxLayout()
    oldFormatButton = QPushButton("Salva file nel formato v1")
    oldFormatButton.clicked.connect(lambda: save_file(0, window))
    newFormatButton = QPushButton("Salva file nel formato v2")
    newFormatButton.clicked.connect(lambda: save_file(1, window))

    outputRow.addWidget(oldFormatButton)
    outputRow.addWidget(newFormatButton)

    layout.addWidget(QLabel("Istruzioni:"))
    layout.addWidget(QLabel("Scegli il file Excel (e.g. cashflow.xlsx) prodotto da VeriERP."))
    layout.addWidget(QLabel("Premi su \"Elabora file\""))
    layout.addWidget(QLabel("Seleziona il formato in cui vuoi esportare la nuova tabella generata:"))
    lab1 = QLabel("v1: colonne mensili solo per il fatturato"); lab1.setStyleSheet("margin-left:20px")
    lab2 = QLabel("v2: colonne mensili per fatturato e incasso"); lab2.setStyleSheet("margin-left:20px")
    layout.addWidget(lab1); layout.addWidget(lab2)
    layout.addWidget(QLabel("L'esportazione avrà successo solo previa elaborazione."))
    layout.addWidget(QLabel("Durante l'esportazione, l'applicazione provvederà ad aggiungere l'estensione \".xlsx\" qualora mancante."))
    layout.addSpacing(50)
    
    layout.addLayout(inputRow)
    layout.addSpacing(100)
    layout.addLayout(processRow)
    layout.addSpacing(100)
    layout.addLayout(outputRow)

    window.setLayout(layout)
    window.setWindowTitle("MM/CF-CONV")
    window.setFixedWidth(800)
    window.setFixedHeight(600)
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()