from functions import render_xlsx, process_xlsx
import sys
from datetime import datetime
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QWidget, QApplication, QLabel, QPushButton, QFileDialog, QVBoxLayout, QHBoxLayout

current_year = datetime.now().year
filename_in = ""
filename_out_old = ""
filaname_out_new = ""

processed_final_old, processed_final_new = None, None

months = {
 1:"Gennaio", 2:"Febbraio", 3:"Marzo",
 4:"Aprile", 5:"Maggio", 6:"Giugno",
 7:"Luglio", 8:"Agosto", 9:"Settembre",
 10:"Ottobre", 11:"Novembre", 12:"Dicembre"   
}

def openFileNameDialog(window, label=None):
    global filename_in
    options = QFileDialog.Options()
    # options |= QFileDialog.DontUseNativeDialog
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
    processButton.setStyleSheet("margin-left: 60px; margin-right:60px; font-size:18px")
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