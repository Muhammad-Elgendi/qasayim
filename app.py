import sys
from PyQt5.QtWidgets import *
from PyQt5.QtGui import QIcon, QIntValidator
from PyQt5.QtCore import pyqtSlot
from PyQt5 import QtCore, QtGui, QtPrintSupport
import pandas as pd
import json
from PandasView import DataFrameModel
from functools import partial
from datetime import datetime
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from arabic_reshaper import ArabicReshaper
from bidi.algorithm import get_display

# Read template file that has rules
template = pd.read_excel('qasaym_template.xlsx')

# new dataframe to save the data to
new_df = pd.DataFrame(columns=template.columns)

# keep up with the latest selected row
last_selected_row = None

# Add/Save dialog
def dialog():
    msgBox = QMessageBox()
    msgBox.setWindowTitle("Qasaym")
    msgBox.setText("تم إضافة قسيمة جديدة")
    msgBox.setInformativeText("هل تريد حفظ التعديلات؟")
    msgBox.setStandardButtons(QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel)
    msgBox.setDefaultButton(QMessageBox.Save)
    ret = msgBox.exec_()

def add_new_qasema(number,name,status,widget):
    record = template[ template['الحاله'] == status ].copy().reset_index(drop=True)
    auto_sum = True
    for key in record.iloc[0][record.iloc[0].str.contains('دخل', na=False)].to_dict().keys():
        if key == 'الجمله':
            auto_sum = False
        value, done = QInputDialog.getDouble(
			widget, 'Input Dialog', str(record[key].iloc[0])+' في '+key+' :')
        if done:
            record[key] = value
        else:
            return None

    record['الاسـم'] = name
    record['رقم القسيمة'] = number   
    record['التاريخ'] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    if auto_sum:
        # convert to float
        numric_cols = list(record.columns)
        numric_cols.remove('الاسـم')
        numric_cols.remove('رقم القسيمة')
        numric_cols.remove('التاريخ')
        numric_cols.remove('الحاله')
        numric_cols.remove('الجمله')        
        for col in numric_cols:
            record[col] = record[col].astype(float)
        record['الجمله'] = record[numric_cols].sum(axis=1)

    global new_df
    new_df = pd.concat([new_df, record], ignore_index = True)
    new_df = new_df.reset_index(drop=True)
    return new_df


def add_click(widget):
    global nameTextEdit
    global recTextEdit
    name = nameTextEdit.text().strip()
    rec = recTextEdit.text().strip()
    status = statusComboBox.currentText()
    global new_df
    global table
    if name and rec:
        new_df = add_new_qasema(rec,name,status,widget)
        model = DataFrameModel(new_df)
        table.setModel(model)
        nameTextEdit.setText('')
        recTextEdit.setText(str(int(rec)+1))



def clicked_table():
    global table
    index = table.selectedIndexes()[0]
    data = table.model().data(index)
    # print ("data : " + str(data))
    # print ("row : " , index.row())
    # print ("col : " , index.column())
    global last_selected_row
    last_selected_row = index.row()

def remove_dialog():
    global last_selected_row

    if last_selected_row == None:
        return None

    global new_df
    number = new_df.iloc[last_selected_row]['رقم القسيمة']
    name  = new_df.iloc[last_selected_row]['الاسـم']
    status = new_df.iloc[last_selected_row]['الحاله']
    msgBox = QMessageBox()
    msgBox.setWindowTitle("Qasaym")
    msgBox.setText("القسيمة رقم "+number+"\n"+" باسم "+name+"\n"+ " نوع القسيمة "+status)
    msgBox.setInformativeText("هل تريد حذف القسيمة؟")
    msgBox.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
    msgBox.setDefaultButton(QMessageBox.Cancel)
    ret = msgBox.exec_()
 
    global table
    if ret == QMessageBox.Ok:
        new_df = new_df.drop([last_selected_row])
        new_df = new_df.reset_index(drop=True)
        model = DataFrameModel(new_df)
        table.setModel(model)
        last_selected_row = None

def save_excel(file_name,default_name = True):
    if default_name:
        file_name = datetime.now().strftime("%d-%m-%Y__%H-%M-%S")+'.xlsx'
    new_df.drop(['التاريخ'],axis = 1).to_excel(file_name, sheet_name='Qasaym 1', index=False, engine='openpyxl')    
    show_saved_message(file_name)

def save_pdf(file_name,default_name = True):
    if default_name:
        file_name = datetime.now().strftime("%d-%m-%Y__%H-%M-%S")+'.pdf'
    df = new_df.drop(['التاريخ'],axis = 1)

    # show arabic header correctly,
    # cuz, matplotlib doesn't support it well
    new_arabic_cols = [arabify(col) for col in df.columns]
    # plot dataframe in matplotlib
    fig, ax =plt.subplots(figsize=(12,4))
    ax.axis('tight')
    ax.axis('off')
    the_table = ax.table(cellText=df.values,colLabels=new_arabic_cols,loc='center') 
    # the_table.auto_set_font_size(False)   
    # the_table.set_fontsize(24)
    the_table.scale(1.5, 1.5)  # may help

    # Export Matplotlib table to PDF
    pp = PdfPages(file_name)
    pp.savefig(fig, bbox_inches='tight')
    pp.close()

    show_saved_message(file_name)

def show_saved_message(file_name):
    mbox = QMessageBox()
    mbox.setText("تم حفظ القسائم في ملف \n"+file_name)
    mbox.setStandardButtons(QMessageBox.Ok)
    mbox.setDefaultButton(QMessageBox.Ok)
    # mbox.setDetailedText("تم حفظ المعلومات الخاصة بالقسائم")
    # mbox.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
    mbox.exec_()

def saveFileDialog(widget,file_type):
    if len(new_df) == 0:
        return None
    
    options = QFileDialog.Options()
    options |= QFileDialog.DontUseNativeDialog

    fileName, _ = QFileDialog.getSaveFileName(widget,"QFileDialog.getSaveFileName()","","ALL FILES (*)", options=options)
    
    if fileName:
        # export file based on its type
        if file_type == 'excel':
            save_excel(fileName+".xls",False)
        elif file_type == 'pdf':
            save_pdf(fileName+".pdf",False)

def print_widget(widget):    
    previewDialog = QtPrintSupport.QPrintPreviewDialog()
    previewDialog.paintRequested.connect(handlePaintRequest)  
    previewDialog.exec_()

def handlePaintRequest(printer):
    document = QtGui.QTextDocument()
    cursor = QtGui.QTextCursor(document)
    global table
    global new_df
    print_table = cursor.insertTable(
        table.model().rowCount(), table.model().columnCount())

    # insert header
    for col in new_df.columns:
        cursor.insertText(str(col))
        cursor.movePosition(QtGui.QTextCursor.NextCell)

    # insert data 
    for index, row in new_df.iterrows():
        for col in new_df.columns:            
            cursor.insertText(str(row[col]))
            cursor.movePosition(QtGui.QTextCursor.NextCell)
    document.print_(printer)

def arabify(arabic_text):
    #reshape the text       
    configuration = {
        'use_unshaped_instead_of_isolated': True,
    }
    reshaper = ArabicReshaper(configuration=configuration)
    rehaped_text = reshaper.reshape(arabic_text)
    bidi_text = get_display(rehaped_text)
    return bidi_text

# def select_row():
#     global table
#     indexes = table.selectionModel().selectedRows()
#     for index in sorted(indexes):
#         print('Row %d is selected' % index.row())

# if __name__ == "__main__":

app = QApplication(sys.argv)
w = QWidget()
w.resize(800,800)
w.setWindowTitle("Qasaym")

statusComboBox = QComboBox()
statusComboBox.addItems(template['الحاله'].to_dict().values())
recLael = QLabel("&رقم القسيمة")
recTextEdit = QLineEdit()
recTextEdit.resize(50,40)
recLael.setBuddy(recTextEdit)

# allow only integers
onlyInt = QIntValidator()
onlyInt.setRange(100000, 999999)
recTextEdit.setValidator(onlyInt)

nameLael = QLabel("&الإسم")
nameTextEdit = QLineEdit()
nameTextEdit.resize(50,40)
nameLael.setBuddy(nameTextEdit)

statusLabel = QLabel("&حالة")
statusLabel.setBuddy(statusComboBox)

addbtn = QPushButton()
addbtn.setText('&اضافة')
add_click_func = partial(add_click,addbtn)
addbtn.clicked.connect(add_click_func)

topLayout = QHBoxLayout()
topLayout.addWidget(recTextEdit)
topLayout.addWidget(recLael)
topLayout.addWidget(nameTextEdit)
topLayout.addWidget(nameLael)
topLayout.addWidget(statusComboBox)
topLayout.addWidget(statusLabel)    
topLayout.addWidget(addbtn)    

removebtn = QPushButton()
removebtn.setText('&حذف')
removebtn.clicked.connect(remove_dialog)

exportbtn = QPushButton()
exportbtn.setText('&حفظ إلي إكسيل')
exportbtn.clicked.connect(partial(saveFileDialog,exportbtn,'excel'))

pdfbtn = QPushButton()
pdfbtn.setText('&حفظ التقرير للطباعة')
pdfbtn.clicked.connect(partial(saveFileDialog,pdfbtn,'pdf'))

printbtn = QPushButton()
printbtn.setText('&طباعة')


buttonsLayout = QHBoxLayout()
buttonsLayout.addWidget(removebtn)
buttonsLayout.addStretch(1)
buttonsLayout.addWidget(exportbtn)
buttonsLayout.addWidget(pdfbtn)
# Uncomment to show print preview button
# buttonsLayout.addWidget(printbtn)

model = DataFrameModel(new_df)
table = QTableView()
table.setModel(model)
table.doubleClicked.connect(clicked_table)
printbtn.clicked.connect(partial(print_widget,table))

grid = QGridLayout()
grid.addLayout(topLayout,0,0,1,2)
grid.addLayout(buttonsLayout,1,0,1,2)
grid.addWidget(table, 2, 0,3,2)
w.setLayout(grid)    
w.show()
sys.exit(app.exec_())