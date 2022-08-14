import sys
from PyQt5.QtWidgets import *
from PyQt5.QtGui import QIcon, QIntValidator
from PyQt5.QtCore import pyqtSlot
from PyQt5 import QtCore, QtGui
import pandas as pd
import json
from PandasView import DataFrameModel
from functools import partial
from datetime import datetime




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
    # mbox = QMessageBox()
    # mbox.setText("Your allegiance has been noted")
    # mbox.setDetailedText("You are now a disciple and subject of the all-knowing Guru")
    # mbox.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
    # mbox.exec_()

def add_new_qasema(number,name,status,taqseem):
    record = template[ template['الحاله'] == status ].copy()
    record['الاسـم'] = name
    record['رقم القسيمة'] = number
    for key,val in taqseem.items():
        record[key] = val
    
    record['التاريخ'] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    global new_df
    new_df = pd.concat([new_df, record], ignore_index = True)
    new_df = new_df.reset_index(drop=True)
    return new_df


def add_click():
    name = nameTextEdit.text().strip()
    rec = recTextEdit.text().strip()
    status = statusComboBox.currentText()
    global new_df
    global table
    if name and rec:
        new_df = add_new_qasema(rec,name,status,{})
        model = DataFrameModel(new_df)
        table.setModel(model)

def clicked_table():
    global table
    index = table.selectedIndexes()[0]
    data = table.model().data(index)
    print ("data : " + str(data))
    print ("row : " , index.row())
    print ("col : " , index.column())
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

def save_excel():
    file_name = datetime.now().strftime("%d-%m-%Y__%H-%M-%S")+'.xlsx'
    new_df.drop(['التاريخ'],axis = 1).to_excel(file_name, sheet_name='Qasaym 1', index=False, engine='openpyxl')    
    mbox = QMessageBox()
    mbox.setText("تم حفظ القسائم في ملف "+file_name)
    mbox.setStandardButtons(QMessageBox.Ok)
    mbox.setDefaultButton(QMessageBox.Ok)
    mbox.exec_()

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
addbtn.clicked.connect(add_click)

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
exportbtn.clicked.connect(save_excel)

buttonsLayout = QHBoxLayout()
buttonsLayout.addWidget(removebtn)
buttonsLayout.addStretch(1)
buttonsLayout.addWidget(exportbtn)

model = DataFrameModel(new_df)
table = QTableView()
table.setModel(model)
table.doubleClicked.connect(clicked_table)

grid = QGridLayout()
grid.addLayout(topLayout,0,0,1,2)
grid.addLayout(buttonsLayout,1,0,1,2)
grid.addWidget(table, 2, 0,3,2)
w.setLayout(grid)    
w.show()
sys.exit(app.exec_())
