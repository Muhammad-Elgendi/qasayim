from fbs_runtime.application_context.PyQt5 import ApplicationContext
import sys
from PyQt5.QtWidgets import *
from PyQt5.QtGui import QIcon, QIntValidator
from PyQt5.QtCore import pyqtSlot
from PyQt5 import QtCore, QtGui, QtPrintSupport
import pandas as pd
import json
from PandasView import DataFrameModel
# from AnotherWindow import JsonEditor
from functools import partial
from datetime import datetime
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from arabic_reshaper import ArabicReshaper
from bidi.algorithm import get_display

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        # self._createMenuBar()

        # Flag to indicate if data is saved
        self.saved = False

        # Read template file that has rules
        self.template = pd.read_excel(
                        appctxt.get_resource("qasayim_template.xlsx"),
                        engine='openpyxl',
                        sheet_name='rules')

        self.sum_rules = pd.read_excel(
                        appctxt.get_resource("qasayim_template.xlsx"),
                        engine='openpyxl',
                        sheet_name='sum_rules')        

        # new dataframe to save the data to
        self.new_df = pd.DataFrame(columns=self.template.columns)
  
        # keep up with the latest selected row
        self.last_selected_row = None

        self.resize(800,800)
        self.setWindowTitle("Qasayim")

        self.statusComboBox = QComboBox()
        self.statusComboBox.addItems([str(option) for option in self.template['الحاله'].dropna().to_list()])
        recLael = QLabel("&رقم القسيمة")
        self.recTextEdit = QLineEdit()
        self.recTextEdit.resize(50,40)
        recLael.setBuddy(self.recTextEdit)

        # allow only integers validator
        onlyInt = QIntValidator()
        onlyInt.setRange(100000, 999999)
        self.recTextEdit.setValidator(onlyInt)

        nameLael = QLabel("&الإسم")
        self.nameTextEdit = QLineEdit()
        self.nameTextEdit.resize(50,40)
        nameLael.setBuddy(self.nameTextEdit)

        statusLabel = QLabel("&حالة")
        statusLabel.setBuddy(self.statusComboBox)

        addbtn = QPushButton()
        addbtn.setText('&اضافة')
        addbtn.setToolTip('أدخل رقم القسيمة و الإسم و الحالة أولا للإضافة')
        addbtn.resize(addbtn.sizeHint())
        addbtn.clicked.connect(self.add_btn_click)

        topLayout = QHBoxLayout()
        topLayout.addWidget(self.recTextEdit)
        topLayout.addWidget(recLael)
        topLayout.addWidget(self.nameTextEdit)
        topLayout.addWidget(nameLael)
        topLayout.addWidget(self.statusComboBox)
        topLayout.addWidget(statusLabel)    
        topLayout.addWidget(addbtn)    

        removebtn = QPushButton()
        removebtn.setText('&حذف')
        removebtn.setToolTip('إضغط مرتين علي القسيمة المراد حذفها بالأسفل أولاً')
        removebtn.resize(removebtn.sizeHint())
        removebtn.clicked.connect(self.remove_qasema)

        exportbtn = QPushButton()
        exportbtn.setText('&حفظ إلي إكسيل')
        exportbtn.clicked.connect(partial(self.saveFileDialog,'excel'))

        pdfbtn = QPushButton()
        pdfbtn.setText('&حفظ التقرير للطباعة')
        pdfbtn.clicked.connect(partial(self.saveFileDialog,'pdf'))

        openbtn = QPushButton()
        openbtn.setText('&فتح ملف')
        openbtn.clicked.connect(self.openFileDialog)

        printbtn = QPushButton()
        printbtn.setText('&طباعة')

        buttonsLayout = QHBoxLayout()
        buttonsLayout.addWidget(removebtn)
        buttonsLayout.addStretch(1)
        buttonsLayout.addWidget(openbtn)
        buttonsLayout.addStretch(1)
        buttonsLayout.addWidget(exportbtn)
        buttonsLayout.addWidget(pdfbtn)
        # Uncomment to show print preview button
        # buttonsLayout.addWidget(printbtn)

        model = DataFrameModel(self.new_df)
        self.table = QTableView()
        self.table.setModel(model)
        self.table.doubleClicked.connect(self.clicked_table)
        printbtn.clicked.connect(self.print_widget)

        grid = QGridLayout()
        grid.addLayout(topLayout,0,0,1,2)
        grid.addLayout(buttonsLayout,1,0,1,2)
        grid.addWidget(self.table, 2, 0,3,2)
        self.setLayout(grid)

    def closeEvent(self, event):
        reply = QMessageBox.question(self, 'Qasayim Close', 'هل تريد حفظ التغيرات و الخروج؟',
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            # TODO Storing in Database

            # save to excel
            if len(self.new_df) > 0:
                self.saveFileDialog("excel")

            # close Qasayim
            event.accept()
            print('Qasayim closed')
        else:
            event.ignore()
        
    def add_new_qasema(self,number,name,status):
        record = self.template[ self.template['الحاله'] == status ].copy().reset_index(drop=True)
        # if status == 'حاله يدويه':
        #     self.w = JsonEditor(record.iloc[0].to_dict())
        #     self.w.show()
        #     self.hide()

        auto_sum = True
        for key in record.iloc[0][record.iloc[0].str.contains('دخل', na=False)].to_dict().keys():
            if key == 'الجمله':
                auto_sum = False
            value, done = QInputDialog.getDouble(
                self, 'Input Dialog', str(record[key].iloc[0])+' في '+key+' :')
            if done:
                record[key] = value
                items = [str(x) for x in self.sum_rules[key].dropna().to_list()]
                if len(items) == 1 and status != 'حاله يدويه':
                    record[key] = value * float(items[0])

                if len(items) > 1 and status != 'حاله يدويه':
                    item, ok = QInputDialog().getItem(self, "إختيار القيمة المطلوبة",
                                                        str(key)+" X :", items, 0, False)
                    if ok and item:
                        record[key] = value * float(item)
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

        self.new_df = pd.concat([self.new_df, record], ignore_index = True)
        self.new_df = self.new_df.reset_index(drop=True)
        return self.new_df

    def add_btn_click(self):
        name = self.nameTextEdit.text().strip()
        rec = self.recTextEdit.text().strip()
        status = self.statusComboBox.currentText()
        if name and rec:
            temp_df = self.add_new_qasema(rec,name,status)
            # if add_new_qasema() returns a vaild df 
            if temp_df is not None:
                self.new_df = temp_df
                model = DataFrameModel(self.new_df)
                self.table.setModel(model)
                self.nameTextEdit.setText('')
                self.recTextEdit.setText(str(int(rec)+1))
                self.saved = False

    def clicked_table(self):
        index = self.table.selectedIndexes()[0]
        data = self.table.model().data(index)
        # print ("data : " + str(data))
        # print ("row : " , index.row())
        # print ("col : " , index.column())
        self.last_selected_row = index.row()

    def remove_qasema(self):

        if self.last_selected_row == None:
            return None

        number = self.new_df.iloc[self.last_selected_row]['رقم القسيمة']
        name  = self.new_df.iloc[self.last_selected_row]['الاسـم']
        status = self.new_df.iloc[self.last_selected_row]['الحاله']
        msgBox = QMessageBox()
        msgBox.setWindowTitle("Qasayim")
        msgBox.setText("القسيمة رقم : "+str(number)+"\n"+"بإسم : "+str(name)+"\n"+ "نوع القسيمة : "+str(status))
        msgBox.setInformativeText("هل تريد حذف القسيمة؟")
        msgBox.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        msgBox.setDefaultButton(QMessageBox.Cancel)
        ret = msgBox.exec_()
    
        if ret == QMessageBox.Ok:
            self.new_df = self.new_df.drop([self.last_selected_row])
            self.new_df = self.new_df.reset_index(drop=True)
            model = DataFrameModel(self.new_df)
            self.table.setModel(model)
            self.last_selected_row = None
            self.saved = False


    def save_excel(self,file_name,default_name = True):
        if default_name:
            file_name = datetime.now().strftime("%d-%m-%Y__%H-%M-%S")+'.xlsx'
        
        # add horizental total            
        temp = self.new_df.copy().reset_index(drop=True)
        temp.loc['مجموع العمود']= temp.sum(numeric_only=True, axis=0)       
        temp.drop(['التاريخ'], axis = 1, errors='ignore').to_excel(file_name, sheet_name='Qasayim 1', index=True, engine='openpyxl')    
        show_saved_message(file_name)
        self.saved = True

    def save_pdf(self,file_name,default_name = True):
        if default_name:
            file_name = datetime.now().strftime("%d-%m-%Y__%H-%M-%S")+'.pdf'
        df = self.new_df.drop(['التاريخ'],axis = 1,errors='ignore')

        # show arabic chars correctly
        df['الاسـم'] = df['الاسـم'].apply(arabify) 
        df['الحاله'] = df['الحاله'].apply(arabify) 
        # show arabic header correctly,
        # cuz, matplotlib doesn't support it well
        new_arabic_cols = [arabify(col) for col in df.columns]
        # plot dataframe in matplotlib
        fig, ax =plt.subplots(figsize=(12,4))
        ax.axis('tight')
        ax.axis('off')
        the_table = ax.table(cellText=df.values,colLabels=new_arabic_cols,loc='center') 
        # the_table.auto_set_font_size(False)   
        # the_table.set_fontsize(10)
        the_table.scale(3, 3)  # may help

        # Export Matplotlib table to PDF
        pp = PdfPages(file_name)
        pp.savefig(fig, bbox_inches='tight')
        pp.close()

        show_saved_message(file_name)

    def openFileDialog(self):
        if len(self.new_df) > 0 and not self.saved:
            reply = QMessageBox.question(self, 'تأكد من حفظ الملف', 'هل تريد فتح ملف جديد دون حفظ التغيرات؟',
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        
            if reply == QMessageBox.Yes:
                options = QFileDialog.Options()
                options |= QFileDialog.DontUseNativeDialog

                fileName, _ = QFileDialog.getOpenFileName(self, "Open File",
                                            "/home",
                                            "Excel files (*.xls *.xlsx)")        
                if fileName:
                    df = pd.read_excel(
                            fileName,
                            engine='openpyxl',
                            sheet_name=0,
                            usecols= self.template.columns)

                    self.new_df = df.dropna()

                    model = DataFrameModel(self.new_df)
                    self.table.setModel(model)
            else:
                return None

       
    def saveFileDialog(self,file_type):
        if len(self.new_df) == 0:
            return None
        
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog

        fileName, _ = QFileDialog.getSaveFileName(self,"حفظ القسائم","","ALL FILES (*)", options=options)
        
        if fileName:
            # export file based on its type
            if file_type == 'excel':
                self.save_excel(fileName+".xlsx",False)
            elif file_type == 'pdf':
                self.save_pdf(fileName+".pdf",False)
    
    # print preview methods
    def print_widget(self):    
        previewDialog = QtPrintSupport.QPrintPreviewDialog()
        previewDialog.paintRequested.connect(self.handlePaintRequest)  
        previewDialog.exec_()

    def handlePaintRequest(self,printer):
        document = QtGui.QTextDocument()

        #############################################
        #### Working but bellow is better design ####
        # insert custom table
        cursor = QtGui.QTextCursor(document)

        table_format = QtGui.QTextTableFormat()
        table_format.setCellPadding(3)
        table_format.setHeaderRowCount(0)
        table_format.clearColumnWidthConstraints()

        df_to_print = self.new_df.drop(['التاريخ'],axis = 1, errors='ignore')

        print_table = cursor.insertTable(
            len(df_to_print),
            len(df_to_print.columns),
            table_format
            )        

        # insert header
        for col in df_to_print.columns:
            cursor.insertText(str(col))
            cursor.movePosition(QtGui.QTextCursor.NextCell)

        # insert data        
        for index, row in df_to_print.iterrows():
            for col in df_to_print.columns:            
                cursor.insertText(str(row[col]))
                cursor.movePosition(QtGui.QTextCursor.NextCell)
        document.print_(printer)

        ###################################################
        #### Create a html document from the dataframe ####
        # html = self.new_df.drop(['التاريخ'],axis = 1,errors='ignore').to_html(buf=None, index=False)
        # document = QtGui.QTextDocument()
        # document.setHtml(html)
        # document.print_(printer)

    def _createMenuBar(self):
        menuBar = QMenuBar(self)
        # Creating menus using a QMenu object
        fileMenu = QMenu("&File", self)
        menuBar.addMenu(fileMenu)
        # Creating menus using a title
        editMenu = menuBar.addMenu("&Edit")
        helpMenu = menuBar.addMenu("&Help")
############## UTILITIES #################

# Add dialog
def dialog():
    msgBox = QMessageBox()
    msgBox.setWindowTitle("Qasayim")
    msgBox.setText("تم إضافة قسيمة جديدة")
    msgBox.setInformativeText("هل تريد حفظ التعديلات؟")
    msgBox.setStandardButtons(QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel)
    msgBox.setDefaultButton(QMessageBox.Save)
    ret = msgBox.exec_()

# show saved file location
def show_saved_message(file_name):
    mbox = QMessageBox()
    mbox.setWindowTitle("Qasayim")

    mbox.setText("تم حفظ القسائم في ملف \n"+file_name)
    mbox.setStandardButtons(QMessageBox.Ok)
    mbox.setDefaultButton(QMessageBox.Ok)
    # mbox.setDetailedText("تم حفظ المعلومات الخاصة بالقسائم")
    # mbox.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
    mbox.exec_()

# reshape arabic words in matplotlib PDF
def arabify(arabic_text):
    #reshape the text       
    configuration = {
        'use_unshaped_instead_of_isolated': True,
    }
    reshaper = ArabicReshaper(configuration=configuration)
    rehaped_text = reshaper.reshape(arabic_text)
    bidi_text = get_display(rehaped_text)
    return bidi_text

# Doing some clean up right before closing application
def cleanup():
    print('CleanUp executed')

if __name__ == "__main__":
    appctxt = ApplicationContext()       # 1. Instantiate ApplicationContext
    # app = QApplication(sys.argv)
    qasayim = MainWindow()
    qasayim.show()
    # app.aboutToQuit.connect(cleanup)
    # sys.exit(app.exec_())
    exit_code = appctxt.app.exec()      # 2. Invoke appctxt.app.exec()
    sys.exit(exit_code)