from fbs_runtime.application_context.PyQt5 import ApplicationContext
import sys
from PyQt5.QtWidgets import *
from PyQt5.QtGui import QIcon, QIntValidator
from PyQt5.QtCore import pyqtSlot
from PyQt5 import QtCore, QtGui, QtPrintSupport
import pandas as pd
import json
from PandasView import DataFrameModel
from ManualCaseWindow import ManualCaseWindow
from functools import partial
from datetime import datetime
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from arabic_reshaper import ArabicReshaper
from bidi.algorithm import get_display
import base64
import sqlite3
from sqlite3 import Error

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        # center window
        centerPoint = QDesktopWidget().availableGeometry().center()
        qtRectangle = self.frameGeometry()
        qtRectangle.moveCenter(centerPoint)
        self.move(qtRectangle.topLeft())

        self.setWindowModality(QtCore.Qt.ApplicationModal)

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

        dbbtn = QPushButton()
        dbbtn.setText('&إستخراج التقارير')
        dbbtn.setToolTip('تأكد من حفظ البيانات أولاً')
        dbbtn.resize(dbbtn.sizeHint())
        dbbtn.clicked.connect(self.exportdb)

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
        buttonsLayout.addWidget(dbbtn)
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

        self.manualWindow = QCheckBox("&تمكين إدخال الحالة اليدوية في نافذة منفصلة")
        self.manualWindow.setChecked(False)

        grid = QGridLayout()
        grid.addLayout(topLayout,0,0,1,2)
        grid.addLayout(buttonsLayout,1,0,1,2)
        grid.addWidget(self.manualWindow, 2, 0,1,2)
        grid.addWidget(self.table, 3, 0,3,2)
        self.setLayout(grid)

    def closeEvent(self, event):
        reply = QMessageBox.question(self, 'Qasayim Close', 'هل تريد حفظ التغيرات و الخروج؟',
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            if len(self.new_df) > 0:
                # Storing in Database
                try:
                    conn = create_connection(appctxt.get_resource("qasayim.db"))
                    if conn is not None:
                        ##push the dataframe to sql 
                        self.new_df.to_sql("qasayim", conn, if_exists="append", index = False)
                except Error as e:
                    print(e)

                # save to excel            
                self.saveFileDialog("excel")

            # close Qasayim
            event.accept()
            print('Qasayim closed')
        else:
            event.ignore()
        
    def add_new_qasema(self,number,name,status,record_dict = None):
        if record_dict == None:
            record = self.template[ self.template['الحاله'] == status ].copy().reset_index(drop=True)
        else:
            record = self.template[ self.template['الحاله'] == 'شهاده أصليه ماجستير' ].copy().reset_index(drop=True)
            for col in record_dict.keys():
                record[col] = record_dict[col]
            record['الحاله'] = status

        auto_sum = True
        if status == 'حاله يدويه' and self.manualWindow.isChecked() and record_dict == None:
            dictionary = record.drop(['الحاله','الجمله','الاسـم','رقم القسيمة'],axis=1)
            dictionary.loc[0,:] = 0.0
            dictionary = dictionary.iloc[0].to_dict()
            self.w = ManualCaseWindow(self,dictionary)
            self.w.show()
            self.hide()
            return None
        else:
            for key in record.iloc[0][record.iloc[0].str.contains('دخل', na=False)].to_dict().keys():
                if key == 'الجمله':
                    auto_sum = False
                value, done = QInputDialog.getDouble(
                    self, 'إدخال قيمة عشرية', str(record[key].iloc[0])+' في '+key+' :')
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

    def exportdb(self):
        self.dbwindow = MyDBWindow(self)
        self.dbwindow.show()
        self.hide()

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
        if (len(self.new_df) > 0 and not self.saved) or len(self.new_df) == 0:
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

def decode_pass(encoded):
    """
    Decoding a Base64 string is essentially a reverse of the encoding process.
    We decode the Base64 string into bytes of unencoded data.
    We then convert the bytes-like object into a string.
    """
    base64_bytes = str(encoded).strip().encode('utf-8')
    message_bytes = base64.b64decode(base64_bytes)
    return str(message_bytes.decode('utf-8')).strip()

def create_connection(db_file):
    """ create a database connection to a SQLite database """
    conn = None
    try:
        conn = sqlite3.connect(db_file)
    except Error as e:
        print(e)
    finally:
        if conn:
            return conn
            
class LoginForm(QWidget):
    def __init__(self):
        super().__init__()

        # center window
        centerPoint = QDesktopWidget().availableGeometry().center()
        qtRectangle = self.frameGeometry()
        qtRectangle.moveCenter(centerPoint)
        self.move(qtRectangle.topLeft())

        self.setWindowTitle('تسجيل الدخول لقسائم')
        self.resize(500, 120)

        # Read template file that has logins
        self.logins = pd.read_excel(
                        appctxt.get_resource("qasayim_template.xlsx"),
                        engine='openpyxl',
                        sheet_name='logins')

        layout = QGridLayout()

        label_name = QLabel('<font size="4"> إسم المستخدم </font>')
        self.lineEdit_username = QLineEdit()
        self.lineEdit_username.setPlaceholderText('أدخل إسم المستخدم')
        layout.addWidget(label_name, 0, 0)
        layout.addWidget(self.lineEdit_username, 0, 1)

        label_password = QLabel('<font size="4"> كلمة المرور </font>')
        self.lineEdit_password = QLineEdit()
        self.lineEdit_password.setPlaceholderText('أدخل كلمة المرور')
        layout.addWidget(label_password, 1, 0)
        layout.addWidget(self.lineEdit_password, 1, 1)

        button_login = QPushButton('دخول')
        button_login.clicked.connect(self.check_password)
        layout.addWidget(button_login, 2, 0, 1, 2)
        layout.setRowMinimumHeight(2, 75)

        self.setLayout(layout)

    def check_password(self):
        msg = QMessageBox()
        access = False
        for index, row in self.logins.iterrows():
            if self.lineEdit_username.text() == decode_pass(row['usr']) and self.lineEdit_password.text() == decode_pass(row['pass']):
                access = True
                break

        if access:
            msg.setText('تم تسجيل الدخول بنجاح')
            # hide message
            # msg.exec_()
            self.hide() 
            qasayim = MainWindow()
            qasayim.show()
        else:
            msg.setText('كلمة المرور خاطئة ، سيتم إغلاق البرنامج')
            msg.exec_()
            appctxt.app.quit()


class MyDBWindow(QWidget):
    def __init__(self,parent_widget):
        super().__init__()

        self.parent_widget = parent_widget
        # center window
        centerPoint = QDesktopWidget().availableGeometry().center()
        qtRectangle = self.frameGeometry()
        qtRectangle.moveCenter(centerPoint)
        self.move(qtRectangle.topLeft())

        # Read template file that has rules
        self.template = pd.read_excel(
                        appctxt.get_resource("qasayim_template.xlsx"),
                        engine='openpyxl',
                        sheet_name='rules')

        self.setWindowTitle('إستخراج التقارير من قاعدة البيانات')
        self.resize(800, 800)   

        layout = QGridLayout()

        label_name = QLabel('<font size="4"> من تاريخ </font>')
        self.start_date_edit = QLineEdit()
        self.start_date_edit.setPlaceholderText('أدخل تاريخ البداية')
        layout.addWidget(label_name, 0, 0)
        layout.addWidget(self.start_date_edit, 0, 1)

        label_password = QLabel('<font size="4"> إلي تاريخ </font>')
        self.end_date_edit = QLineEdit()
        self.end_date_edit.setPlaceholderText('أدخل تاريج النهاية')
        layout.addWidget(label_password, 1, 0)
        layout.addWidget(self.end_date_edit, 1, 1)

        label_columns = QLabel('<font size="4"> إختيار التقرير </font>')
        self.colComboBox = QComboBox()
        options = ['الكل']
        options.extend([str(option) for option in self.template['الحاله'].dropna().to_list()])
        self.colComboBox.addItems(options)
        layout.addWidget(label_columns, 2, 0)
        layout.addWidget(self.colComboBox, 2, 1)

        btn_export = QPushButton('إستخراج التقرير إلي إكسيل')
        btn_export.clicked.connect(self.export_db)
        layout.addWidget(btn_export, 3, 0, 1, 2)
        layout.setRowMinimumHeight(3, 75)

        self.setLayout(layout)

    def export_db(self):
        try:
            conn = create_connection(appctxt.get_resource("qasayim.db"))
            if conn is not None:
                df = pd.read_sql('SELECT * FROM qasayim', conn, parse_dates={"التاريخ": {"format":'%d-%m-%Y %H:%M:%S'}}) 
                file_name = "export_qasayim_"+datetime.now().strftime("%d-%m-%Y__%H-%M-%S")+'.xlsx'
                df = df[(df['التاريخ'] >= self.start_date_edit.text() ) & (df['التاريخ'] <= self.end_date_edit.text() )]
                col = self.colComboBox.currentText()
                if col == 'الكل':
                    self.parent_widget.new_df = df
                else:
                    self.parent_widget.new_df = df[col]
                self.hide()                
                self.parent_widget.show()
                self.parent_widget.saveFileDialog('excel')
        except Error as e:
            print(e)

if __name__ == "__main__":
    appctxt = ApplicationContext()       # 1. Instantiate ApplicationContext
    # app = QApplication(sys.argv)
    # app.aboutToQuit.connect(cleanup)
    # sys.exit(app.exec_())
    form = LoginForm()
    form.show()   
    appctxt.app.aboutToQuit.connect(cleanup)
    exit_code = appctxt.app.exec()      # 2. Invoke appctxt.app.exec()
    sys.exit(exit_code)