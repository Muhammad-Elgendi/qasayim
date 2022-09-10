import sys
from PyQt5.QtWidgets import *
from PyQt5 import QtCore
from PandasView import DataFrameModel

# copy from https://stackoverflow.com/a/45790404/9758790
def deleteItemsOfLayout(layout):
    if layout is not None:
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.setParent(None)
            else:
                deleteItemsOfLayout(item.layout())

class ManualCaseWindow(QWidget):
    def __init__(self,parent_widget, dictionary):
        super().__init__()

        # center window
        centerPoint = QDesktopWidget().availableGeometry().center()
        qtRectangle = self.frameGeometry()
        qtRectangle.moveCenter(centerPoint)
        self.move(qtRectangle.topLeft())

        self.parent_widget = parent_widget
        self.setWindowModality(QtCore.Qt.ApplicationModal)
        self.dictionary = dictionary
        self.main_layout = QVBoxLayout(self)
        self.scroll = QScrollArea(self)
        self.scroll.resize(800, 480)
        self.main_layout.addWidget(self.scroll)
        self.scroll.setWidgetResizable(True)
        self.scrollContent = QWidget(self.scroll)
        self.scrollLayout = QVBoxLayout(self.scrollContent)
        self.scrollContent.setLayout(self.scrollLayout)
        self.scroll.setWidget(self.scrollContent)
        self.setWindowTitle("إضافة حالة يدوية")
        # self.setGeometry(200, 200, 800, 100)
        self.resize(810, 490)


        self.createDynamicForm()
        self.createUpdateButton()

    def createDynamicForm(self):
        self.dynamiclayout = QFormLayout()
        self.dynamic_dictionary = {}
        for key, value in self.dictionary.items():
            # print (key, value)
            self.dynamic_dictionary[key] = QLineEdit(str(value))
            self.dynamiclayout.addRow(key, self.dynamic_dictionary[key])
        self.scrollLayout.insertLayout(0, self.dynamiclayout)

    def createUpdateButton(self):
        self.update_button = QPushButton('إضافة القسيمة')
        self.scrollLayout.addWidget(self.update_button)
        self.update_button.clicked.connect(self.updateDictionary)

    def closeEvent(self, e):
        e.ignore()
        self.hide()
        self.parent_widget.show()

    def updateDictionary(self):
        # dictionary2 = {}
        # dictionary2['foo2'] = 'foo_string2'
        # dictionary2['bar2'] = 'bar_string2'
        # dictionary2['foo_bar'] = 'foo_bar_string2'
        # self.dictionary = dictionary2
        # deleteItemsOfLayout(self.dynamiclayout)
        # self.createDynamicForm()
        new_dictionary = {}
        for row in range(len(self.dictionary)):
            key = self.dynamiclayout.itemAt(row,0).widget().text()
            value = self.dynamiclayout.itemAt(row,1).widget().text()
            new_dictionary[key] = value
        # print(new_dictionary)

        temp_df = self.parent_widget.add_new_qasema(
            self.parent_widget.recTextEdit.text().strip(),
            self.parent_widget.nameTextEdit.text().strip(),
            'حاله يدويه',
            new_dictionary)

        # if add_new_qasema() returns a vaild df 
        if temp_df is not None:
            self.parent_widget.new_df = temp_df
            model = DataFrameModel(self.parent_widget.new_df)
            self.parent_widget.table.setModel(model)
            self.parent_widget.nameTextEdit.setText('')
            self.parent_widget.recTextEdit.setText(str(int(self.parent_widget.recTextEdit.text().strip())+1))
            self.parent_widget.saved = False
            
        self.close()






# if __name__ == "__main__":
#     dictionary1 = {}
#     dictionary1['foo'] = 'foo_string'
#     dictionary1['bar'] = 'bar_string'
#     test_app = QApplication(sys.argv)
#     MainWindow = JsonEditor(dictionary1)
#     MainWindow.show()
#     sys.exit(test_app.exec_())