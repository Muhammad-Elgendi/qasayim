from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QMainWindow, QLineEdit, QPushButton,
                             QTableWidget, QTableWidgetItem, QHeaderView)
import sys

class TableEditUI(QMainWindow):
    def __init__(self,df):
        super().__init__()

        self.setWindowTitle('إضافة حالة يدوية')
        self.setMinimumWidth(500)

        self.main_widget = QWidget(self)
        self.setCentralWidget(self.main_widget)

        self.create_widgets()
        self.create_layouts()

        self.toggle = True
        self.set_dictionary()
        self.refresh_table()

        self.create_connections()

    def create_widgets(self):
        self.table = QTableWidget()
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["Attribute", "Data"])
        header_view = self.table.horizontalHeader()
        header_view.setSectionResizeMode(1, QHeaderView.Stretch)
        self.toggle_dictionary_btn = QPushButton('toggle dictionary')

    def refresh_table(self):
        self.table.setRowCount(0)
        self.dynamic_dictionary = {}
        for count, (key, value) in enumerate(self.dictionary.items()):
            print(count, key, value)
            self.dynamic_dictionary[key] = QLineEdit(value)
            self.table.insertRow(count)

            self.insert_item(count, 0, str(key))
            self.table.setCellWidget(count, 1, self.dynamic_dictionary[key])

    def insert_item(self, row, column, text):
        item = QTableWidgetItem(text)
        self.table.setItem(row, column, item)

    def showEvent(self, e):
        super(TableEditUI, self).showEvent(e)
        self.refresh_table

    def create_layouts(self):
        self.main_layout = QVBoxLayout(self.main_widget)
        self.main_layout.addWidget(self.table)
        self.main_layout.addWidget(self.toggle_dictionary_btn)

    def create_connections(self):
        self.toggle_dictionary_btn.clicked.connect(self.set_dictionary)

    def set_dictionary(self):
        self.toggle = not self.toggle
        dictionary1 = {}
        dictionary1['foo'] = 'foo_string'
        dictionary1['bar'] = 'bar_string'

        dictionary2 = {}
        dictionary2['foo2'] = 'foo_string2'
        dictionary2['bar2'] = 'bar_string2'
        dictionary2['foo_bar'] = 'foo_bar_string2'

        if self.toggle:
            self.dictionary = dictionary1
        else:
            self.dictionary = dictionary2

        self.refresh_table()


if __name__ == "__main__":

    app = QApplication(sys.argv)
    TableEditUIWindow = TableEditUI()
    TableEditUIWindow.show()
    sys.exit(app.exec())