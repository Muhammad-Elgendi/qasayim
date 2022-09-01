import sys
from PyQt5.QtWidgets import *

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

class JsonEditor(QMainWindow):
    def __init__(self, dictionary):
        super().__init__()
        self.dictionary = dictionary
        scroll = QScrollArea()

        self.setWindowTitle("Main Window")
        self.setGeometry(200, 200, 800, 100)

        self.main_widget = QWidget(self)
        self.setCentralWidget(self.main_widget)
        self.main_layout = QVBoxLayout(self.main_widget)
        scroll.setWidget(self)
        scroll.setWidgetResizable(True)
        scroll.setFixedHeight(400)
        self.main_layout.addWidget(scroll)

        self.createDynamicForm()
        self.createUpdateButton()

    def createDynamicForm(self):
        self.dynamiclayout = QFormLayout()
        self.dynamic_dictionary = {}
        for key, value in self.dictionary.items():
            print (key, value)
            self.dynamic_dictionary[key] = QLineEdit(str(value))
            self.dynamiclayout.addRow(key, self.dynamic_dictionary[key])
        self.main_layout.insertLayout(0, self.dynamiclayout)

    def createUpdateButton(self):
        self.update_button = QPushButton('update')
        self.main_layout.addWidget(self.update_button)
        self.update_button.clicked.connect(self.updateDictionary)

    def updateDictionary(self):
        dictionary2 = {}
        dictionary2['foo2'] = 'foo_string2'
        dictionary2['bar2'] = 'bar_string2'
        dictionary2['foo_bar'] = 'foo_bar_string2'
        self.dictionary = dictionary2

        deleteItemsOfLayout(self.dynamiclayout)
        self.createDynamicForm()

dictionary1 = {}
dictionary1['foo'] = 'foo_string'
dictionary1['bar'] = 'bar_string'


if __name__ == "__main__":
    test_app = QApplication(sys.argv)
    MainWindow = JsonEditor(dictionary1)
    MainWindow.show()
    sys.exit(test_app.exec_())