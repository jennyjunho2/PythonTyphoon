import sys
from PyQt5.QtWidgets import *
from PyQt5 import uic

form_class = uic.loadUiType("test.ui")[0]

class WindowClass(QDialog, form_class) :
    def __init__(self) :
        super().__init__()
        self.setupUi(self)
        self.setFixedSize(1151, 481)

class first_tab(QWidget):
    def __init__(self):
        super().__init__()


    def getText(self):
        excel_directory = self.QTextEdit_excel_directory.toPlainText()
        print(excel_directory)

if __name__ == "__main__" :
    app = QApplication(sys.argv)
    myWindow = WindowClass()
    myWindow.show()
    app.exec_()