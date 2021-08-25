import sys
from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5 import QtGui
from PyQt5.QtGui import QIcon


class WindowClass(QDialog) :
    def __init__(self) :
        super().__init__()
        self.ui = uic.loadUi("test.ui", self)
        self.setWindowTitle("검토용파일 제작 프로그램")
        self.setWindowIcon(QIcon("icon.png"))
        self.pushButton_execute.clicked.connect(self.execute_function)

    def getText_excel_directory(self):
        excel_directory = self.QTextEdit_excel_directory.toPlainText()

    def getText_test_name(self):
        test_name = self.QTextEdit_test_name.toPlainText()

    def execute_function(self):
        excel_directory = self.QTextEdit_excel_directory.toPlainText()
        test_name = self.QTextEdit_test_name.toPlainText()
        print(excel_directory)
        print(test_name)

if __name__ == "__main__" :
    app = QApplication(sys.argv)
    myWindow = WindowClass()
    myWindow.show()
    app.exec_()