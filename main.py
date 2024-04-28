""" from PyQt5.QtWidgets import QApplication,QMainWindow
from interface import Ui_MainWindow
from sum.suma import apply_sum

class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        self.ui.pushButton.clicked.connect(self.sumar)

    def sumar(self):
        a = self.ui.lineEdit.text()
        b = self.ui.lineEdit_2.text()
        result = apply_sum(a,b)
        self.ui.lineEdit_3.setText(str(result))

def main():
    app = QApplication([])
    window = MyWindow()
    window.show()
    app.exec_()

if __name__ == "__main__": 
    main() """