import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QAxContainer import *


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Python 로그인")
        self.ocx = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self.ocx.OnEventConnect.connect(self._handler_login)

        # login
        # 10초후에 로그인 실행
        QTimer.singleShot(10, self.CommConnect)


    def CommConnect(self):
        self.ocx.dynamicCall("CommConnect()")


    def _handler_login(self):
        self.ocx.dynamicCall("KOA_Functions(QString, QString)", "ShowAccountWindow", "")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = MyWindow()
    win.show()
    app.exec_()

