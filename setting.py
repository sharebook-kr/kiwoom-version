import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QAxContainer import *


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setGeometry(300, 300, 200, 200)

        self.ocx = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self.ocx.OnEventConnect.connect(self._handler_login)

        btn_login = QPushButton("로그인", self)
        btn_login.resize(160, 40)
        btn_login.move(20, 20)
        btn_login.clicked.connect(self.CommConnect)

        btn_password = QPushButton("계좌비밀번호 입력", self)
        btn_password.resize(160, 40)
        btn_password.move(20, 70)
        btn_password.clicked.connect(self.set_password)

    def CommConnect(self):
        self.ocx.dynamicCall("CommConnect()")

    def set_password(self):
        self.ocx.dynamicCall("KOA_Functions(QString, QString)", "ShowAccountWindow", "")

    def _handler_login(self, err_code):
        if err_code == 0:
            msg = "로그인 완료"
        else:
            msg = "로그인 실패"
        self.statusBar().showMessage(msg)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = MyWindow()
    win.show()
    app.exec_()
