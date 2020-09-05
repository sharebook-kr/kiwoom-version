import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QAxContainer import *
from util import *
from PyQt5 import QtTest

f = open("user.txt")
lines = f.readlines()
PATH = lines[0].strip()
USER_ID = lines[1].strip()
USER_PW = lines[2].strip()
USER_CR = lines[3].strip()
USER_PW2 = lines[4].strip()
f.close()


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        # 키움 객체 생성
        self.ocx = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self.ocx.OnEventConnect.connect(self._handler_login)
        # 5초 후에 로그인  
        QTimer.singleShot(5 * 1000, self.CommConnect)
        
        # 2초 마다 버전처리 창이 있는지 확인
        self.timer = QTimer(self)
        self.timer.start(2000)
        self.timer.timeout.connect(self.check_version_window)

    def CommConnect(self):
        self.ocx.dynamicCall("CommConnect()")
        QTimer.singleShot(5 * 1000, self.input_keys)

    def _handler_login(self):
        QTimer.singleShot(4 * 1000, self.auto_off)
        self.ocx.dynamicCall("KOA_Functions(QString, QString)", "ShowAccountWindow", "")
        # 자동로그인 off 되면 프로그램 자동종료 
        QApplication.instance().quit()

    def input_keys(self):
        try:
            hwnd = find_window("Open API Login")
            edit_id = win32gui.GetDlgItem(hwnd, 0x3E8)
            edit_pass = win32gui.GetDlgItem(hwnd, 0x3E9)
            edit_cert = win32gui.GetDlgItem(hwnd, 0x3EA)
            button = win32gui.GetDlgItem(hwnd, 0x1)

            enter_keys(edit_id, USER_ID)
            enter_keys(edit_pass, USER_PW)
            enter_keys(edit_cert, USER_CR)
            click_button(button)
        except:
            pass
            
    def auto_off(self):
        try:
            hwnd = find_window("계좌비밀번호")
            checkbox = win32gui.GetDlgItem(hwnd, 0xCD)
            checked = win32gui.SendMessage(checkbox, win32con.BM_GETCHECK)
            if checked:
                click_button(checkbox)
                #win32gui.SendMessage(checkbox, win32con.BM_SETCHECK, 0)

            QtTest.QTest.qWait(1000)
            button= win32gui.GetDlgItem(hwnd, 0x01)
            click_button(button)
        except:
            print("auto 해제 실패")

    def check_version_window(self):
        try:
            hwnd = find_window("opstarter")
            if hwnd != 0:
                static_hwnd = win32gui.GetDlgItem(hwnd, 0xFFFF)
                text = win32gui.GetWindowText(static_hwnd)
                if '버전처리를' in text:
                    QApplication.instance().quit()
        except:
            pass


if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = MyWindow()
    app.exec_()