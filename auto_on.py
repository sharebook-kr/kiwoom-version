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
        self.ocx = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self.ocx.OnEventConnect.connect(self._handler_login)
        QTimer.singleShot(5 * 1000, self.CommConnect)

    def CommConnect(self):
        self.ocx.dynamicCall("CommConnect()")
        QTimer.singleShot(5 * 1000, self.input_keys)

    def _handler_login(self):
        QTimer.singleShot(5 * 1000, self.auto_on)
        self.ocx.dynamicCall("KOA_Functions(QString, QString)", "ShowAccountWindow", "")
        # 자동로그인 off 되면 프로그램 자동종료 
        QApplication.instance().quit()

    def input_keys(self):
        print("input keys")
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
            
    def auto_on(self):
        try:
            hwnd = find_window("계좌비밀번호")
            if hwnd != 0:
                # 비밀번호등록
                edit = win32gui.GetDlgItem(hwnd, 0xCA)

                win32gui.SendMessage(edit, win32con.WM_SETTEXT, 0, USER_PW2)
                #enter_keys(edit, USER_PW2)
                print(edit, "enter password")

                win32api.Sleep(100)
                button_register_all = win32gui.GetDlgItem(hwnd, 0xCE)
                click_button(button_register_all)

                # 체크박스 체크 
                checkbox = win32gui.GetDlgItem(hwnd, 0xCD)
                checked = win32gui.SendMessage(checkbox, win32con.BM_GETCHECK)
                if not checked:
                    click_button(checkbox)

                QtTest.QTest.qWait(1000)
                button= win32gui.GetDlgItem(hwnd, 0x01)
                click_button(button)
        except:
            print("auto on 실패")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = MyWindow()
    app.exec_()