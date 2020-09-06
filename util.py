import win32gui
import win32con
import win32api
import time

def window_enumeration_handler(hwnd, top_windows):
    top_windows.append((hwnd, win32gui.GetWindowText(hwnd)))

def enum_windows():
    windows = []
    win32gui.EnumWindows(window_enumeration_handler, windows)
    return windows

def find_window(caption):
    hwnd = win32gui.FindWindow(None, caption)
    if hwnd == 0:
        windows = enum_windows()
        for handle, title in windows:
            if caption in title:
                hwnd = handle
                break
    return hwnd

def enter_keys(hwnd, password):
    win32gui.SetForegroundWindow(hwnd)
    win32api.Sleep(100)

    for c in password:
        win32api.SendMessage(hwnd, win32con.WM_CHAR, ord(c), 0)
        win32api.Sleep(100)

def click_button(btn_hwnd):
    win32api.PostMessage(btn_hwnd, win32con.WM_LBUTTONDOWN, 0, 0)
    win32api.Sleep(100)
    win32api.PostMessage(btn_hwnd, win32con.WM_LBUTTONUP, 0, 0)
    win32api.Sleep(100)

def close_window(title, secs=5):
    hwnd = find_window(title)
    if hwnd !=0:
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
        print("{} secs sleep ...".format(secs))
        time.sleep(secs)

def wait_secs(msg, secs=10):
    while secs > 0:
        time.sleep(1)
        print(f"{msg} waiting: {secs}")
        secs = secs - 1

