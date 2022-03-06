import time

import win32api
import win32con


def key_down(code: int):
    win32api.keybd_event(code, win32api.MapVirtualKey(code, 0), 0, 0)


def key_up(code: int):
    win32api.keybd_event(code, win32api.MapVirtualKey(code, 0), win32con.KEYEVENTF_KEYUP, 0)


def key_press(code: int):
    key_down(code)
    time.sleep(0.02)
    key_up(code)
