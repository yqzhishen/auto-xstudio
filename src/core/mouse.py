import win32api
import win32con


def move_wheel(distance: int):
    win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, 0, 0, distance)
