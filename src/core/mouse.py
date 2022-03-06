import uiautomation as auto
import win32api
import win32con


def move_wheel(distance: int):
    win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, 0, 0, distance)


def scroll_inside(target: auto.Control, bound: auto.Control):
    while True:
        if target.BoundingRectangle.top < bound.BoundingRectangle.top:
            bound.MoveCursorToMyCenter(simulateMove=False)
            move_wheel(bound.BoundingRectangle.top - target.BoundingRectangle.top)
        elif target.BoundingRectangle.bottom > bound.BoundingRectangle.bottom:
            bound.MoveCursorToMyCenter(simulateMove=False)
            move_wheel(bound.BoundingRectangle.bottom - target.BoundingRectangle.bottom)
        else:
            break
