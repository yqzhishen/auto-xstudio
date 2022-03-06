import uiautomation as auto


def press_key_combination(*keys: int):
    for k in keys:
        auto.PressKey(k, waitTime=0)
    for k in keys[::-1]:
        auto.ReleaseKey(k, waitTime=0)


def scroll_wheel_inside(target: auto.Control, bound: auto.Control):
    while True:
        if target.BoundingRectangle.top < bound.BoundingRectangle.top:
            bound.MoveCursorToMyCenter(simulateMove=False)
            auto.WheelUp(waitTime=0)
        elif target.BoundingRectangle.bottom > bound.BoundingRectangle.bottom:
            bound.MoveCursorToMyCenter(simulateMove=False)
            auto.WheelDown(waitTime=0)
        else:
            break
