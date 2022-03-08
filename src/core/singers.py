import uiautomation as auto

from exception import AutoXStudioException
import log

logger = log.logger


def choose_singer(name: str):
    """
    选择一名歌手。歌手市场必须处于打开状态。
    :param name: 歌手名字
    """
    singer_market = auto.WindowControl(searchDepth=2, Name='歌手市场')
    singer_market.HyperlinkControl(searchDepth=9, Name='全部歌手').Click(simulateMove=False)
    browser_pane = singer_market.PaneControl(searchDepth=3, ClassName='CefBrowserWindow')
    bottom = browser_pane.BoundingRectangle.bottom
    while True:
        singer_text = browser_pane.TextControl(searchDepth=14, Name=name)
        bottom_text = browser_pane.TextControl(searchDepth=14, Name='已经到底了')
        if singer_text.Exists(maxSearchSeconds=0.1) and 0 < singer_text.BoundingRectangle.bottom < bottom:
            singer_text.Click(simulateMove=False)
            break
        elif bottom_text.Exists(maxSearchSeconds=0.1) and bottom_text.BoundingRectangle.bottom > 0:
            singer_market.ButtonControl(searchDepth=1, AutomationId='btnClose').Click(simulateMove=False)
            logger.error('指定的歌手“%s”不存在。' % name)
            raise AutoXStudioException()
        else:
            browser_pane.MoveCursorToMyCenter(simulateMove=False)
            auto.WheelDown(wheelTimes=2, waitTime=1)
    if singer_market.ButtonControl(searchDepth=17, Name='待解锁').Exists(maxSearchSeconds=0.5):
        singer_market.ImageControl(Depth=17).Click(simulateMove=False)
        logger.error('指定的歌手“%s”未解锁。' % name)
        raise AutoXStudioException()
    singer_market.ButtonControl(searchDepth=17, Name='选中').Click(simulateMove=False)
