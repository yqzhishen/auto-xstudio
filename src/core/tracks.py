import uiautomation as auto

import log
import mouse
import singers

logger = log.logger


class Track:
    def __init__(self, index: int):
        if index < 1:
            logger.error('轨道编号最小为 1。')
            exit(1)
        self.track_window = auto.WindowControl(searchDepth=1, RegexName='X Studio .*').CustomControl(searchDepth=1, ClassName='TrackWin')
        self.index = index
        self.pane = None

    def exists(self) -> bool:
        self.pane = self.track_window.CustomControl(searchDepth=2, foundIndex=self.index, ClassName='TrackChannelControlPanel')
        return self.pane.Exists(maxSearchSeconds=0.5)

    def is_instrumental(self) -> bool:
        return self.pane.ComboBoxControl(searchDepth=1, ClassName='ComboBox').IsOffscreen

    def switch_singer(self, singer: str):
        """
        切换歌手。
        :param singer: 歌手名字
        """
        if not self.exists():
            logger.error('未找到对应序号的轨道。')
            exit(1)
        if self.is_instrumental():
            logger.error('指定的轨道不是演唱轨。')
            exit(1)
        bound = self.track_window.PaneControl(searchDepth=1, ClassName='ScrollViewer').BoundingRectangle
        top, bottom = bound.top, bound.bottom
        switch_button = self.pane.ButtonControl(searchDepth=2, AutomationId='switchSingerButton')
        while True:
            if switch_button.BoundingRectangle.top < top:
                self.track_window.MoveCursorToMyCenter(simulateMove=False)
                mouse.move_wheel(top - switch_button.BoundingRectangle.top)
            elif switch_button.BoundingRectangle.bottom > bottom:
                self.track_window.MoveCursorToMyCenter(simulateMove=False)
                mouse.move_wheel(bottom - switch_button.BoundingRectangle.bottom)
            else:
                switch_button.DoubleClick(simulateMove=False)
                break
        singers.choose_singer(name=singer)
        logger.info('为轨道 %d 切换歌手：%s。' % (self.index, singer))
