import uiautomation as auto

import log
import mouse
import singers

logger = log.logger


all_tracks = []


def enum_tracks() -> list:
    all_tracks.clear()
    index = 1
    track = Track(index)
    while track.exists():
        all_tracks.append(track)
        index += 1
        track = Track(index)
    return all_tracks


def count_tracks() -> int:
    return len(enum_tracks())


class Track:
    def __init__(self, index: int):
        if index < 1:
            logger.error('轨道编号最小为 1。')
            exit(1)
        self.index = index
        self.track_window = auto.WindowControl(searchDepth=1, RegexName='X Studio .*').CustomControl(searchDepth=1, ClassName='TrackWin')
        self.scroll_bound = self.track_window.PaneControl(searchDepth=1, ClassName='ScrollViewer')
        self.pane = self.track_window.CustomControl(searchDepth=2, foundIndex=self.index, ClassName='TrackChannelControlPanel')
        self.mute_button = self.pane.ButtonControl(searchDepth=1, Name='UnMute')
        self.unmute_button = self.pane.ButtonControl(searchDepth=1, Name='Mute')
        self.solo_button = self.pane.ButtonControl(searchDepth=1, Name='notSolo')
        self.notsolo_button = self.pane.ButtonControl(searchDepth=1, Name='Solo')
        self.switch_button = self.pane.ButtonControl(searchDepth=2, AutomationId='switchSingerButton')

    def exists(self) -> bool:
        return self.pane.Exists(maxSearchSeconds=0.5)

    def is_instrumental(self) -> bool:
        return self.pane.ComboBoxControl(searchDepth=1, ClassName='ComboBox').IsOffscreen

    def is_muted(self) -> bool:
        return self.unmute_button.Exists(maxSearchSeconds=0.5)

    def is_solo(self) -> bool:
        return self.notsolo_button.Exists(maxSearchSeconds=0.5)

    def set_muted(self, muted: bool):
        if muted and not self.is_muted():
            mouse.scroll_inside(target=self.mute_button, bound=self.scroll_bound)
            self.mute_button.Click(simulateMove=False)
        elif not muted and self.is_muted():
            mouse.scroll_inside(target=self.unmute_button, bound=self.scroll_bound)
            self.unmute_button.Click(simulateMove=False)
        if muted:
            logger.info('静音轨道 %d。' % self.index)
        else:
            logger.info('取消静音轨道 %d。' % self.index)

    def set_solo(self, solo: bool):
        if solo and not self.is_solo():
            mouse.scroll_inside(target=self.solo_button, bound=self.scroll_bound)
            self.solo_button.Click(simulateMove=False)
        elif not solo and self.is_solo():
            mouse.scroll_inside(target=self.notsolo_button, bound=self.scroll_bound)
            self.notsolo_button.Click(simulateMove=False)
        if solo:
            logger.info('独奏轨道 %d。' % self.index)
        else:
            logger.info('取消独奏轨道 %d。' % self.index)

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
        mouse.scroll_inside(target=self.switch_button, bound=self.scroll_bound)
        self.switch_button.DoubleClick(simulateMove=False)
        singers.choose_singer(name=singer)
        logger.info('为轨道 %d 切换歌手：%s。' % (self.index, singer))
