# UI AUTOMATION FOR X STUDIO
# AUTHOR: YQ之神
# VERSION: 1.1.0 (2022.3.5)

import logging
import os
import time
import winreg

import win32api
import win32con

import uiautomation as auto
import colorlog


logger = logging.getLogger()
log_colors_config = {
    'DEBUG': 'white',
    'INFO': 'green',
    'WARNING': 'yellow',
    'ERROR': 'red',
    'CRITICAL': 'bold_red',
}
console_handler = logging.StreamHandler()
logger.setLevel(logging.INFO)
console_handler.setLevel(logging.INFO)
console_formatter = colorlog.ColoredFormatter(
    fmt='%(log_color)s[%(asctime)s.%(msecs)03d] %(filename)s -> %(funcName)s line:%(lineno)d [%(levelname)s] : %(message)s',
    datefmt='%Y-%m-%d  %H:%M:%S',
    log_colors=log_colors_config
)
console_handler.setFormatter(console_formatter)
if not logger.handlers:
    logger.addHandler(console_handler)
console_handler.close()


def _verify_startup():
    """
    验证 X Studio 是否成功启动。
    """
    warning_window = auto.WindowControl(searchDepth=1, Name='提示')
    if warning_window.Exists(maxSearchSeconds=3):
        warning = warning_window.TextControl(searchDepth=1, AutomationId='Tbx').Name
        warning_window.ButtonControl(searchDepth=1, AutomationId='OkBtn').Click(simulateMove=False)
        logger.error(warning)
        exit(1)


def _verify_opening(base):
    """
    验证工程是否成功打开。
    """
    warning_window = base.WindowControl(searchDepth=1, ClassName='#32770')
    if warning_window.Exists(maxSearchSeconds=1):
        warning = warning_window.TextControl(searchDepth=2).Name.replace('\r\n', ' ').replace('。 ', '。')
        if warning.startswith('无法读取伴奏文件'):
            logger.warning('已自动忽略：无法读取伴奏文件。')
            while warning_window.Exists(maxSearchSeconds=1):
                warning_window.ButtonControl(searchDepth=1, Name='确定').Click(simulateMove=False)
                warning_window = auto.WindowControl(searchDepth=2, ClassName='#32770')
        elif '正在使用' in warning:
            warning_window.ButtonControl(searchDepth=2, Name='确定').Click(simulateMove=False)
            base.ButtonControl(searchDepth=1, Name='取消').Click(simulateMove=False)
            logger.error(warning)
            exit(1)
        else:
            warning_window.ButtonControl(searchDepth=1, Name='确定').Click(simulateMove=False)
            base.WindowControl(searchDepth=1, Name='X Studio').ButtonControl(searchDepth=1, AutomationId='OkBtn').Click(simulateMove=False)
            base.WindowControl(searchDepth=1, Name='X Studio').ButtonControl(searchDepth=1, AutomationId='btnClose').Click(simulateMove=False)
            logger.error(warning)
            exit(1)


def _verify_updates():
    """
    验证 X Studio 更新，并关闭更新提示窗。
    """
    update_window = auto.WindowControl(searchDepth=2, Name='检测到新版本')
    if update_window.Exists(maxSearchSeconds=3):
        update_window.ButtonControl(searchDepth=1, AutomationId='btnClose').Click(simulateMove=False)


def _key_down(code: int):
    win32api.keybd_event(code, win32api.MapVirtualKey(code, 0), 0, 0)


def _key_up(code: int):
    win32api.keybd_event(code, win32api.MapVirtualKey(code, 0), win32con.KEYEVENTF_KEYUP, 0)


def _key_press(code: int):
    _key_down(code)
    time.sleep(0.02)
    _key_up(code)


def find_xstudio() -> str:
    """
    根据注册表查找 X Studio 主程序路径。
    :return: XStudioSinger.exe 的路径
    """
    key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 'SOFTWARE\\Classes\\svipfile\\shell\\open\\command')
    value = winreg.QueryValueEx(key, '')
    return value[0].split('"')[1]


def start_xstudio(engine: str = None, project: str = None, singer: str = '陈水若'):
    """
    启动 X Studio。
    :param engine: 手动指定 X Studio 主程序路径
    :param project: 启动时需要打开的工程文件路径，默认打开空白工程
    :param singer: 若打开空白工程，可指定初始歌手名称
    """
    if engine:
        engine = os.path.abspath(engine)
        if not os.path.exists(engine):
            logger.error('指定的主程序路径不存在。')
            exit(1)
        if not os.path.isfile(engine) or not engine.endswith('.exe'):
            logger.error('指定的主程序不是合法的可执行 (.exe) 文件。')
            exit(1)
        logger.info('指定的主程序：%s。' % engine)
    if project:
        project = os.path.abspath(project)
        if not os.path.exists(project):
            logger.error('工程文件不存在。')
            exit(1)
        if not os.path.isfile(project) or not project.endswith('.svip'):
            logger.error('不是合法的 X Studio 工程 (.svip) 文件。')
            exit(1)
        if engine:
            os.popen(f'"{engine}" "{project}"')
        else:
            os.popen(f'"{project}"')
        _verify_startup()
        _verify_opening(auto)
        logger.info('启动 X Studio 并打开工程：%s。' % project)
        _verify_updates()
    else:
        if engine:
            os.popen(f'"{engine}"')
        else:
            os.popen(f'"{find_xstudio()}"')
        _verify_startup()
        auto.WindowControl(searchDepth=1, Name='X Studio').TextControl(searchDepth=2, Name='开始创作').Click(simulateMove=False)
        choose_singer(name=singer)
        logger.info('启动 X Studio 并创建空白工程，初始歌手：%s。' % singer)
        _verify_updates()


def quit_xstudio():
    """
    退出 X Studio。
    """
    auto.WindowControl(searchDepth=1, RegexName='X Studio .*').ButtonControl(searchDepth=1, AutomationId='btnClose').Click(simulateMove=False)
    confirm_window = auto.WindowControl(searchDepth=1, Name='X Studio')
    text = confirm_window.TextControl(searchDepth=1, AutomationId='Tbx').Name
    if text.startswith('确认'):
        confirm_window.ButtonControl(searchDepth=1, AutomationId='OkBtn').Click(simulateMove=False)
    else:
        confirm_window.ButtonControl(searchDepth=1, AutomationId='NoBtn').Click(simulateMove=False)
    logger.info('退出 X Studio。')


def new_project(singer: str = None):
    """
    新建工程。X Studio 必须已处于启动状态。
    :param singer: 可指定新工程的初始歌手
    """
    _key_down(17)
    _key_press(78)
    _key_up(17)
    confirm_window = auto.WindowControl(searchDepth=1, Name='X Studio')
    if confirm_window.Exists(maxSearchSeconds=1):
        confirm_window.ButtonControl(searchDepth=1, AutomationId='NoBtn').Click(simulateMove=False)
    if singer:
        track_window = auto.WindowControl(searchDepth=1, RegexName='X Studio .*').CustomControl(searchDepth=1, ClassName='TrackWin')
        track_window.CustomControl(searchDepth=2, ClassName='TrackChannelControlPanel').ButtonControl(searchDepth=2, AutomationId='switchSingerButton').DoubleClick(simulateMove=False)
        choose_singer(singer)
        logger.info('创建新工程，初始歌手：%s。' % singer)
    else:
        logger.info('创建新工程。')


def open_project(filename: str, folder: str = None):
    """
    打开工程。X Studio 必须已处于启动状态。
    :param filename: 工程文件名
    :param folder: 工程所处文件夹路径，默认为 X Studio 上一次打开工程的路径
    """
    if folder:
        project = os.path.abspath(os.path.join(folder, filename))
        if not os.path.exists(project):
            logger.error('工程文件不存在。')
            exit(1)
    else:
        project = filename
    if not filename.endswith('.svip'):
        logger.error('不是一个可打开的 X Studio 工程 (.svip) 文件。')
        exit(1)
    _key_down(17)
    _key_press(79)
    _key_up(17)
    confirm_window = auto.WindowControl(searchDepth=1, Name='X Studio')
    if confirm_window.Exists(maxSearchSeconds=1):
        confirm_window.ButtonControl(searchDepth=1, AutomationId='NoBtn').Click(simulateMove=False)
    main_window = auto.WindowControl(searchDepth=1, RegexName='X Studio .*')
    open_window = main_window.WindowControl(searchDepth=1, Name='打开文件')
    open_window.EditControl(searchDepth=3, Name='文件名(N):').GetValuePattern().SetValue(project)
    open_window.ButtonControl(searchDepth=1, Name='打开(O)').Click(simulateMove=False)
    warning_window = open_window.WindowControl(searchDepth=1, ClassName='#32770')
    if warning_window.Exists(maxSearchSeconds=1):
        warning = warning_window.TextControl(searchDepth=2).Name
        warning_window.ButtonControl(searchDepth=2, Name='确定').Click(simulateMove=False)
        open_window.ButtonControl(searchDepth=1, Name='取消').Click(simulateMove=False)
        logger.error(warning.replace('\r\n', ' ').replace('。 ', '。'))
        exit(1)
    _verify_opening(main_window)
    logger.info('打开工程：%s。' % project)


def export_project(title: str = None, folder: str = None, format: str = 'mp3', samplerate: int = 48000):
    """
    导出当前打开的工程。
    :param title: 目标文件名，默认与工程同名
    :param folder: 目标文件夹路径，默认为工程所在文件夹
    :param format: 导出格式 (mp3/wav/midi)，默认为 mp3
    :param samplerate: 采样率 (48000/44100)，默认为 48000
    """
    if format not in ['mp3', 'wav', 'midi']:
        logger.error('只能保存为 mp3, wav 或 midi 格式。')
        exit(1)
    if format == 'midi':
        samplerate = None
    elif samplerate != 48000 and samplerate != 44100:
        logger.error('采样率只能为 48000 或 44100。')
        exit(1)
    if folder and not os.path.exists(folder):
        folder = folder.replace('/', '\\')
        os.makedirs(folder)
    auto.ButtonControl(searchDepth=2, Name='导出').Click(simulateMove=False)
    setting_window = auto.WindowControl(searchDepth=2, Name='导出设置')
    if title:
        setting_window.EditControl(searchDepth=1, AutomationId='FileNameTbx').GetValuePattern().SetValue(title)
    if folder:
        logger.warning('当前尚不支持指定导出文件夹路径。')
        setting_window.EditControl(searchDepth=1, AutomationId='DestTbx').SendKeys(folder, interval=0.05)
    if format != 'mp3':
        format_box = setting_window.ComboBoxControl(searchDepth=1, AutomationId='FormatComboBox')
        format_box.Click(simulateMove=False)
        if format == 'wav':
            format_box.ListItemControl(searchDepth=1, Name='WAVE文件').Click(simulateMove=False)
        else:
            format_box.ListItemControl(searchDepth=1, Name='Midi文件').Click(simulateMove=False)
    if samplerate == 44100:
        samplerate_box = setting_window.ComboBoxControl(searchDepth=1, AutomationId='SampleRateComboBox')
        samplerate_box.Click(simulateMove=False)
        samplerate_box.ListItemControl(searchDepth=1, Name='44100HZ').Click(simulateMove=False)
    setting_window.ButtonControl(searchDepth=1, Name='导出').Click(simulateMove=False)
    export_window = auto.WindowControl(searchDepth=2, RegexName='导出.*')
    label = export_window.TextControl(searchDepth=1, ClassName='TextBlock', AutomationId='label')
    while True:
        message_window = auto.WindowControl(searchDepth=2, ClassName='#32770')
        if message_window.Exists(maxSearchSeconds=1):
            message = message_window.TextControl(searchDepth=1, ClassName='Static').Name
            message_window.ButtonControl(searchDepth=1, Name='确定').Click(simulateMove=False)
            logger.error(message + '。')
            exit(1)
        if label.Name == '导出成功':
            break
        elif label.Name.startswith('导出失败'):
            logger.error('导出失败，请稍后再试。')
            exit(1)
    export_window.ButtonControl(searchDepth=1, AutomationId='okBtn').Click(simulateMove=False)
    logger.info('导出工程：%s, 格式 %s, 采样率 %d Hz。' % (title, format, samplerate))


def save_project(filename: str = None, folder: str = None):
    """
    保存或另存为当前打开的工程。
    :param filename: 另存为的工程文件名
    :param folder: 另存为的文件夹路径，默认为工程所在文件夹
    """
    if folder:
        if not filename:
            logger.error('另存为工程时必须指定文件名。')
            exit(1)
        if not os.path.exists(folder):
            os.makedirs(folder)
        folder = os.path.abspath(folder.replace('/', '\\'))
    if not filename:
        _key_down(17)
        _key_press(83)
        _key_up(17)
        logger.info('保存工程。')
    else:
        if folder:
            project = os.path.join(folder, filename)
        else:
            project = filename
        _key_down(17)
        _key_down(16)
        _key_press(83)
        _key_up(16)
        _key_up(17)
        save_window = auto.WindowControl(searchDepth=1, RegexName='X Studio .*').WindowControl(searchDepth=1, Name='另存为')
        save_window.EditControl(searchDepth=6, Name='文件名:').GetValuePattern().SetValue(project)
        save_window.ButtonControl(searchDepth=1, Name='保存(S)').Click(simulateMove=False)
        confirm_window = save_window.WindowControl(searchDepth=1, ClassName='#32770')
        if confirm_window.Exists(maxSearchSeconds=1):
            warning = confirm_window.TextControl(searchDepth=2).Name
            if warning.endswith('是否替换它?'):
                confirm_window.ButtonControl(searchDepth=1, Name='是(Y)').Click(simulateMove=False)
            else:
                confirm_window.ButtonControl(searchDepth=2, Name='确定').Click(simulateMove=False)
                save_window.ButtonControl(searchDepth=1, Name='取消').Click(simulateMove=False)
                logger.error(warning.replace('\r\n', ' ').replace('。 ', '。'))
                exit(1)
        logger.info('另存为工程：%s。' % project)


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
        if singer_text.Exists(maxSearchSeconds=0.5) and 0 < singer_text.BoundingRectangle.bottom < bottom:
            singer_text.Click(simulateMove=False)
            break
        elif bottom_text.Exists(maxSearchSeconds=0.5) and bottom_text.BoundingRectangle.bottom > 0:
            singer_market.ButtonControl(searchDepth=1, AutomationId='btnClose').Click(simulateMove=False)
            logger.error('指定的歌手“%s”不存在。' % name)
            exit(1)
        else:
            browser_pane.MoveCursorToMyCenter(simulateMove=False)
            win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, 0, 0, -1500)
            time.sleep(1)
    if singer_market.ButtonControl(searchDepth=17, Name='待解锁').Exists(maxSearchSeconds=0.5):
        singer_market.ImageControl(Depth=17).Click(simulateMove=False)
        logger.error('指定的歌手“%s”未解锁。' % name)
        exit(1)
    singer_market.ButtonControl(searchDepth=17, Name='选中').Click(simulateMove=False)


def switch_singer(name: str, track: int = 1):
    """
    为指定的轨道切换歌手。
    :param track: 目标演唱轨序号（即工程中的第几条轨道），从 1 开始，默认为 1
    :param name: 歌手名字
    """
    if track < 1:
        logger.error('轨道编号最小为 1。')
        exit(1)
    track_window = auto.WindowControl(searchDepth=1, RegexName='X Studio .*').CustomControl(searchDepth=1, ClassName='TrackWin')
    channel_pane = track_window.CustomControl(searchDepth=2, foundIndex=track, ClassName='TrackChannelControlPanel')
    if not channel_pane.Exists(maxSearchSeconds=2):
        logger.error('未找到对应序号的轨道。')
        exit(1)
    if channel_pane.ComboBoxControl(searchDepth=1, ClassName='ComboBox').IsOffscreen:
        logger.error('指定的轨道不是演唱轨。')
        exit(1)
    bottom = track_window.PaneControl(searchDepth=1, ClassName='ScrollViewer').BoundingRectangle.bottom
    switch_button = channel_pane.ButtonControl(searchDepth=2, AutomationId='switchSingerButton')
    while True:
        if switch_button.BoundingRectangle.bottom < bottom:
            switch_button.DoubleClick(simulateMove=False)
            break
        else:
            track_window.MoveCursorToMyCenter(simulateMove=False)
            win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, 0, 0, -500)
    choose_singer(name=name)


if __name__ == '__main__':
    # demo: 导出同一个工程文件的不同歌手演唱音频（公测版歌手不全无法使用，目前仅可运行在 2.0.0 beta2 版本上）
    singers = ['陈水若', '方念', '果妹', '小傻']  # 需要导出的所有歌手
    path = r'PATH_TO_PROJECT'  # 工程文件存放路径
    prefix = '示例'  # 各歌手演唱音频的公共前缀。本例中保存为“示例 - 陈水若.mp3”，“示例 - 方念.mp3”等
    start_xstudio(engine=r'E:\YQ数据空间\YQ实验室\实验室：XStudioSinger\内测\XStudioSinger_2.0.0_beta2.exe', project=path)
    for s in singers:
        switch_singer(track=1, name=s)
        export_project(title=f'{prefix} - {s}')
    quit_xstudio()
