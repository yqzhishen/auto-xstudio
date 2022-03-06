import os

import uiautomation as auto

import keybd
import log
import verify
import singers

logger = log.logger


def new_project(singer: str = None):
    """
    新建工程。X Studio 必须已处于启动状态。
    :param singer: 可指定新工程的初始歌手
    """
    keybd.key_down(17)
    keybd.key_press(78)
    keybd.key_up(17)
    confirm_window = auto.WindowControl(searchDepth=1, Name='X Studio')
    if confirm_window.Exists(maxSearchSeconds=1):
        confirm_window.ButtonControl(searchDepth=1, AutomationId='NoBtn').Click(simulateMove=False)
    if singer:
        track_window = auto.WindowControl(searchDepth=1, RegexName='X Studio .*').CustomControl(searchDepth=1, ClassName='TrackWin')
        track_window.CustomControl(searchDepth=2, ClassName='TrackChannelControlPanel').ButtonControl(searchDepth=2, AutomationId='switchSingerButton').DoubleClick(simulateMove=False)
        singers.choose_singer(singer)
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
    keybd.key_down(17)
    keybd.key_press(79)
    keybd.key_up(17)
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
    verify.verify_opening(main_window)
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
    else:
        title = setting_window.EditControl(searchDepth=1, AutomationId='FileNameTbx').GetValuePattern().Value
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
        keybd.key_down(17)
        keybd.key_press(83)
        keybd.key_up(17)
        logger.info('保存工程。')
    else:
        if folder:
            project = os.path.join(folder, filename)
        else:
            project = filename
        keybd.key_down(17)
        keybd.key_down(16)
        keybd.key_press(83)
        keybd.key_up(16)
        keybd.key_up(17)
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
