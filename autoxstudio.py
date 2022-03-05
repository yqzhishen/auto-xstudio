# UI AUTOMATION FOR X STUDIO
# AUTHOR: YQ之神
# VERSION: 1.0.0 (2022.3.4)

import logging
import os
import winreg

import uiautomation as auto

logger = logging.getLogger()


def _init():
    import colorlog
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
    warning_window = auto.WindowControl(searchDepth=1, Name='提示', maxSearchSeconds=3)
    if warning_window.Exists():
        warning = warning_window.TextControl(searchDepth=1, AutomationId='Tbx').Name
        warning_window.ButtonControl(searchDepth=1, AutomationId='OkBtn').Click(simulateMove=False)
        logger.error(warning)
        exit(1)


def _verify_opening(base):
    """
    验证工程是否成功打开。
    """
    warning_window = base.WindowControl(searchDepth=1, ClassName='#32770', maxSearchSeconds=1)
    if warning_window.Exists():
        warning = warning_window.TextControl(searchDepth=2).Name.replace('\r\n', ' ').replace('。 ', '。')
        if warning.startswith('无法读取伴奏文件'):
            logger.warning('已自动忽略：无法读取伴奏文件。')
            while warning_window.Exists():
                warning_window.ButtonControl(searchDepth=1, Name='确定').Click(simulateMove=False)
                warning_window = auto.WindowControl(searchDepth=2, ClassName='#32770', maxSearchSeconds=1)
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


def find_xstudio() -> str:
    """
    根据注册表查找 X Studio 主程序路径。
    :return: XStudioSinger.exe 的路径
    """
    key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 'SOFTWARE\\Classes\\svipfile\\shell\\open\\command')
    value = winreg.QueryValueEx(key, '')
    return value[0].split('"')[1]


def start_xstudio(project: str = None):
    """
    启动 X Studio。
    :param project: 启动时需要打开的工程文件路径，默认打开空白工程
    """
    if project:
        if not os.path.exists(project):
            logger.error('文件不存在。')
            exit(1)
        if not project.endswith('.svip'):
            logger.error('不是一个可打开的 X Studio 工程 (.svip) 文件。')
            exit(1)
        os.startfile(project.replace('/', '\\'))
        _verify_startup()
        _verify_opening(auto)
        logger.info('启动 X Studio 并打开工程：%s。' % project)
    else:
        os.startfile(find_xstudio())
        _verify_startup()
        auto.WindowControl(searchDepth=1, Name='X Studio').TextControl(searchDepth=2, Name='开始创作').Click(simulateMove=False)
        singer_market = auto.WindowControl(searchDepth=2, Name='歌手市场')
        singer_market.HyperlinkControl(searchDepth=9, Name='全部歌手').Click(simulateMove=False)
        singer_market.TextControl(searchDepth=16, Name='陈水若').Click(simulateMove=False)
        singer_market.ButtonControl(searchDepth=17, Name='选中').Click(simulateMove=False)
        logger.info('启动 X Studio 并创建空白工程。')


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


def new_project():
    """
    新建工程。X Studio 必须已处于启动状态。
    """
    main_window = auto.WindowControl(searchDepth=1, RegexName='X Studio .*')
    main_window.MenuItemControl(searchDepth=2, Name='菜单').Click(simulateMove=False)
    main_window.MenuItemControl(searchDepth=3, Name='文件').Click(simulateMove=False)
    main_window.MenuItemControl(searchDepth=4, Name='新建工程').Click(simulateMove=False)
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
            logger.error('文件不存在。')
            exit(1)
    else:
        project = filename
    if not filename.endswith('.svip'):
        logger.error('不是一个可打开的 X Studio 工程 (.svip) 文件。')
        exit(1)
    main_window = auto.WindowControl(searchDepth=1, RegexName='X Studio .*')
    main_window.MenuItemControl(searchDepth=2, Name='菜单').Click(simulateMove=False)
    main_window.MenuItemControl(searchDepth=3, Name='文件').Click(simulateMove=False)
    main_window.MenuItemControl(searchDepth=4, Name='打开工程').Click(simulateMove=False)
    open_window = main_window.WindowControl(searchDepth=1, Name='打开文件')
    open_window.EditControl(searchDepth=3, Name='文件名(N):').GetValuePattern().SetValue(project)
    open_window.ButtonControl(searchDepth=1, Name='打开(O)').Click(simulateMove=False)
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
        message_window = auto.WindowControl(searchDepth=2, ClassName='#32770', maxSearchSeconds=0.5)
        if message_window.Exists():
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
    logger.info('导出工程：%s, %s 格式, 采样率 %d Hz。' % (title, format, samplerate))


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
    main_window = auto.WindowControl(searchDepth=1, RegexName='X Studio .*')
    main_window.MenuItemControl(searchDepth=2, Name='菜单').Click(simulateMove=False)
    main_window.MenuItemControl(searchDepth=3, Name='文件').Click(simulateMove=False)
    if not filename:
        main_window.MenuItemControl(searchDepth=4, Name='保存工程').Click(simulateMove=False)
        logger.info('保存工程。')
    else:
        if folder:
            project = os.path.join(folder, filename)
        else:
            project = filename
        main_window.MenuItemControl(searchDepth=4, Name='工程另存为').Click(simulateMove=False)
        save_window = main_window.WindowControl(searchDepth=1, Name='另存为')
        save_window.EditControl(searchDepth=6, Name='文件名:').GetValuePattern().SetValue(project)
        save_window.ButtonControl(searchDepth=1, Name='保存(S)').Click(simulateMove=False)
        confirm_window = save_window.WindowControl(searchDepth=1, ClassName='#32770', maxSearchSeconds=1)
        if confirm_window.Exists():
            warning = confirm_window.TextControl(searchDepth=2).Name
            if warning.endswith('是否替换它?'):
                confirm_window.ButtonControl(searchDepth=1, Name='是(Y)').Click(simulateMove=False)
            else:
                confirm_window.ButtonControl(searchDepth=2, Name='确定').Click(simulateMove=False)
                save_window.ButtonControl(searchDepth=1, Name='取消').Click(simulateMove=False)
                logger.error(warning.replace('\r\n', ' ').replace('。 ', '。'))
                exit(1)
        logger.info('另存为工程。')


if __name__ == '__main__':
    _init()
    # demo: 导出某文件夹下所有的工程
    path = r'PATH_TO_PROJECTS'
    filelist = []
    for name in os.listdir(path):
        if os.path.isfile(os.path.join(path, name)) and name.endswith('.svip'):
            filelist.append(name)
    num = len(filelist)
    if num > 0:
        start_xstudio(os.path.join(path, filelist[0]))
        export_project(format='wav', samplerate=48000)
        if num > 1:
            open_project(filename=filelist[1], folder=path)
            export_project(format='wav', samplerate=48000)
            for file in filelist[2:]:
                open_project(filename=file)
                export_project(format='wav', samplerate=48000)
        quit_xstudio()
