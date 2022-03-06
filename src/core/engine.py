import os
import winreg

import uiautomation as auto

import log
import singers
import verify

logger = log.logger


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
        verify.verify_startup()
        verify.verify_opening(auto)
        logger.info('启动 X Studio 并打开工程：%s。' % project)
        verify.verify_updates()
    else:
        if engine:
            os.popen(f'"{engine}"')
        else:
            os.popen(f'"{find_xstudio()}"')
        verify.verify_startup()
        auto.WindowControl(searchDepth=1, Name='X Studio').TextControl(searchDepth=2, Name='开始创作').Click(simulateMove=False)
        singers.choose_singer(name=singer)
        logger.info('启动 X Studio 并创建空白工程，初始歌手：%s。' % singer)
        verify.verify_updates()


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


if __name__ == '__main__':
    start_xstudio(singer='陈水若')
