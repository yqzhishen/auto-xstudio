import uiautomation as auto

from exception import AutoXStudioException
import log

logger = log.logger


def verify_startup():
    """
    验证 X Studio 是否成功启动。
    """
    warning_window = auto.WindowControl(searchDepth=1, Name='提示')
    if warning_window.Exists(maxSearchSeconds=3):
        warning = warning_window.TextControl(searchDepth=1, AutomationId='Tbx').Name
        warning_window.ButtonControl(searchDepth=1, AutomationId='OkBtn').Click(simulateMove=False)
        logger.error(warning)
        raise AutoXStudioException()


def verify_opening(base):
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
            raise AutoXStudioException()
        else:
            warning_window.ButtonControl(searchDepth=1, Name='确定').Click(simulateMove=False)
            base.WindowControl(searchDepth=1, Name='X Studio').ButtonControl(searchDepth=1, AutomationId='OkBtn').Click(simulateMove=False)
            base.WindowControl(searchDepth=1, Name='X Studio').ButtonControl(searchDepth=1, AutomationId='btnClose').Click(simulateMove=False)
            logger.error(warning)
            raise AutoXStudioException()


def verify_updates():
    """
    验证 X Studio 更新，并关闭更新提示窗。
    """
    update_window = auto.WindowControl(searchDepth=2, Name='检测到新版本')
    if update_window.Exists(maxSearchSeconds=3):
        update_window.ButtonControl(searchDepth=1, AutomationId='btnClose').Click(simulateMove=False)
