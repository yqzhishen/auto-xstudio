# Auto X Studio

X Studio · 歌手 UI 自动化  |  UI Automation for X Studio Singer



## 项目简介与功能说明  |  Introduction & Features

本项目基于 [Python-UIAutomation-for-Windows](https://github.com/yinkaisheng/Python-UIAutomation-for-Windows) 开发，用于自动控制 X Studio · 歌手软件完成若干常用操作。

功能大致包括：

- 启动和退出 X Studio
- 新建、打开、保存、导出工程

可能的应用场景：

- 批量导出若干份工程
- 导出一个工程的若干版本
- 批量编辑并另存为工程

欢迎加入仓库或提交 Pull requests。若发现 bug，欢迎在本仓库提交 Issue，或[联系作者](https://space.bilibili.com/102844209)。



## 使用方法  |  Usage

模块中定义的函数可供外部调用，可自由组合或添加其他操作（详见参考资料）。每个版本均带有一个典型的实用 Demo，稍加改动即可立即运行。

需求建议与定制：在本仓库提交 issue，或[联系作者](https://space.bilibili.com/102844209)。



## 环境要求  |  Requirements & Dependencies

- Windows 7 及以上操作系统
- X Studio 2.0.0 及以上版本
- Python 3 （3.7.6 和 3.8.1 除外）
- 第三方模块：pywin32, comtypes, typing, uiautomation, colorlog



## 更新日志   |  Update Log

#### 1.0.0 (2022.03.04)

> - 启动和退出 X Studio
> - 新建、打开、保存、导出工程
> - Demo - 批量导出某个文件夹下的所有工程



## 参考资料与相关链接  |  References & Links

- [Python-UIAutomation-for-Windows](https://github.com/yinkaisheng/Python-UIAutomation-for-Windows)
- [UI Automation - Win32 apps | Microsoft Docs](https://docs.microsoft.com/en-us/windows/win32/winauto/entry-uiauto-win32)
- [X Studio · 歌手 - 官方网站](https://singer.xiaoice.com/)
- [作者主页 - YQ之神的个人空间](https://space.bilibili.com/102844209)
- [X Studio · 歌手 - 视频教程](https://www.bilibili.com/video/BV1nk4y117AC)
