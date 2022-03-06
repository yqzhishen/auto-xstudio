# Auto X Studio

X Studio · 歌手 UI 自动化  |  UI Automation for X Studio Singer



## 项目简介与功能说明  |  Introduction & Features

本项目基于 [Python-UIAutomation-for-Windows](https://github.com/yinkaisheng/Python-UIAutomation-for-Windows) 开发，用于自动控制 X Studio · 歌手软件完成若干常用操作。

功能大致包括：

- 启动和退出 X Studio
- 新建、打开、保存、导出工程
- 为某条轨道切换歌手
- 静音、独奏某条轨道

可能的应用场景：

- 批量导出若干份工程
- 分轨导出一个工程
- 导出一个工程的若干版本
- 批量编辑并另存为工程
- **工程在线试听**（欢迎网站站长合作！）

欢迎加入仓库或提交 pull requests。若发现 bug 或有功能建议，欢迎在本仓库提交 issue，或[联系作者](https://space.bilibili.com/102844209)。



## 环境要求  |  Requirements & Dependencies

- Windows 7 及以上操作系统
- X Studio · 歌手 2.0.0 及以上版本
- Python 3（除 3.7.6 和 3.8.1 外）
- 第三方模块：comtypes, typing, uiautomation, colorlog



## 使用方法  |  How to Use

发行版中包含若干个 demo，安装好所需依赖后即可运行 main.py 查看效果。

UI 自动化执行的成功与否受到系统流畅度等客观因素影响。本模块考虑了一些异常特殊情况的处理，但为了确保执行顺畅、无误，使用时请关闭不需要的应用以保证软件界面不卡顿，并保持网络畅通，防止出现窗口或组件未响应、卡加载、服务器忙碌等意外情况。

模块中定义的函数可供外部调用，可自由组合或添加其他操作（详见代码注释与参考资料）。若需借助本工程实现自定义的自动化流程，需将 src 文件夹复制到自己的项目中，随后可以 import 使用。

不会编程或需要定制自动化流程？在本仓库提交 issue，或[联系作者](https://space.bilibili.com/102844209)。



## 更新日志   |  Update Log

#### 1.0.0 (2022.03.04)

> - 启动和退出 X Studio
> - 新建、打开、保存、导出工程
> - Demo - 批量导出某个文件夹下的所有工程

#### 1.1.0 (2022.03.05)

> - 打开或新建空白工程时，支持指定初始歌手
> - 支持为某条轨道切换歌手
> - 支持指定 X Studio 主程序目录（适用于多版本共存与内测版的情况）
> - 新增若干特殊异常情况的处理
> - 部分步骤采用快捷键，简化流程
> - 缩短等待时间，提升运行速度
> - Demo - 为同一份工程文件导出不同歌手的演唱音频

#### 1.2.0 (2022.03.06)

> - 修复音轨操作不能向上滚动的问题
> - 切换歌手时将打印日志
> - 重构部分代码，优化使用方式

#### 1.3.0 (2022.03.06)

> - 支持静音、取消静音、独奏、取消独奏
> - 调整项目结构，优化部分代码
> - Demo - 分轨导出一份工程文件

#### 1.3.1 (2022.03.06)

> - 解除了对第三方库 pywin32 的依赖



## 参考资料与相关链接  |  References & Links

- [Python-UIAutomation-for-Windows](https://github.com/yinkaisheng/Python-UIAutomation-for-Windows)
- [UI Automation - Win32 apps | Microsoft Docs](https://docs.microsoft.com/en-us/windows/win32/winauto/entry-uiauto-win32)
- [X Studio · 歌手 - 官方网站](https://singer.xiaoice.com/)
- [作者主页 - YQ之神的个人空间](https://space.bilibili.com/102844209)
- [X Studio · 歌手 - 视频教程](https://www.bilibili.com/video/BV1nk4y117AC)
