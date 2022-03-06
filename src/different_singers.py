import sys

sys.path.append('core')

from core import engine, projects, tracks


if __name__ == '__main__':
    # demo: 导出同一个工程文件的不同歌手演唱音频（公测版歌手不全无法使用，目前仅可运行在 2.0.0 beta2 版本上）
    singers = ['陈水若', '果妹']  # 需要导出的所有歌手
    path = r'..\泡沫.svip'  # 工程文件存放路径
    prefix = '示例'  # 各歌手演唱音频的公共前缀。本例中保存为“示例 - 陈水若.mp3”，“示例 - 果妹.mp3”等
    engine.start_xstudio(engine=r'C:\Users\YQ之神\AppData\Local\warp\packages\XStudioSinger_2.0.0_beta2.exe\XStudioSinger.exe', project=path)
    for s in singers:
        tracks.Track(1).switch_singer(singer=s)
        projects.export_project(title=f'{prefix} - {s}')
    engine.quit_xstudio()
