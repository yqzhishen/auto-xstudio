import sys

sys.path.append('core')

from core import engine, projects, tracks


if __name__ == '__main__':
    # demo: 导出同一个工程文件的不同歌手演唱音频（公测版歌手不全无法使用，目前仅可运行在 2.0.0 beta2 版本上）
    singers = ['陈水若', '方念', '果妹', '小傻']  # 需要导出的所有歌手
    path = r'..\泡沫.svip'  # 工程文件存放路径
    prefix = '示例'  # 各歌手演唱音频的公共前缀。本例中保存为“示例 - 陈水若.mp3”，“示例 - 方念.mp3”等
    engine.start_xstudio(engine=r'E:\YQ数据空间\YQ实验室\实验室：XStudioSinger\内测\XStudioSinger_2.0.0_beta2.exe', project=path)
    for s in singers:
        tracks.Track(1).switch_singer(singer=s)
        projects.export_project(title=f'{prefix} - {s}')
    engine.quit_xstudio()
