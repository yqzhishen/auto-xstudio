import sys

sys.path.append('core')

from core import engine, projects, tracks

if __name__ == '__main__':
    path = r'..\demo\separate_tracks\assets\示例.svip'
    prefix = '示例'
    engine.start_xstudio(engine=r'E:\YQ数据空间\YQ实验室\实验室：XStudioSinger\内测\XStudioSinger_2.0.0_beta2.exe', project=path)
    for track in tracks.enum_tracks():
        track.set_solo(True)
        projects.export_project(title=f'{prefix}_轨道{track.index}')
    engine.quit_xstudio()
