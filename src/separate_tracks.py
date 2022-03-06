import sys

sys.path.append('core')

from core import engine, projects, tracks

if __name__ == '__main__':
    path = r'..\demo\separate_tracks\assets\示例.svip'
    prefix = '示例'
    engine.start_xstudio(engine=r'C:\Users\YQ之神\AppData\Local\warp\packages\XStudioSinger_2.0.0_beta2.exe\XStudioSinger.exe', project=path)
    for track in tracks.enum_tracks()[::-1]:
        track.set_solo(True)
        projects.export_project(title=f'{prefix}_轨道{track.index}')
    engine.quit_xstudio()
