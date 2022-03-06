import os
import sys

sys.path.append('core')

from core import engine, projects


if __name__ == '__main__':
    # demo: 导出某文件夹下所有的工程
    path = r'PATH_TO_PROJECTS'
    filelist = []
    for name in os.listdir(path):
        if os.path.isfile(os.path.join(path, name)) and name.endswith('.svip'):
            filelist.append(name)
    num = len(filelist)
    if num > 0:
        engine.start_xstudio(os.path.join(path, filelist[0]))
        projects.export_project(format='wav', samplerate=48000)
        if num > 1:
            projects.open_project(filename=filelist[1], folder=path)
            projects.export_project(format='wav', samplerate=48000)
            for file in filelist[2:]:
                projects.open_project(filename=file)
                projects.export_project(format='wav', samplerate=48000)
        engine.quit_xstudio()
