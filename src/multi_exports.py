from core import *


if __name__ == '__main__':
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
