from cx_Freeze import setup, Executable


include_files = ["settings.json", "scania-symbol.png"]
includes = []
excludes = []
packages = []

setup(name = 'BubblesToXLSX',
            version = '1.0',
            description = 'Bubbles to Excel',
            options = {'build_exe': {'includes': includes, 'excludes': excludes, 'packages': packages, 'include_files': include_files}},
            executables = [Executable("main.py", base= None, target_name= "BubblesToXLSX.exe")])