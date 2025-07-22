import PyInstaller.__main__

PyInstaller.__main__.run([
    'main2.py',
    '--onefile',
    '--clean',
    '--collect-submodules=openpyxl',
    '--collect-submodules=xlsxwriter',
    '--noconsole',
    '--windowed',
    '--icon=blueprint.ico',   # ✅ fixed quote
    '--name=AddUnused'        # ✅ sets the final exe name
])