# -*- mode: python ; coding: utf-8 -*-

a = Analysis(['confautomation.py'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=['FixTk', 'tcl', 'tk', '_tkinter', 'tkinter', 'Tkinter'],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data)
exe = EXE(pyz,
          a.scripts,
          [],
          exclude_binaries=True,
          name='confautomation',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=True,
          icon='camera-128.ico')
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='confautomation')
