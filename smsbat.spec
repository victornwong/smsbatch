# -*- mode: python -*-
a = Analysis(['smsbat.py'],
             pathex=['Z:\\smsbat'],
             hiddenimports=[],
             hookspath=None,
             runtime_hooks=None)
pyz = PYZ(a.pure)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='smsbat.exe',
          debug=False,
          strip=None,
          upx=True,
          console=False )
