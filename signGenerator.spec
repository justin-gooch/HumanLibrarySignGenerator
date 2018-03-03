# -*- mode: python -*-
import sys
from os import path
site_packages = next(p for p in sys.path if 'site-packages' in p)
block_cipher = None

block_cipher = None


a = Analysis(['signGenerator.py'],
             pathex=['C:\\Users\\Justin\\Documents\\GitHub\\HumanLibrarySignGenerator'],
             binaries=[],
             datas=[(path.join(site_packages, 'docx','templates'), 'docx/templates')],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          exclude_binaries=True,
          name='signGenerator',
          debug=False,
          strip=False,
          upx=True,
          console=True )
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               name='signGenerator')
