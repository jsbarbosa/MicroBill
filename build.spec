# -*- mode: python -*-

block_cipher = None

a = Analysis(['run.py'],
             pathex=[],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes = ["microbill.config", "microbill.login"],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)

pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)

exe = EXE(pyz,
          a.scripts,
          exclude_binaries=True,
          name='MicroBill',
          debug=False,
          strip=False,
          upx=True,
          console=False,
          icon='microbill/icon.ico',
          version='microbill/fileversion.txt')

coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               name='MicroBill')
