# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(['audit_helper.py'],
             pathex=['.'],
             binaries=[],
             datas=[],
             hiddenimports=['tkinter'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)

a.datas += [("combine_icon.png", "."), ("gui.ui", ".")]

pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)

exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='audit_helper',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          upx_include=[],
          runtime_tmpdir=None,
          console=True,
          icon='combine_icon.png')

