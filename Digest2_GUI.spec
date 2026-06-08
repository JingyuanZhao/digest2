# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# 收集所有digest2的文件
datas = [
    ('digest2_icon.ico', '.'),
    ('digest2_icon_256.png', '.'),
]
datas += collect_data_files('digest2', include_py_files=True)

# 收集所有digest2的子模块
hiddenimports = ['digest2'] + collect_submodules('digest2')

# 添加docx相关
hiddenimports += [
    'docx',
    'docx.api',
    'docx.document',
    'docx.table',
    'docx.oxml',
    'docx.oxml.table',
    'docx.oxml.text',
    'docx.image',
    'docx.package',
    'docx.opc',
    'docx.dml',
    'docx.drawing',
    'docx.enum',
    'docx.shape',
    'docx.section',
    'docx.settings',
    'docx.types',
    'docx.shared',
    'docx.blkcntnr',
    'docx.comments',
]

a = Analysis(
    ['digest2_gui.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['PyQt5', 'PySide6', 'PySide2', 'matplotlib', 'pytest', 'sphinx', 'IPython', 'pandas', 'scipy'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='Asterorbit_GUI',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['digest2_icon.ico'],
)