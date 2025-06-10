# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['extracao_itens-sei.py'],
    pathex=[],
    binaries=[],
    datas=[('C:\\python\\extracao_itens-sei\\excel\\', 'excel'), ('C:\\python\\extracao_itens-sei\\chromedriver-win64\\*', 'chromedriver-win64'), ('C:\\python\\extracao_itens-sei\\login_sei.py', '.'), ('C:\\python\\extracao_itens-sei\\README.pdf', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='extracao_itens-sei',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['C:\\python\\extracao_itens-sei\\icon\\download_sei.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='extracao_itens-sei',
)
