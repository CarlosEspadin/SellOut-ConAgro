# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['C:\\Users\\carlo\\OneDrive\\Documentos\\SellOut ConAgro\\SellOut.py'],
    pathex=[],
    binaries=[],
    datas=[('C:\\Users\\carlo\\OneDrive\\Documentos\\SellOut ConAgro\\agriculture_6739552.png', '.'), ('C:\\Users\\carlo\\OneDrive\\Documentos\\SellOut ConAgro\\ConAgro.ico', '.'), ('C:\\Users\\carlo\\OneDrive\\Documentos\\SellOut ConAgro\\ConAgro_icon_big.png', '.'), ('C:\\Users\\carlo\\OneDrive\\Documentos\\SellOut ConAgro\\ConAgro_icon_small.png', '.'), ('C:\\Users\\carlo\\OneDrive\\Documentos\\SellOut ConAgro\\home_2115185.png', '.'), ('C:\\Users\\carlo\\OneDrive\\Documentos\\SellOut ConAgro\\menu.png', '.'), ('C:\\Users\\carlo\\OneDrive\\Documentos\\SellOut ConAgro\\onboarding_14753055.png', '.'), ('C:\\Users\\carlo\\OneDrive\\Documentos\\SellOut ConAgro\\README.md', '.'), ('C:\\Users\\carlo\\OneDrive\\Documentos\\SellOut ConAgro\\review-document_14752610.png', '.'), ('C:\\Users\\carlo\\OneDrive\\Documentos\\SellOut ConAgro\\setting_1146744.png', '.'), ('C:\\Users\\carlo\\OneDrive\\Documentos\\SellOut ConAgro\\workflow_14254834.png', '.')],
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
    name='SellOut',
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
    icon=['C:\\Users\\carlo\\OneDrive\\Documentos\\SellOut ConAgro\\ConAgro.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='SellOut',
)
