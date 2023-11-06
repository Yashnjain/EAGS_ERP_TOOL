# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(
    ['EAGS_Quote_Generator.py'],
    pathex=['C:\\Users\\imam.khan\\Documents\\EAGS_New_Test\\venv\\Lib\\site-packages'],
    binaries=[],
    datas=[('biourjaLogo.png', '.'), ('sound1.png', '.'), ('Entry1.png', '.'), ('Entry1New.png', '.'), ('Entry2.png', '.'), ('Entry2New.png', '.'), ('Entry3.png', '.'), ('Entry3New.png', '.'), ('Entry4.png', '.'), ('Entry4New.png', '.'), ('center.png', '.'), ('home(2).png', '.'), ('home(4).png', '.'), ('addRowS.png', '.'), ('addRow2S.png', '.'), ('deleteRowS.png', '.'), ('deleteRow2S.png', '.'), ('previewButtonS.png', '.'), ('previewButton2S.png', '.'), ('submitButtonS.png', '.'), ('submitButton2S.png', '.'), ('reportGenerator1.png', '.'), ('reportGenerator2.png', '.'), ('off.png', '.'), ('on.png', '.'), ('no_order_why1.png', '.'), ('no_order_why2.png', '.'), ('pdfCreator.xlsm', '.'), ('c:\\users\\imam.khan\\appdata\\roaming\\python\\python38\\site-packages\\customtkinter', 'customtkinter\\')],
    hiddenimports=['tkPDFViewer', 'tkcalendar', 'snowflake-connector-python', 'xlwings', 'babel.numbers'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='EAGS_Quote_Generator',
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
    icon='biourjaLogo.ico',
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='EAGS_Quote_Generator',
)
