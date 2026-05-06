# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = [('_internal/Fonts', 'Fonts')]
binaries = []
hiddenimports = ['openpyxl', 'openpyxl.workbook', 'openpyxl.worksheet', 'xlrd', 'et_xmlfile', 'lxml', 'lxml.etree', 'pandas', 'pandas._libs.tslibs.np_datetime', 'pandas._libs.tslibs.nattype', 'numpy', 'tkinter', 'tkinter.messagebox', 'tkinter.filedialog', 'PIL', 'PIL.Image', 'drafthorse', 'pypdf', 'src', 'src.handle_pdf', 'src.handle_zugferd', 'src.handle_girocode', 'src.handle_ini_file', 'src.middleware', 'src.steuerung', 'src.oberflaeche_base', 'src.oberflaeche_excel2zugferd', 'src.oberflaeche_ini', 'src.oberflaeche_steuerung', 'src.oberflaeche_excelpositions', 'src.oberflaeche_excelsteuerung', 'src.invoice_collection', 'src.invoice', 'src.excel_content', 'src.adresse', 'src.constants', 'src.konto', 'src.kunde', 'src.lieferant', 'src.stammdaten', 'src.windowseventlog']
tmp_ret = collect_all('encodings')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('drafthorse')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('src')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['excel2zugferd.py'],
    pathex=['C:\\Users\\Charis\\Projekte\\excel2zugferd'],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
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
    name='excel2zugferd',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='excel2zugferd',
)
