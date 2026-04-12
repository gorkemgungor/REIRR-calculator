# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# Collect data files for packages that need them
reportlab_datas = collect_data_files('reportlab')
lxml_datas      = collect_data_files('lxml')

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('icon_source.png', '.'),
        ('EN.png', '.'),
    ] + reportlab_datas + lxml_datas,
    hiddenimports=[
        'PyQt6.sip',
        'PyQt6.QtCore',
        'PyQt6.QtGui',
        'PyQt6.QtWidgets',
        'PyQt6.QtPrintSupport',
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.styles.fonts',
        'openpyxl.styles.fills',
        'openpyxl.styles.borders',
        'openpyxl.styles.alignment',
        'reportlab.pdfbase',
        'reportlab.pdfbase.ttfonts',
        'reportlab.pdfbase.pdfmetrics',
        'reportlab.platypus',
        'reportlab.lib',
        'reportlab.lib.styles',
        'reportlab.lib.pagesizes',
        'reportlab.lib.units',
        'lxml',
        'lxml.etree',
        'lxml._elementpath',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='REIR Calculator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='REIR Calculator',
)

app = BUNDLE(
    coll,
    name='REIR Calculator.app',
    icon='AppIcon.icns',
    bundle_identifier='com.reircalculator.app',
    info_plist={
        'NSHighResolutionCapable': True,
        'NSRequiresAquaSystemAppearance': False,
        'CFBundleShortVersionString': '1.1.1',
        'CFBundleVersion': '1.1.1',
        'CFBundleName': 'REIR Calculator',
        'CFBundleDisplayName': 'REIR Calculator',
        'LSMinimumSystemVersion': '11.0',
    },
)
