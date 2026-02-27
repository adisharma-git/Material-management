# pharmacy_inventory.spec
# PyInstaller spec – works for both Windows (.exe) and macOS (.app)

import sys
import os
from PyInstaller.building.build_main import Analysis, PYZ, EXE, COLLECT, BUNDLE

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        # Include the src package
        ('src/*.py', 'src'),
    ],
    hiddenimports=[
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.utils',
        'pandas',
        'numpy',
        'math',
        'threading',
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'src.script1_global_stock',
        'src.script2_main_store_stock',
        'src.script4_inventory_calc',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'scipy',
        'PIL',
        'IPython',
        'jupyter',
        'notebook',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='PharmacyInventory',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,           # No console window (GUI only)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # Windows icon (optional – place icon.ico in repo root)
    # icon='assets/icon.ico',
)

# ── macOS .app bundle ─────────────────────────────────────────────────────────
if sys.platform == 'darwin':
    app = BUNDLE(
        exe,
        name='PharmacyInventory.app',
        # icon='assets/icon.icns',
        bundle_identifier='com.hospital.pharmacy.inventory',
        info_plist={
            'NSHighResolutionCapable': True,
            'CFBundleShortVersionString': '1.0.0',
            'CFBundleVersion': '1.0.0',
            'CFBundleName': 'Pharmacy Inventory',
            'NSHumanReadableCopyright': 'Hospital Pharmacy',
        },
    )
