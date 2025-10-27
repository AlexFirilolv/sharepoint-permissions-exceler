# -*- mode: python ; coding: utf-8 -*-

import sys
import os
from pathlib import Path

# Get the parent directory (project root) - use current working directory approach
project_root = Path(os.getcwd()).parent

# Add project root to sys.path so PyInstaller can find modules
sys.path.insert(0, str(project_root))

a = Analysis(
    [str(project_root / 'gui.py')],
    pathex=[str(project_root)],
    binaries=[],
    datas=[
        # Include .env.example for reference
        (str(project_root / '.env.example'), '.'),
    ],
    hiddenimports=[
        # Explicitly include modules that might not be auto-detected
        'pandas',
        'openpyxl',
        'msal',
        'requests',
        'dotenv',
        'PyQt6.QtCore',
        'PyQt6.QtWidgets',
        'PyQt6.QtGui',
        'main',  # Our main module
        # Fix numpy/pandas circular import issues
        'numpy.random._common',
        'numpy.random._bounded_integers',
        'numpy.random._mt19937',
        'numpy.random._pcg64',
        'numpy.random._philox',
        'numpy.random._sfc64',
        'numpy.random.bit_generator',
        'numpy.random.mtrand',
        'pandas._libs.tslibs.base',
        'pandas._libs.tslibs.np_datetime',
        'pandas._libs.tslibs.nattype',
        'pandas._libs.tslibs.timezones',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Exclude unnecessary modules to reduce file size
        'tkinter',
        'matplotlib',
        'IPython',
        'jupyter',
        'notebook',
        'qtconsole',
        'sphinx',
        'pytest',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='SharePoint-Permissions-Exceler',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # No console window for GUI app
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # Add icon path here if you have one
    version=None,  # Add version info here if needed
)