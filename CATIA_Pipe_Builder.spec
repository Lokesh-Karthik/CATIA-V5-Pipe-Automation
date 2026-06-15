# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for CATIA Parametric Pipe Builder.

Produces a single-file .exe with:
  - No console window (windowed mode for the Tkinter GUI)
  - config/defaults.yaml bundled alongside
  - All src/ modules included
"""
import os

block_cipher = None

# Project root (where this .spec file lives)
ROOT = os.path.dirname(os.path.abspath(SPEC))

a = Analysis(
    ['main.py'],
    pathex=[ROOT],
    binaries=[],
    datas=[
        # Bundle the config directory so defaults.yaml is available at runtime
        (os.path.join(ROOT, 'config', 'defaults.yaml'), 'config'),
    ],
    hiddenimports=[
        # win32com sub-modules that PyInstaller sometimes misses
        'win32com',
        'win32com.client',
        'win32com.server',
        'pythoncom',
        'pywintypes',
        'yaml',
        # All src modules explicitly
        'src',
        'src.gui_app',
        'src.gui_pipeline',
        'src.centerline_builder_v2',
        'src.pipe_body_builder',
        'src.catia_connection',
        'src.solid_converter',
        'src.surface_builder',
        'src.utils',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Not needed — reduce exe size
        'matplotlib', 'scipy', 'pandas', 'PIL', 'numpy',
        'pytest', 'unittest',
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
    name='CATIA_Pipe_Builder',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # No console window — Tkinter GUI only
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
