# -*- mode: python ; coding: utf-8 -*-

import os
import sys

block_cipher = None

added_files = [
    ('frontend', 'frontend'),
    ('source', 'source'),
    ('target', 'target'),
    ('mapping_import_accounting_journal.json', '.'),
    ('mapping_submit_accounting_journal.json', '.'),
    ('HOW_IT_WORKS.md', '.'),
    ('agent.md', '.')
]

a = Analysis(
    ['desktop_app.py'],
    pathex=['.'],
    binaries=[],
    datas=added_files,
    hiddenimports=[
        'clr',
        'pythonnet',
        'webview',
        'webview.platforms.winforms',
        'openpyxl',
        'csv',
        'math',
        'datetime',
        'shutil'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'pandas',
        'numpy',
        'pyarrow',
        'tkinter',
        'sqlite3',
        'unittest',
        'pydoc',
        'torch',
        'matplotlib',
        'scipy',
        'litellm',
        'gradio',
        'werkzeug',
        'PIL',
        'sqlglot',
        'tifffile',
        'datamodel_code_generator',
        'rich',
        'playwright',
        'pytest',
        'IPython',
        'zmq',
        'cryptography',
        'tornado'
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
    name='WorkdayEIBStudio',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

