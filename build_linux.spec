# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec file for Linux

block_cipher = None

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('templates', 'templates'),
        ('static', 'static'),
    ],
    hiddenimports=[
        'flask',
        'PIL',
        'PIL.Image',
        'fitz',
        'img2pdf',
        'docx',
        'openpyxl',
        'pptx',
        'reportlab',
        'reportlab.lib',
        'reportlab.platypus',
        'reportlab.pdfbase',
        'reportlab.pdfbase.ttfonts',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'torch', 'torchvision', 'tensorflow', 'keras',
        'numpy.distutils', 'scipy', 'pandas', 'matplotlib',
        'sklearn', 'cv2', 'ray', 'transformers', 'datasets',
        'pytest', 'sphinx', 'IPython', 'jupyter',
        'PyQt5', 'PyQt6', 'PySide2', 'PySide6', 'tkinter',
        'boto3', 'botocore', 'google', 'grpc',
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
    name='everything2pdf',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # Linux下保留控制台以便查看输出
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
