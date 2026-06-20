# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['MoonOrangeBulletinPDFDownloader.py'],
    pathex=[],
    binaries=[],
    datas=[('user_penguin.ico', '.')],
    hiddenimports=['requests'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib', 'numpy', 'pandas', 'scipy', 'sympy', 'sklearn',
        'IPython', 'jupyter', 'notebook', 'ipykernel', 'nbconvert', 'nbformat',
        'sphinx', 'docutils', 'pygments',
        'tornado', 'zmq', 'jinja2', 'markupsafe',
        'sqlalchemy', 'pymysql', 'sqlite3',
        'lxml', 'xmltodict', 'defusedxml',
        'pytest', 'setuptools', 'pip', 'wheel', 'distutils',
        'pyarrow', 'h5py', 'tables',
        'botocore', 'boto3',
        'selenium', 'webdriver',
        'cv2', 'opencv_python',
        'cryptography', 'OpenSSL',
        'win32com', 'win32evtlog', 'win32pdh', 'win32trace', 'win32ui',
        'pythoncom', 'pywintypes',
        'tkinter.test', 'unittest',
        'asyncio', 'xmlrpc', 'http.server',
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
    name='MoonOrangeBulletinPDFDownloader',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
