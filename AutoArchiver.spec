import tkinterdnd2
from os.path import join, dirname

tkdnd_path = join(dirname(tkinterdnd2.__file__), 'tkdnd')

a = Analysis(
    ['batch_archiver.py'],
    pathex=[],
    binaries=[],
    datas=[('sparkgy.ico', '.'), (tkdnd_path, 'tkinterdnd2/tkdnd')],
    hiddenimports=['tkinterdnd2'],
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
    a.binaries,
    a.datas,
    [],
    name='AutoArchiver',
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
    icon=['sparkgy.ico'],
)
