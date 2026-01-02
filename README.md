# Contingency-Comparator-v2

py -m pip install --upgrade pywin32 pywin32-ctypes `
  --trusted-host pypi.org `
  --trusted-host files.pythonhosted.org `
  --trusted-host pypi.python.org

py -m pywin32_postinstall -install

py -m PyInstaller --onefile --windowed ContingencyComparaterV2.py

py -m PyInstaller --noconfirm --clean --onedir --windowed --name ContingencyComparatorV2 `
  --icon assets\app.ico `
  --add-data "assets\app.ico;assets" `
  --add-data "assets\app_256.png;assets" `
  main.py



fast onedir


a = Analysis(
    ['ContingencyComparaterV2.py'],
    pathex=['.'],
    binaries=[],
    datas=[('assets', 'assets')],
    hiddenimports=[],
    excludes=['matplotlib', 'IPython'],
    optimize=2,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='ContingencyComparaterV2',
    icon='assets/app.ico',
    console=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    name='ContingencyComparaterV2',
)



clean onedir


a = Analysis(
    ['ContingencyComparaterV2.py'],
    pathex=['.'],
    binaries=[],
    datas=[('assets', 'assets')],
    hiddenimports=[],
    excludes=['matplotlib'],
    optimize=2,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='ContingencyComparaterV2',
    icon='assets/app.ico',
    console=False,
    upx=False,
)





# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['ContingencyComparaterV2\\ContingencyComparaterV2.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
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
    name='ContingencyComparaterV2',
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
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='ContingencyComparaterV2',
)





cd dist\ContingencyComparaterV2
gci _internal -Directory | % {
  $size = (gci $_.FullName -Recurse -File | measure Length -Sum).Sum
  [pscustomobject]@{Name=$_.Name; MB=[math]::Round($size/1MB,1)}
} | sort MB -desc | select -first 25