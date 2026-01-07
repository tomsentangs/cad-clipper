# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = [('C:\\Users\\13987\\.conda\\envs\\cad_clip\\Library\\bin', 'osgeo'),('C:\\Users\\13987\\.conda\\envs\\cad_clip\\Library\\share\\gdal', 'gdal_data'), ('C:\\Users\\13987\\.conda\\envs\\cad_clip\\Library\\share\\proj', 'proj_data')]
binaries = []
hiddenimports = ['osgeo', 'osgeo.gdal', 'osgeo._gdal', 'osgeo._osr', 'osgeo._ogr', 'fiona', 'geopandas', 'shapely', 'pyproj', 'ezdxf', 'flask']
tmp_ret = collect_all('osgeo')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('fiona')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('shapely')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('pyproj')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['clip_dxf.py'],
    pathex=[],
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
    a.binaries,
    a.datas,
    [],
    name='CADClipService',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
