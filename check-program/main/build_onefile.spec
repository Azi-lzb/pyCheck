# PyInstaller spec: 单文件 exe，控制台程序（精简体积）
# 使用: pyinstaller build_onefile.spec（在 check-program/main 目录下执行）

# -*- mode: python ; coding: utf-8 -*-

# 排除本程序未用的大型库，减小 exe 体积（仅依赖 openpyxl + 标准库 + tkinter）
EXCLUDES = [
    'numpy', 'numpy.*', 'PIL', 'PIL.*', 'matplotlib', 'matplotlib.*',
    'pandas', 'pandas.*', 'scipy', 'scipy.*', 'cv2', 'sklearn',
    'IPython', 'jupyter', 'jupyter_core', 'notebook', 'qtpy', 'PyQt5', 'PyQt6',
    'pytest', 'setuptools', 'pkg_resources', 'distutils',
]

a = Analysis(
    ['build_strategy1.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['openpyxl', 'openpyxl.cell._writer', 'tkinter'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=EXCLUDES,
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
    name='策略一核查',
    debug=False,
    bootloader_ignore_signals=False,
    strip=True,   # 去掉符号表，减小体积
    upx=True,     # 用 UPX 压缩（需本机安装 UPX 才生效）
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # 保留控制台，方便看打印与交互
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
