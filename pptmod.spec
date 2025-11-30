# -*- mode: python ; coding: utf-8 -*-

datas = [('pptmodconfig.json', '.')]
binaries = []
hiddenimports = ['wx.grid', 'win32com.client', 'pptx']


a = Analysis(
    ['gui.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Exclude unused wx modules
        'wx.lib.pubsub', 'wx.lib.agw', 'wx.lib.plot', 'wx.lib.floatcanvas',
        'wx.html2', 'wx.media', 'wx.richtext', 'wx.ribbon', 'wx.stc',
        'wx.py', 'wx.tools', 'wx.lib.ogl', 'wx.propgrid', 'wx.aui',
        'wx.lib', 'wx.dataview', 'wx.glcanvas', 'wx.xml', 'wx.xrc',
        # Exclude other GUI frameworks
        'tkinter', 'PyQt5', 'PyQt6', 'PySide2', 'PySide6',
        # Exclude data science libraries
        'matplotlib', 'numpy', 'scipy', 'pandas', 'IPython', 'jupyter',
        # Exclude test frameworks
        'unittest', 'pytest', 'nose', 'doctest',
        # Exclude unused stdlib modules
        'pdb', 'pydoc', 'pydoc_data', 'test', 'tests',
        'distutils', 'setuptools', 'pip', 'wheel',
        # Exclude lxml optional dependencies
        'lxml.cssselect', 'lxml.html5lib', 'bs4', 'BeautifulSoup',
        'html5lib', 'cssselect', 'cython',
        # Exclude PIL optional modules
        'PIL.ImageQt', 'PIL.FpxImagePlugin', 'PIL.MicImagePlugin',
    ],
    noarchive=False,
    optimize=2,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='pptmod',
    debug=False,
    bootloader_ignore_signals=False,
    strip=True,
    upx=True,
    upx_exclude=[
        'vcruntime140.dll',
        'python313.dll',
    ],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='NONE',
)
