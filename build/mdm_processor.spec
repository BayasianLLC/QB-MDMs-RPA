
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    [r'c:\Users\sabar\Documents\GitHub\QB-MDMs-RPA\src\mdm_processor_ui.py'],
    pathex=[r'c:\Users\sabar\Documents\GitHub\QB-MDMs-RPA'],
    binaries=[],
    datas=[
        (r'c:\Users\sabar\Documents\GitHub\QB-MDMs-RPA\config\credentials.json', 'config'),
                (r'c:\Users\sabar\Documents\GitHub\QB-MDMs-RPA\config\settings.py', 'config'),
                (r'c:\Users\sabar\Documents\GitHub\QB-MDMs-RPA\src\pseg_mdm_processor.py', 'src'),
                (r'c:\Users\sabar\Documents\GitHub\QB-MDMs-RPA\src\pse_mdm_processor.py', 'src'),
                (r'c:\Users\sabar\Documents\GitHub\QB-MDMs-RPA\src\sce_mdm_processor.py', 'src'),
                (r'c:\Users\sabar\Documents\GitHub\QB-MDMs-RPA\src\sdge_mdm_processor.py', 'src')
    ],
    hiddenimports=['pandas', 'office365', 'quickbase_client'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    name='MDM_Processor',
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
    version='version_info.txt',
    icon=None,
)
