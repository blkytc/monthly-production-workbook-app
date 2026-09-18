# -*- mode: python ; coding: utf-8 -*-

import sys


app_name = "月度生产表生成器"
a = Analysis(
    ["launcher.py"],
    pathex=["src"],
    binaries=[],
    datas=[("src/monthly_workbook/help.html", "monthly_workbook")],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)
if sys.platform == "darwin":
    exe = EXE(
        pyz, a.scripts, [], exclude_binaries=True, name=app_name, debug=False,
        bootloader_ignore_signals=False, strip=False, upx=True, console=False,
        disable_windowed_traceback=False, argv_emulation=False,
        target_arch=None, codesign_identity=None, entitlements_file=None,
    )
    collected = COLLECT(
        exe, a.binaries, a.datas, strip=False, upx=True, name=app_name,
    )
    app = BUNDLE(
        collected,
        name=f"{app_name}.app",
        icon=None,
        bundle_identifier="com.local.monthlyworkbook",
        info_plist={"NSHighResolutionCapable": True},
    )
else:
    exe = EXE(
        pyz, a.scripts, a.binaries, a.datas, [], name=app_name, debug=False,
        bootloader_ignore_signals=False, strip=False, upx=True, console=False,
        disable_windowed_traceback=False, argv_emulation=False,
        target_arch=None, codesign_identity=None, entitlements_file=None,
    )
