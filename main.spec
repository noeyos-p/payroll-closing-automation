# PyInstaller build spec for 급여대장 자동 변환.
#
# Build (run ON the target OS — PyInstaller does not cross-compile):
#   pip install -r requirements-build.txt
#   pyinstaller main.spec
#
# Output: dist/PayrollAutomation_Maumson&Care24/PayrollAutomation_Maumson&Care24.exe
# (Windows) or .app (Mac).
# --onedir build (a folder, not a single .exe) — startup is much faster
# than --onefile since nothing needs to be unpacked to a temp folder on
# every launch. Zip the whole dist/PayrollAutomation_Maumson&Care24/ folder
# to hand it to someone else; they unzip it and run the .exe inside, no
# Python install required.
#
# The build output name is ASCII-only on purpose (not the app's Korean
# window title) — the sister project (wehago-payroll-tool-py) hit a real
# bug where a Korean folder/exe name broke cmd.exe's own batch-file
# parsing depending on the Windows system's active codepage. The app's
# actual window title (ui/main_window.py) is unaffected and stays Korean.

# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_data_files

# pyhwp's HTML transform reads its .xsl stylesheets from hwp5/xsl/ at
# runtime (core/hwp_reader.py calls hwp5.hwp5html.HTMLTransform directly,
# not the `hwp5html` console script — see that module's docstring for why)
# — these are non-.py resource files PyInstaller won't pick up on its own.
hwp5_datas = collect_data_files("hwp5")

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=hwp5_datas + [
        ('RELEASE_README.txt', '.'),
        ('assets/fonts', 'assets/fonts'),
        ('assets/app_icon.ico', 'assets'),
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
# Bundled tool runtimes on PATH can supply incompatible Windows/ICU DLLs.
# Refuse such a build; run PyInstaller with those runtime paths removed.
foreign_dlls = [entry[0] for entry in a.binaries if 'codex-runtimes' in str(entry[1]).lower()]
if foreign_dlls:
    raise RuntimeError('Unexpected tool-runtime DLLs; clean PATH before building: ' + ', '.join(foreign_dlls))
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='PayrollAutomation_Maumson&Care24',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # GUI app — no console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='assets/app_icon.ico',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='PayrollAutomation_Maumson&Care24',
)
