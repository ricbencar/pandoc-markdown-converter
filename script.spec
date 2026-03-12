# -*- mode: python ; coding: utf-8 -*-
from __future__ import annotations

from pathlib import Path
from PyInstaller.utils.hooks import collect_submodules, copy_metadata

block_cipher = None

project_dir = Path(SPEC).resolve().parent
script_file = project_dir / "script.py"
icon_file = project_dir / "app.ico"
vendor_dir = project_dir / "vendor"
runtime_hook_file = project_dir / "rthook_path_env.py"


def tree(src: Path, dst_prefix: str):
    items = []
    if not src.exists():
        return items
    for p in src.rglob("*"):
        if p.is_file():
            rel_parent = p.parent.relative_to(src)
            dest = str(Path(dst_prefix) / rel_parent).replace("\\", "/")
            items.append((str(p), dest))
    return items


def safe_copy_metadata(package_name: str):
    try:
        return copy_metadata(package_name)
    except Exception:
        return []


datas = []
datas += safe_copy_metadata("pypandoc")
datas += safe_copy_metadata("pypdf")

datas += tree(vendor_dir / "pandoc", "vendor/pandoc")
datas += tree(vendor_dir / "pandoc-data", "vendor/pandoc-data")
datas += tree(vendor_dir / "tex", "vendor/tex")

hiddenimports = []
hiddenimports += collect_submodules("pypdf")
hiddenimports += collect_submodules("pypandoc")
hiddenimports += collect_submodules("tkinter")

excludes = [
    "matplotlib",
    "numpy",
    "pandas",
    "IPython",
    "pytest",
]

runtime_hooks = [str(runtime_hook_file)] if runtime_hook_file.exists() else []


a = Analysis(
    [str(script_file)],
    pathex=[str(project_dir)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=runtime_hooks,
    excludes=excludes,
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="PandocMarkdownConverter",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(icon_file) if icon_file.exists() else None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="PandocMarkdownConverter",
)
