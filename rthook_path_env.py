from __future__ import annotations

import os
import sys
from pathlib import Path


def _existing_dirs(paths):
    result = []
    seen = set()
    for path in paths:
        try:
            path = Path(path)
        except Exception:
            continue
        if path.exists() and path.is_dir():
            key = str(path.resolve()).lower()
            if key not in seen:
                seen.add(key)
                result.append(path)
    return result


def _prepend_paths(paths) -> None:
    existing = _existing_dirs(paths)
    if not existing:
        return

    current_parts = [p for p in os.environ.get("PATH", "").split(os.pathsep) if p]
    current_keys = {str(Path(p).resolve()).lower() for p in current_parts if Path(p).exists()}

    new_parts = []
    for path in existing:
        key = str(path.resolve()).lower()
        if key not in current_keys:
            new_parts.append(str(path))
            current_keys.add(key)

    os.environ["PATH"] = os.pathsep.join(new_parts + current_parts) if current_parts else os.pathsep.join(new_parts)


base_dir = Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
app_dir = Path(sys.executable).resolve().parent

pandoc_dir = base_dir / "vendor" / "pandoc"
pandoc_data_dir = base_dir / "vendor" / "pandoc-data"
tex_dir = base_dir / "vendor" / "tex"

tex_candidate_dirs = [
    tex_dir,
    tex_dir / "miktex" / "bin" / "x64",
    tex_dir / "miktex" / "bin",
    tex_dir / "bin" / "windows",
    tex_dir / "bin" / "win32",
    tex_dir / "bin",
]

_prepend_paths([pandoc_dir] + tex_candidate_dirs)

pandoc_exe = pandoc_dir / "pandoc.exe"
if pandoc_exe.exists():
    os.environ["PYPANDOC_PANDOC"] = str(pandoc_exe)

if pandoc_data_dir.exists():
    os.environ["PANDOC_DATA_DIR"] = str(pandoc_data_dir)

# Helpful for applications that want the real executable directory instead of the
# temporary extraction directory used by onefile mode.
os.environ.setdefault("APP_BUNDLE_DIR", str(app_dir))
