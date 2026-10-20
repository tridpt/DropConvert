"""Debug missing procedures in bundled Windows DLLs."""

from pathlib import Path

import pefile


root = Path(__file__).resolve().parent.parent / "dist" / "ThaVaDoi" / "_internal"
search = [root / "PySide6", root / "shiboken6", root]
for source in (root / "PySide6" / "QtCore.pyd", root / "PySide6" / "Qt6Core.dll"):
    print(source.name)
    binary = pefile.PE(str(source))
    for entry in binary.DIRECTORY_ENTRY_IMPORT:
        name = entry.dll.decode()
        candidates = [folder / name for folder in search if (folder / name).exists()]
        if not candidates:
            continue
        library = pefile.PE(str(candidates[0]))
        exports = {item.name for item in library.DIRECTORY_ENTRY_EXPORT.symbols if item.name}
        missing = [item.name.decode(errors="replace") for item in entry.imports if item.name and item.name not in exports]
        if missing:
            print(" ", name, "missing", len(missing), missing[:8])
