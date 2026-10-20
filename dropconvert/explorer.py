"""Optional, per-user File Explorer context menu integration on Windows."""

from __future__ import annotations

import sys
import os
import shutil
import tempfile
import uuid
from pathlib import Path

from .core import SUPPORTED_SUFFIXES


VERB_NAME = "ThaVaDoi"
ROOT = r"Software\Classes\SystemFileAssociations"
MANAGED_VALUE = "DropConvertManaged"
SENDTO_NAME = "Thả & Đổi.lnk"


def available() -> bool:
    return sys.platform == "win32" and bool(getattr(sys, "frozen", False))


def _verb_path(suffix: str) -> str:
    return f"{ROOT}\\{suffix}\\shell\\{VERB_NAME}"


def _command(executable: Path) -> str:
    return f'"{executable}" --open "%1"'


def _sendto_path() -> Path:
    roaming = Path(os.environ.get("APPDATA", str(Path.home() / "AppData" / "Roaming")))
    return roaming / "Microsoft" / "Windows" / "SendTo" / SENDTO_NAME


def _shortcut_owned(path: Path, executable: Path | None = None) -> bool:
    if not path.is_file():
        return False
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    try:
        # WScript.Shell can mis-handle accented shortcut paths on some Windows locales.
        with tempfile.TemporaryDirectory() as temporary:
            inspection_path = Path(temporary) / "ThaVaDoi.lnk"
            shutil.copy2(path, inspection_path)
            shortcut = win32com.client.Dispatch("WScript.Shell").CreateShortcut(str(inspection_path))
            target = Path(shortcut.TargetPath)
            owned = "DropConvertManaged" in shortcut.Description and shortcut.Arguments.strip() == "--open"
            del shortcut
        return owned and (target.resolve() == executable.resolve() if executable else True)
    finally:
        pythoncom.CoUninitialize()


def is_registered(executable: Path | None = None) -> bool:
    if sys.platform != "win32":
        return False
    import winreg

    target = _command((executable or Path(sys.executable)).resolve())
    for suffix in SUPPORTED_SUFFIXES:
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _verb_path(suffix)) as key:
                if winreg.QueryValueEx(key, MANAGED_VALUE)[0] != 1:
                    return False
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _verb_path(suffix) + r"\command") as key:
                if winreg.QueryValueEx(key, "")[0] != target:
                    return False
        except OSError:
            return False
    return _shortcut_owned(_sendto_path(), (executable or Path(sys.executable)).resolve())


def register(executable: Path | None = None) -> None:
    if sys.platform != "win32":
        raise RuntimeError("Tích hợp File Explorer chỉ hỗ trợ Windows")
    import winreg

    exe = (executable or Path(sys.executable)).resolve()
    if not exe.is_file() or exe.suffix.lower() != ".exe":
        raise ValueError("Không tìm thấy tệp ứng dụng .exe")
    shortcut_path = _sendto_path()
    if shortcut_path.exists() and not _shortcut_owned(shortcut_path):
        raise RuntimeError("Lối tắt Gửi tới → Thả & Đổi đã được ứng dụng khác sử dụng")
    for suffix in SUPPORTED_SUFFIXES:
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _verb_path(suffix)) as key:
                try:
                    managed = winreg.QueryValueEx(key, MANAGED_VALUE)[0] == 1
                except FileNotFoundError:
                    managed = False
                if not managed:
                    raise RuntimeError(f"Lệnh chuột phải cho {suffix} đã được ứng dụng khác sử dụng")
        except FileNotFoundError:
            pass
    for suffix in sorted(SUPPORTED_SUFFIXES):
        with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, _verb_path(suffix), 0, winreg.KEY_WRITE) as key:
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "Mở bằng Thả & Đổi")
            winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, str(exe))
            winreg.SetValueEx(key, "MultiSelectModel", 0, winreg.REG_SZ, "Single")
            winreg.SetValueEx(key, MANAGED_VALUE, 0, winreg.REG_DWORD, 1)
        with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, _verb_path(suffix) + r"\command", 0, winreg.KEY_WRITE) as key:
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, _command(exe))
    import pythoncom
    import win32com.client

    shortcut_path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path = shortcut_path.parent / ("ThaVaDoi-" + uuid.uuid4().hex + ".lnk")
    pythoncom.CoInitialize()
    try:
        shortcut = win32com.client.Dispatch("WScript.Shell").CreateShortcut(str(temporary_path))
        shortcut.TargetPath = str(exe)
        shortcut.Arguments = "--open"
        shortcut.IconLocation = str(exe)
        shortcut.Description = "Gửi nhiều tệp tới Thả & Đổi (DropConvertManaged)"
        shortcut.Save()
        del shortcut
        temporary_path.replace(shortcut_path)
    finally:
        pythoncom.CoUninitialize()
        temporary_path.unlink(missing_ok=True)


def unregister() -> None:
    if sys.platform != "win32":
        return
    import winreg

    for suffix in sorted(SUPPORTED_SUFFIXES):
        path = _verb_path(suffix)
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as key:
                if winreg.QueryValueEx(key, MANAGED_VALUE)[0] != 1:
                    continue
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, path + r"\command")
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, path)
        except FileNotFoundError:
            continue
    shortcut_path = _sendto_path()
    if _shortcut_owned(shortcut_path):
        shortcut_path.unlink()

