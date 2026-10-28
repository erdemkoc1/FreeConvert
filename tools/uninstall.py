"""FreeConvert kaldırıcı (uninstall.exe).

Başlat menüsü kısayolunu, ayar/geçmiş klasörünü ve FreeConvert.exe ile bu
dosyanın bulunduğu klasörü siler. Klasörde başka dosya varsa klasöre
dokunmaz; yalnızca FreeConvert dosyalarını kaldırır.
"""
from __future__ import annotations

import os
import subprocess
import sys
import tkinter as tk
from tkinter import messagebox

SHORTCUT_NAME = "FreeConvert.lnk"
APP_EXE = "FreeConvert.exe"
UNINSTALL_EXE = "uninstall.exe"
NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)
DETACHED = getattr(subprocess, "DETACHED_PROCESS", 0)


def _app_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _start_menu_shortcut() -> str:
    appdata = os.environ.get("APPDATA") or os.path.expanduser("~")
    return os.path.join(appdata, "Microsoft", "Windows", "Start Menu",
                        "Programs", SHORTCUT_NAME)


def _data_dir() -> str:
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    return os.path.join(base, "FreeConvert")


def _remove_file(path: str) -> None:
    try:
        if os.path.isfile(path):
            os.remove(path)
    except OSError:
        pass


def _schedule_app_folder_cleanup(app_dir: str) -> None:
    """uninstall.exe kendini çalışırken silemez; kısa gecikmeyle siletir.

    Yalnızca FreeConvert.exe ve uninstall.exe silinir, klasör boşsa kaldırılır.
    rmdir /s kullanılmaz; böylece yanlış klasörde çalışsa bile başka dosyalar
    silinmez.
    """
    app = os.path.join(app_dir, APP_EXE)
    me = os.path.join(app_dir, UNINSTALL_EXE)
    command = (
        f'ping -n 3 127.0.0.1 >nul & del /q "{app}" "{me}" 2>nul '
        f'& rmdir "{app_dir}" 2>nul'
    )
    subprocess.Popen(["cmd", "/c", command],
                     creationflags=NO_WINDOW | DETACHED, close_fds=True)


def main() -> None:
    root = tk.Tk()
    root.withdraw()
    app_dir = _app_dir()
    answer = messagebox.askyesno(
        "FreeConvert - Kaldır",
        "FreeConvert kaldırılsın mı?\n\n"
        "Başlat menüsü kısayolu, ayarlar ve dönüşüm geçmişi silinecek.\n"
        f"{app_dir} klasöründeki FreeConvert.exe ve uninstall.exe de silinecek.",
        icon="warning",
    )
    if not answer:
        root.destroy()
        return

    _remove_file(_start_menu_shortcut())
    data_dir = _data_dir()
    try:
        for name in ("settings.json", "history.json"):
            _remove_file(os.path.join(data_dir, name))
        os.rmdir(data_dir)
    except OSError:
        pass

    _schedule_app_folder_cleanup(app_dir)
    messagebox.showinfo("FreeConvert - Kaldır",
                        "FreeConvert kaldırıldı.")
    root.destroy()


if __name__ == "__main__":
    main()
