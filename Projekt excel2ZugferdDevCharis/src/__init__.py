import os
import sys
from pathlib import Path
import src.oberflaeche_base
import src.oberflaeche_excel2zugferd
import src.oberflaeche_ini
import src.oberflaeche_steuerung
import src.oberflaeche_excelpositions
import src.oberflaeche_excelsteuerung  # noqa F404


def _normalize(arr_in: list) -> list:
    """remove empty elements of array"""
    return list(filter(None, arr_in))


def _setNoneIfEmpty(str_in: str) -> str | None:
    # print("_setNoneIfEmpty:", str_in)
    if str_in is None:
        return None
    trimmed = str_in.strip()
    trimmed = " ".join(trimmed.split())
    return trimmed if trimmed != "" else None


def _get_exe_dir() -> Path:
    """Gibt das Verzeichnis der laufenden EXE zurück (PyInstaller), im Dev-Modus das Projektverzeichnis."""
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).parent.parent


def logo_fn() -> str:
    logo_neben_exe = _get_exe_dir() / "logo.jpg"
    if logo_neben_exe.exists():
        return str(logo_neben_exe)
    return os.path.join(
        os.getenv("APPDATA"), "excel2zugferd", "logo.jpg"  # type: ignore
    )


def sig_fn() -> str:
    """Gibt den Pfad zur Unterschrift-Datei zurück.
    Sucht zuerst neben der EXE, dann in %APPDATA%, dann im assets-Ordner (Dev-Modus).
    """
    sig_neben_exe = _get_exe_dir() / "signatur.png"
    if sig_neben_exe.exists():
        return str(sig_neben_exe)
    appdata_sig = Path(os.getenv("APPDATA", "")) / "excel2zugferd" / "signatur.png"
    if appdata_sig.exists():
        return str(appdata_sig)
    # Dev-Modus: assets/signatur.png relativ zum Projektverzeichnis
    dev_sig = _get_exe_dir() / "assets" / "signatur.png"
    return str(dev_sig)
