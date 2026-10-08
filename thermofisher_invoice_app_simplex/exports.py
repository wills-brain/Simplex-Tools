from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import uuid
from datetime import datetime
from pathlib import Path

IS_WINDOWS = sys.platform == "win32"
IS_MAC = sys.platform == "darwin"
WORD_MAC = Path("/Applications/Microsoft Word.app")


def data_dir() -> Path:
    override = os.environ.get("SIMPLEX_APP_HOME")
    if override:
        path = Path(override)
    elif IS_WINDOWS:
        path = Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local") / "SimplexInvoiceApp"
    elif IS_MAC:
        path = Path.home() / "Library" / "Application Support" / "SimplexInvoiceApp"
    else:
        path = Path(os.environ.get("XDG_DATA_HOME") or Path.home() / ".local" / "share") / "SimplexInvoiceApp"
    path.mkdir(parents=True, exist_ok=True)
    return path


def document_name(kind: str, code: str) -> str:
    cleaned = re.sub(r'[\\/:*?"<>|\r\n\t]+', " ", str(code or "")).strip()
    return f"{kind.upper()} {cleaned}".strip()


def _soffice() -> str | None:
    candidates = [shutil.which("soffice"), shutil.which("libreoffice")]
    if IS_MAC:
        candidates.append("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    if IS_WINDOWS:
        for base in (os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)")):
            if base:
                candidates.append(str(Path(base) / "LibreOffice" / "program" / "soffice.exe"))
    return next((c for c in candidates if c and Path(c).is_file()), None)


def _word_windows() -> bool:
    if not IS_WINDOWS:
        return False
    import winreg

    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CLSID"):
            return True
    except OSError:
        return False


def pdf_converter() -> str | None:
    if IS_MAC and WORD_MAC.exists():
        return "word"
    if _word_windows():
        return "word"
    if _soffice():
        return "libreoffice"
    return None


def _run(command: list[str], timeout: int = 180) -> None:
    flags = subprocess.CREATE_NO_WINDOW if IS_WINDOWS else 0
    result = subprocess.run(command, capture_output=True, text=True, timeout=timeout, creationflags=flags)
    if result.returncode != 0:
        raise RuntimeError((result.stderr or result.stdout or "the converter stopped").strip()[-400:])


def _word_mac(source: Path, target: Path) -> None:
    source_text, target_text = (str(path).replace("\\", "\\\\").replace('"', '\\"') for path in (source, target))
    script = f'''
tell application "Microsoft Word"
    set opened to open file name "{source_text}" with read only
    save as opened file name "{target_text}" file format format PDF
    close opened saving no
end tell
'''
    _run(["/usr/bin/osascript", "-e", script])


def _word_win(source: Path, target: Path) -> None:
    source, target = (str(path).replace("'", "''") for path in (source, target))
    script = (
        "$ErrorActionPreference = 'Stop'; $word = New-Object -ComObject Word.Application; $word.Visible = $false; $word.DisplayAlerts = 0; "
        f"try {{ $doc = $word.Documents.Open('{source}', $false, $true); $doc.ExportAsFixedFormat('{target}', 17); $doc.Close(0) }} finally {{ $word.Quit() }}"
    )
    _run(["powershell", "-NoProfile", "-NonInteractive", "-Command", script])


def _work_folder() -> Path:
    if IS_MAC and WORD_MAC.exists():
        base = Path.home() / "Library" / "Containers" / "com.microsoft.Word" / "Data" / "tmp"
    else:
        base = Path(tempfile.gettempdir())
    folder = base / f"simplex-{uuid.uuid4().hex}"
    folder.mkdir(parents=True)
    return folder


def docx_to_pdf(docx_bytes: bytes) -> bytes:
    converter = pdf_converter()
    if converter is None:
        raise RuntimeError("PDF export needs Microsoft Word or LibreOffice on this computer.")
    folder = _work_folder()
    try:
        source, target = folder / "document.docx", folder / "document.pdf"
        source.write_bytes(docx_bytes)
        if converter == "word":
            (_word_mac if IS_MAC else _word_win)(source, target)
        else:
            _run([_soffice(), "--headless", "--convert-to", "pdf", "--outdir", str(folder), str(source)])
        for _ in range(20):
            if target.exists() and target.stat().st_size:
                break
            time.sleep(0.25)
        return target.read_bytes()
    finally:
        shutil.rmtree(folder, ignore_errors=True)


def history_dir() -> Path:
    folder = data_dir() / "history"
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def _index() -> Path:
    return history_dir() / "history.json"


def load_history() -> list[dict]:
    try:
        entries = json.loads(_index().read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return []
    return [e for e in entries if isinstance(e, dict) and (history_dir() / str(e.get("file", ""))).is_file()]


def record(kind: str, name: str, docx_bytes: bytes, customer: str = "", total: str = "") -> None:
    entries = load_history()
    if entries and entries[0].get("name") == name and (history_dir() / entries[0]["file"]).read_bytes() == docx_bytes:
        return
    stamp = datetime.now()
    file_name = f"{stamp:%Y%m%d-%H%M%S}-{uuid.uuid4().hex[:6]}.docx"
    (history_dir() / file_name).write_bytes(docx_bytes)
    entries.insert(0, {"kind": kind, "name": name, "file": file_name, "created": stamp.isoformat(timespec="seconds"), "customer": customer, "total": total})
    _index().write_text(json.dumps(entries, indent=2), encoding="utf-8")


def history_file(entry: dict) -> bytes:
    return (history_dir() / entry["file"]).read_bytes()
