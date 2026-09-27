from __future__ import annotations

import argparse
import hashlib
import json
import os
import platform
import shutil
import subprocess
import sys
import sysconfig
import time
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
DESKTOP = APP_DIR / "desktop.py"
REQUIREMENTS = APP_DIR / "requirements.txt"

SETUP_VERSION = 2
SUPPORTED_PYTHONS = ((3, 10), (3, 14))
CORE_MODULES = "streamlit, pymupdf, docx, docxtpl, pandas, PIL, pytesseract, webview"

IS_WINDOWS = sys.platform == "win32"
IS_MAC = sys.platform == "darwin"
IS_INTEL_MAC = IS_MAC and platform.machine() == "x86_64"

PYTHON_DOWNLOAD_URL = "https://www.python.org/downloads/windows/" if IS_WINDOWS else "https://www.python.org/downloads/macos/"


class SetupError(Exception):
    pass


def say(message: str = "") -> None:
    print(message, flush=True)


def heading(message: str) -> None:
    say()
    say(f"== {message}")


def state_dir() -> Path:
    override = os.environ.get("SIMPLEX_APP_HOME")
    if override:
        return Path(override)
    if IS_WINDOWS:
        if "PythonSoftwareFoundation.Python." in os.path.realpath(sys.executable):
            return Path.home() / "SimplexInvoiceApp"
        return Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local") / "SimplexInvoiceApp"
    if IS_MAC:
        return Path.home() / "Library" / "Application Support" / "SimplexInvoiceApp"
    return Path(os.environ.get("XDG_DATA_HOME") or Path.home() / ".local" / "share") / "simplex-invoice-app"


def venv_python(venv: Path) -> Path:
    return venv / "Scripts" / "python.exe" if IS_WINDOWS else venv / "bin" / "python"


def check_interpreter() -> None:
    low, high = SUPPORTED_PYTHONS
    version = sys.version_info[:2]
    install_hint = f"Install Python 3.13 from {PYTHON_DOWNLOAD_URL}, then start the app again."
    if platform.python_implementation() != "CPython":
        raise SetupError(f"This app needs the standard Python (CPython), not {platform.python_implementation()}. {install_hint}")
    if not low <= version <= high:
        raise SetupError(f"This app needs Python {low[0]}.{low[1]} to {high[0]}.{high[1]}; this is Python {version[0]}.{version[1]}. {install_hint}")
    if sys.maxsize <= 2**32:
        raise SetupError(f"This is a 32 bit Python; the app needs 64 bit Python. {install_hint}")
    if sysconfig.get_config_var("Py_GIL_DISABLED"):
        raise SetupError(f"This Python build (python{version[0]}.{version[1]}t) is not supported by the app's libraries. {install_hint}")
    if IS_WINDOWS and sysconfig.get_platform() != "win-amd64":
        raise SetupError(
            f"This is a {sysconfig.get_platform()} Python. Several of the app's libraries only exist for x64 Windows. "
            f"Install the x64 Python 3.13 from {PYTHON_DOWNLOAD_URL} ('Windows installer (64 bit)'), then start the app again."
        )
    if IS_INTEL_MAC and _rosetta():
        say(f"Note: this Python runs under Rosetta (Intel emulation) on an Apple Silicon Mac. It works, but a native Python 3.13 from {PYTHON_DOWNLOAD_URL} is faster.")


def _rosetta() -> bool:
    try:
        out = subprocess.run(["sysctl", "-in", "sysctl.proc_translated"], capture_output=True, text=True, timeout=5)
        return out.stdout.strip() == "1"
    except (OSError, subprocess.SubprocessError):
        return False


def find_tesseract() -> Path | None:
    configured = os.environ.get("TESSERACT_CMD")
    if configured and Path(configured).is_file():
        return Path(configured)
    on_path = shutil.which("tesseract")
    if on_path:
        return Path(on_path)
    candidates: list[Path] = []
    if IS_WINDOWS:
        for variable, tail in (
            ("ProgramFiles", r"Tesseract-OCR\tesseract.exe"),
            ("LOCALAPPDATA", r"Programs\Tesseract-OCR\tesseract.exe"),
            ("ProgramFiles(x86)", r"Tesseract-OCR\tesseract.exe"),
        ):
            base = os.environ.get(variable)
            if base:
                candidates.append(Path(base) / tail)
        try:
            import winreg

            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Tesseract-OCR", 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY) as key:
                candidates.append(Path(winreg.QueryValueEx(key, "InstallDir")[0]) / "tesseract.exe")
        except OSError:
            pass
    else:
        candidates += [Path("/opt/homebrew/bin/tesseract"), Path("/usr/local/bin/tesseract"), Path("/opt/local/bin/tesseract")]
    return next((c for c in candidates if c.is_file()), None)


def tesseract_help() -> str:
    if IS_WINDOWS:
        return (
            "Scanned (image only) PDFs need the free Tesseract OCR program. To install it, open Command Prompt and run:\n"
            "    winget install --id tesseract-ocr.tesseract -e\n"
            "or download the installer from https://github.com/tesseract-ocr/tesseract/releases\n"
            "PDFs that contain selectable text work without it."
        )
    if IS_MAC:
        brew = "brew install tesseract" if shutil.which("brew") or Path("/opt/homebrew/bin/brew").exists() or Path("/usr/local/bin/brew").exists() else (
            "install Homebrew from https://brew.sh, then run: brew install tesseract"
        )
        return (
            f"Scanned (image only) PDFs need the free Tesseract OCR program. To install it, open Terminal and {brew}.\n"
            "PDFs that contain selectable text work without it."
        )
    return "Scanned (image only) PDFs need Tesseract OCR: install the 'tesseract-ocr' package with your package manager."


class Instance:
    def __init__(self, state: Path) -> None:
        self.lock_path = state / "instance.lock"
        self._lock = None

    def claim(self) -> bool:
        handle = open(self.lock_path, "a+")
        try:
            if IS_WINDOWS:
                import msvcrt

                handle.seek(0)
                msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl

                fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError:
            handle.close()
            return False
        self._lock = handle
        return True

    def release(self) -> None:
        if self._lock is None:
            return
        try:
            if IS_WINDOWS:
                import msvcrt

                self._lock.seek(0)
                msvcrt.locking(self._lock.fileno(), msvcrt.LK_UNLCK, 1)
        except OSError:
            pass
        self._lock.close()
        self._lock = None


def _read_json(path: Path) -> dict:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def _hash_files(*paths: Path) -> str:
    digest = hashlib.sha256(str(SETUP_VERSION).encode())
    for path in paths:
        if path.exists():
            digest.update(path.name.encode())
            digest.update(path.read_bytes())
    return digest.hexdigest()


def _run(command: list, **kwargs) -> subprocess.CompletedProcess:
    return subprocess.run([str(part) for part in command], **kwargs)


def _pip(python: Path, args: list, log: Path) -> bool:
    return _run([python, "-m", "pip", "--disable-pip-version-check", "--log", log, *args]).returncode == 0


def _venv_works(python: Path) -> bool:
    if not python.exists():
        return False
    try:
        return _run([python, "-c", "import sys, pip"], capture_output=True, timeout=120).returncode == 0
    except (OSError, subprocess.SubprocessError):
        return False


def _windows_dll_hint(output: str) -> str:
    if "4551" in output or "Application Control policy" in output:
        return (
            "\nWindows Smart App Control blocked part of the app. Turn it off in Windows Security > "
            "App & browser control > Smart App Control settings, then start the app again."
        )
    if "WinError 1114" in output or "WinError 126" in output:
        return (
            "\nA Windows system component is missing. Install the Microsoft Visual C++ Redistributable (x64) from "
            "https://aka.ms/vs/17/release/vc_redist.x64.exe, then start the app again."
        )
    return ""


def ensure_environment(state: Path, reset: bool) -> Path:
    tag = f"py{sys.version_info[0]}{sys.version_info[1]}-{platform.machine().lower() or 'unknown'}"
    venv = state / f"venv-{tag}"
    python = venv_python(venv)
    stamp_path = venv / "simplex-setup.json"
    wanted = _hash_files(REQUIREMENTS)

    if reset and venv.exists():
        heading("Removing the old app environment")
        shutil.rmtree(venv, ignore_errors=True)

    if _read_json(stamp_path).get("requirements") == wanted and _venv_works(python):
        return python

    heading("Setting up the app (first run, or the app was updated)")
    say("This downloads about 300 MB and can take a few minutes. It only happens once.")
    say(f"Setup files go in: {state}")
    log = state / "setup.log"

    if not _venv_works(python):
        shutil.rmtree(venv, ignore_errors=True)
        say("Creating the app's Python environment...")
        created = _run([sys.executable, "-m", "venv", venv], capture_output=True, text=True)
        if created.returncode != 0 or not python.exists():
            detail = (created.stderr or created.stdout or "").strip().splitlines()[-1:] or [""]
            hint = " On Linux, install your distribution's python3-venv package." if not (IS_WINDOWS or IS_MAC) else ""
            raise SetupError(f"Could not create a Python environment ({detail[0]}).{hint}")

    say("Updating pip...")
    _pip(python, ["install", "--quiet", "--upgrade", "pip"], log)
    say("Installing the app's components...")
    if not _pip(python, ["install", "--prefer-binary", "-r", REQUIREMENTS], log):
        raise SetupError(f"Installing the app's components failed (details above and in {log}). Check the internet connection and start the app again.")
    check = _run([python, "-c", f"import {CORE_MODULES}"], capture_output=True, text=True)
    if check.returncode != 0:
        output = (check.stderr or "") + (check.stdout or "")
        raise SetupError(f"The app's components were installed but do not load:\n{output.strip()[-1500:]}{_windows_dll_hint(output)}")

    stamp_path.write_text(json.dumps({"requirements": wanted, "python": sys.executable, "installed": time.strftime("%Y-%m-%d %H:%M")}, indent=2), encoding="utf-8")
    say("Setup finished.")
    return python


def app_environment(tesseract: Path | None) -> dict[str, str]:
    env = dict(os.environ)
    if tesseract is not None:
        env["PATH"] = str(tesseract.parent) + os.pathsep + env.get("PATH", "")
    env["PYTHONUNBUFFERED"] = "1"
    return env


def run_app(python: Path) -> int:
    tesseract = find_tesseract()
    if tesseract is None:
        heading("Tesseract OCR not found")
        say(tesseract_help())
    heading("Opening the app window")
    say("Keep this window open while you use the app. Closing the app window stops it.")
    process = subprocess.Popen([str(python), str(DESKTOP)], cwd=APP_DIR, env=app_environment(tesseract))
    try:
        code = process.wait()
    except KeyboardInterrupt:
        process.terminate()
        try:
            process.wait(timeout=10)
        except subprocess.TimeoutExpired:
            process.kill()
        return 0
    if code != 0:
        raise SetupError(
            f"The app stopped unexpectedly (exit code {code}); the messages above say why. "
            "If this keeps happening, start the launcher with --reset to rebuild the app's environment."
        )
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Set up and start the Simplex invoice and quote app.")
    parser.add_argument("--reset", action="store_true", help="rebuild the app's Python environment first")
    parser.add_argument("--setup-only", action="store_true", help="set up, but do not start the app")
    args = parser.parse_args(argv)

    say("Simplex Invoices and Quotes")
    try:
        check_interpreter()
        state = state_dir()
        state.mkdir(parents=True, exist_ok=True)
        instance = Instance(state)
        if not instance.claim():
            say("The app is already open in another window.")
            return 0
        try:
            python = ensure_environment(state, args.reset)
            if args.setup_only:
                say(f"Ready. Python environment: {python}")
                return 0
            return run_app(python)
        finally:
            instance.release()
    except SetupError as problem:
        say()
        say("PROBLEM: " + str(problem))
        return 1
    except KeyboardInterrupt:
        return 0


if __name__ == "__main__":
    sys.exit(main())
