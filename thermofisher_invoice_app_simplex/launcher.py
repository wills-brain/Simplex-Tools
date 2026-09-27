"""Set up and start the Simplex invoice & quote app.

Run by the double click launchers ("Start Invoice App (Mac).command" and
"Start Invoice App (Windows).bat") with a suitable system Python:

    python launcher.py               start the app (setting it up first if needed)
    python launcher.py --reset       rebuild the app's Python environment, then start
    python launcher.py --setup-only  set up without starting
    python launcher.py --no-browser  start without opening a browser tab

The first run creates a private Python environment in the user's app data
folder (outside the download, so paths stay short, cloud synced folders are
avoided, and a newly downloaded copy of the app reuses it) and installs the
requirements. Later runs start in a few seconds. The app is served only to
this computer (127.0.0.1).
"""

from __future__ import annotations

import argparse
import hashlib
import http.client
import json
import os
import platform
import shutil
import socket
import subprocess
import sys
import sysconfig
import threading
import time
import urllib.request
import webbrowser
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
APP_SCRIPT = APP_DIR / "app.py"
REQUIREMENTS = APP_DIR / "requirements.txt"
OCR_REQUIREMENTS = APP_DIR / "requirements-ocr.txt"
INTEL_MAC_CONSTRAINTS = APP_DIR / "constraints-intel-mac.txt"

SETUP_VERSION = 1  # bump to make every install rerun setup
SUPPORTED_PYTHONS = ((3, 10), (3, 14))
PORTS = range(8501, 8531)
CORE_MODULES = "streamlit, pymupdf, docx, docxtpl, pandas, PIL, pytesseract"

IS_WINDOWS = sys.platform == "win32"
IS_MAC = sys.platform == "darwin"
IS_INTEL_MAC = IS_MAC and platform.machine() == "x86_64"

PYTHON_DOWNLOAD_URL = "https://www.python.org/downloads/windows/" if IS_WINDOWS else "https://www.python.org/downloads/macos/"


class SetupError(Exception):
    """A problem the user has to fix; the message says how."""


def say(message: str = "") -> None:
    print(message, flush=True)


def heading(message: str) -> None:
    say()
    say(f"== {message}")


# --------------------------------------------------------------------------
# Where things live
# --------------------------------------------------------------------------

def state_dir() -> Path:
    override = os.environ.get("SIMPLEX_APP_HOME")
    if override:
        return Path(override)
    if IS_WINDOWS:
        # Microsoft Store builds of Python silently redirect writes under
        # AppData into their own package folder, so keep their venv elsewhere.
        if "PythonSoftwareFoundation.Python." in os.path.realpath(sys.executable):
            return Path.home() / "SimplexInvoiceApp"
        return Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local") / "SimplexInvoiceApp"
    if IS_MAC:
        return Path.home() / "Library" / "Application Support" / "SimplexInvoiceApp"
    return Path(os.environ.get("XDG_DATA_HOME") or Path.home() / ".local" / "share") / "simplex-invoice-app"


def venv_python(venv: Path) -> Path:
    return venv / "Scripts" / "python.exe" if IS_WINDOWS else venv / "bin" / "python"


# --------------------------------------------------------------------------
# Checks
# --------------------------------------------------------------------------

def check_interpreter() -> None:
    low, high = SUPPORTED_PYTHONS
    version = sys.version_info[:2]
    install_hint = (
        f"Install Python 3.13 from {PYTHON_DOWNLOAD_URL}"
        + (" (on an Intel Mac, 3.12.10 works best)" if IS_INTEL_MAC else "")
        + ", then start the app again."
    )
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
        say(
            "Note: this Python runs under Rosetta (Intel emulation) on an Apple Silicon Mac. It works, but a native "
            f"Python 3.13 from {PYTHON_DOWNLOAD_URL} is faster and supports every optional component."
        )


def _rosetta() -> bool:
    try:
        out = subprocess.run(["sysctl", "-in", "sysctl.proc_translated"], capture_output=True, text=True, timeout=5)
        return out.stdout.strip() == "1"
    except (OSError, subprocess.SubprocessError):
        return False


def find_tesseract() -> Path | None:
    """Tesseract is a separate program (not a Python package) used to read scanned PDFs."""
    configured = os.environ.get("TESSERACT_CMD")
    if configured and Path(configured).is_file():
        return Path(configured)
    on_path = shutil.which("tesseract")
    if on_path:
        return Path(on_path)
    candidates: list[Path] = []
    if IS_WINDOWS:
        # The Windows installer does not add itself to PATH.
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
        # A Finder-launched Terminal may not have Homebrew or MacPorts on PATH.
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
        slow = " (on Intel Macs this builds from source and can take a while)" if IS_INTEL_MAC else ""
        return (
            f"Scanned (image only) PDFs need the free Tesseract OCR program. To install it, open Terminal and {brew}{slow}.\n"
            "PDFs that contain selectable text work without it."
        )
    return "Scanned (image only) PDFs need Tesseract OCR: install the 'tesseract-ocr' package with your package manager."


# --------------------------------------------------------------------------
# One running copy at a time
# --------------------------------------------------------------------------

def _read_json(path: Path) -> dict:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


class Instance:
    """One app per computer user.

    The launcher holds a lock on instance.lock while it runs; the operating
    system releases it when the process ends, including when the window is
    closed. instance.json holds the app's port so a second launcher can open it.
    """

    def __init__(self, state: Path) -> None:
        self.lock_path = state / "instance.lock"
        self.path = state / "instance.json"
        self._lock = None

    def _try_lock(self):
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
            return None
        return handle

    def claim(self) -> bool:
        """Become the running launcher; False if another one is running."""
        self._lock = self._try_lock()
        if self._lock is None:
            return False
        self._write(None)
        return True

    def held_elsewhere(self) -> bool:
        handle = self._try_lock()
        if handle is None:
            return True
        self._unlock(handle)
        return False

    def port(self) -> int | None:
        port = _read_json(self.path).get("port")
        return int(port) if port else None

    def set_port(self, port: int) -> None:
        self._write(port)

    def _write(self, port: int | None) -> None:
        temporary = self.path.with_suffix(".tmp")
        temporary.write_text(json.dumps({"pid": os.getpid(), "port": port}), encoding="utf-8")
        os.replace(temporary, self.path)

    def _unlock(self, handle) -> None:
        try:
            if IS_WINDOWS:
                import msvcrt

                handle.seek(0)
                msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
        except OSError:
            pass
        handle.close()

    def release(self) -> None:
        if self._lock is not None:
            self.path.unlink(missing_ok=True)
            self._unlock(self._lock)
            self._lock = None


# Loopback requests must not go through a proxy configured for the internet.
_LOCAL = urllib.request.build_opener(urllib.request.ProxyHandler({}))


def healthy(port: int, timeout: float = 1.0) -> bool:
    try:
        with _LOCAL.open(f"http://127.0.0.1:{port}/_stcore/health", timeout=timeout) as response:
            return response.status == 200 and response.read().strip() == b"ok"
    except (OSError, ValueError, http.client.HTTPException):
        return False


def free_port() -> int:
    for port in PORTS:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            if not IS_WINDOWS:
                # Like the app server itself, so a port the app just released
                # (still in TIME_WAIT) is reused and the address stays the same.
                sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                sock.bind(("127.0.0.1", port))
            except OSError:
                continue
            return port
    raise SetupError(f"Ports {PORTS.start} to {PORTS.stop - 1} are all in use. Restart the computer and try again.")


def open_browser(url: str, enabled: bool) -> None:
    if enabled:
        webbrowser.open(url)


def wait_for_other_window(instance: Instance, open_tab: bool) -> int:
    say("The app is already starting or running in another window.")
    deadline = time.monotonic() + 30 * 60
    while time.monotonic() < deadline:
        port = instance.port()
        if port and healthy(port):
            url = f"http://127.0.0.1:{port}"
            say(f"Opening {url}. This window can be closed.")
            open_browser(url, open_tab)
            return 0
        if not instance.held_elsewhere():
            say("The other window has closed. Double click the launcher again to start the app.")
            return 1
        time.sleep(2)
    say("The other window is still busy. Check it for progress or errors.")
    return 1


# --------------------------------------------------------------------------
# Setup
# --------------------------------------------------------------------------

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
    command = [python, "-m", "pip", "--disable-pip-version-check", "--log", log, *args]
    return _run(command).returncode == 0


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


def ensure_environment(state: Path, reset: bool) -> tuple[Path, dict]:
    """Create or update the app's Python environment. Returns its python and setup record."""
    tag = f"py{sys.version_info[0]}{sys.version_info[1]}-{platform.machine().lower() or 'unknown'}"
    venv = state / f"venv-{tag}"
    python = venv_python(venv)
    stamp_path = venv / "simplex-setup.json"
    wanted = _hash_files(REQUIREMENTS, OCR_REQUIREMENTS, INTEL_MAC_CONSTRAINTS)

    if reset and venv.exists():
        heading("Removing the old app environment")
        shutil.rmtree(venv, ignore_errors=True)

    stamp = _read_json(stamp_path)
    if stamp.get("requirements") == wanted and _venv_works(python):
        return python, stamp

    heading("Setting up the app (first run, or the app was updated)")
    say("This downloads about 1 GB and can take 5 to 15 minutes. It only happens once.")
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

    # NumPy 1.x for the Intel Mac torch build; it has no builds for Python 3.13+.
    constraints = ["-c", INTEL_MAC_CONSTRAINTS] if IS_INTEL_MAC and sys.version_info[:2] <= (3, 12) else []
    say("Installing the app's components...")
    if not _pip(python, ["install", "--prefer-binary", "-r", REQUIREMENTS, *constraints], log):
        raise SetupError(
            f"Installing the app's components failed (details above and in {log}). "
            "Check the internet connection and start the app again."
        )
    check = _run([python, "-c", f"import {CORE_MODULES}"], capture_output=True, text=True)
    if check.returncode != 0:
        output = (check.stderr or "") + (check.stdout or "")
        raise SetupError(f"The app's components were installed but do not load:\n{output.strip()[-1500:]}{_windows_dll_hint(output)}")

    ocr = install_ocr_helper(python, constraints, log)

    stamp = {"requirements": wanted, "python": sys.executable, "ocr": ocr, "installed": time.strftime("%Y-%m-%d %H:%M")}
    stamp_path.write_text(json.dumps(stamp, indent=2), encoding="utf-8")
    say("Setup finished.")
    return python, stamp


def install_ocr_helper(python: Path, constraints: list, log: Path) -> str:
    """EasyOCR helps read the DR order number on scanned POs. Optional: the app works without it."""
    if IS_INTEL_MAC and sys.version_info[:2] >= (3, 13):
        say("Skipping the optional EasyOCR helper: it needs Python 3.12 or older on Intel Macs.")
        return "skipped: Intel Mac with Python 3.13+"
    say("Installing the optional EasyOCR helper (large download)...")
    if not _pip(python, ["install", "--only-binary=:all:", "-r", OCR_REQUIREMENTS, *constraints], log):
        say("The optional EasyOCR helper could not be installed; the app works without it.")
        return "failed: install"
    say("Downloading the EasyOCR reading models (one time)...")
    try:
        warm = _run(
            [python, "-c", "import socket; socket.setdefaulttimeout(60); import easyocr; easyocr.Reader(['en'], gpu=False, verbose=False)"],
            capture_output=True, text=True, env=app_environment(python, None), timeout=30 * 60,
        )
    except subprocess.TimeoutExpired:
        say("Downloading the EasyOCR reading models timed out; the app works without them.")
        return "failed: model download timed out"
    if warm.returncode != 0:
        output = (warm.stderr or "") + (warm.stdout or "")
        say("The optional EasyOCR helper is installed but did not start; the app works without it." + _windows_dll_hint(output))
        return "failed: " + (output.strip().splitlines() or ["unknown error"])[-1][:200]
    return "installed"


def app_environment(python: Path, tesseract: Path | None) -> dict[str, str]:
    env = dict(os.environ)
    if tesseract is not None:
        env["PATH"] = str(tesseract.parent) + os.pathsep + env.get("PATH", "")
    if IS_MAC and "SSL_CERT_FILE" not in env:
        # python.org's macOS Python ships without root certificates until its
        # "Install Certificates" step is run; use certifi's for HTTPS instead.
        where = _run([python, "-c", "import certifi; print(certifi.where())"], capture_output=True, text=True)
        if where.returncode == 0 and where.stdout.strip():
            env["SSL_CERT_FILE"] = where.stdout.strip()
    env["PYTHONUNBUFFERED"] = "1"
    return env


# --------------------------------------------------------------------------
# Run
# --------------------------------------------------------------------------

def run_app(python: Path, instance: Instance, open_tab: bool) -> int:
    tesseract = find_tesseract()
    if tesseract is None:
        heading("Tesseract OCR not found")
        say(tesseract_help())

    port = free_port()
    instance.set_port(port)
    url = f"http://127.0.0.1:{port}"
    command = [
        python, "-m", "streamlit", "run", APP_SCRIPT,
        "--server.address", "127.0.0.1",
        "--server.port", str(port),
        "--server.headless", "true",  # no first-run email prompt; this launcher opens the browser
        "--browser.gatherUsageStats", "false",
        "--server.fileWatcherType", "none",
    ]  # theme and toolbar settings are in .streamlit/config.toml
    heading("Starting the app")
    process = subprocess.Popen([str(part) for part in command], cwd=APP_DIR, env=app_environment(python, tesseract))

    def announce_when_ready() -> None:
        deadline = time.monotonic() + 180
        while time.monotonic() < deadline and process.poll() is None:
            if healthy(port):
                say()
                say("=" * 64)
                say(f"  The app is running at {url}")
                say("  Keep this window open while you use the app.")
                say("  To stop the app, close this window.")
                say("=" * 64)
                open_browser(url, open_tab)
                return
            time.sleep(0.5)

    threading.Thread(target=announce_when_ready, daemon=True).start()
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
    parser = argparse.ArgumentParser(description="Set up and start the Simplex invoice & quote app.")
    parser.add_argument("--reset", action="store_true", help="rebuild the app's Python environment first")
    parser.add_argument("--setup-only", action="store_true", help="set up, but do not start the app")
    parser.add_argument("--no-browser", action="store_true", help="do not open a browser tab")
    args = parser.parse_args(argv)

    say("Simplex Invoices & Quotes")
    try:
        check_interpreter()
        state = state_dir()
        state.mkdir(parents=True, exist_ok=True)
        instance = Instance(state)
        if not instance.claim():
            return wait_for_other_window(instance, not args.no_browser)
        try:
            python, stamp = ensure_environment(state, args.reset)
            ocr = stamp.get("ocr", "installed")
            if ocr.startswith("failed"):
                say(f"(The optional EasyOCR helper is not active: {ocr}. Start with --reset to try again.)")
            elif ocr.startswith("skipped"):
                say("(The optional EasyOCR helper needs Python 3.12 or older on Intel Macs. The app works without it.)")
            if args.setup_only:
                say(f"Ready. Python environment: {python}")
                return 0
            return run_app(python, instance, not args.no_browser)
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
