import os
import sys
from pathlib import Path

from PyInstaller.utils.hooks import collect_all, copy_metadata

ROOT = Path(SPECPATH).parent
APP = ROOT / "thermofisher_invoice_app_simplex"
NAME = "Simplex Invoices and Quotes"
IS_MAC = sys.platform == "darwin"
IS_WINDOWS = sys.platform == "win32"
TESSERACT = Path(os.environ["TESSERACT_PREFIX"])
TESSDATA = Path(os.environ["TESSDATA_DIR"])
VERSION = os.environ.get("APP_VERSION", "0.0.0")

modules = [p for p in APP.glob("*.py") if p.name not in ("launcher.py", "desktop.py")]
datas = [(str(p), "app") for p in modules]
datas += [(str(p), "app") for p in APP.iterdir() if p.suffix in (".css", ".docx", ".png")]
datas += [(str(APP / ".streamlit" / "config.toml"), "app/.streamlit")]
datas += [(str(TESSDATA / "eng.traineddata"), "tesseract/tessdata")]
for folder in ("configs", "tessconfigs"):
    datas += [(str(p), f"tesseract/tessdata/{folder}") for p in (TESSDATA / folder).glob("*") if p.is_file()]
binaries = []
hiddenimports = [p.stem for p in modules] + ["certifi", "webview"]

collected = collect_all("streamlit")
datas += collected[0]
binaries += collected[1]
hiddenimports += collected[2]
for distribution in ("pymupdf", "docxtpl", "python-docx", "pytesseract", "pywebview", "pandas", "pillow", "numpy"):
    datas += copy_metadata(distribution)

if IS_MAC:
    binaries += [(os.path.realpath(TESSERACT / "bin" / "tesseract"), "tesseract/bin")]
if IS_WINDOWS:
    tesseract_bin = next(p for p in (TESSERACT / "Library" / "bin", TESSERACT) if (p / "tesseract.exe").exists())
    datas += [(str(tesseract_bin / "tesseract.exe"), "tesseract")]
    datas += [(str(p), "tesseract") for p in tesseract_bin.glob("*.dll")]

a = Analysis(
    [str(APP / "desktop.py")],
    pathex=[str(APP)],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    excludes=["tkinter", "IPython", "matplotlib", "easyocr", "torch", "torchvision", "cv2", "scipy", "skimage"],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name=NAME,
    console=False,
    upx=False,
    argv_emulation=False,
    icon=str(ROOT / "packaging" / ("simplex.icns" if IS_MAC else "simplex.ico")),
)
coll = COLLECT(exe, a.binaries, a.datas, upx=False, name=NAME)
if IS_MAC:
    app = BUNDLE(
        coll,
        name=f"{NAME}.app",
        icon=str(ROOT / "packaging" / "simplex.icns"),
        bundle_identifier="com.simplexsciences.invoices",
        info_plist={
            "CFBundleDisplayName": NAME,
            "CFBundleShortVersionString": VERSION,
            "CFBundleVersion": VERSION,
            "LSMinimumSystemVersion": os.environ.get("MIN_MACOS", "14.0"),
            "NSHighResolutionCapable": True,
        },
    )
