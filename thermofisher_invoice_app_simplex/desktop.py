from __future__ import annotations

import json
import os
import platform
import socket
import subprocess
import sys
import threading
import time
import urllib.request
from pathlib import Path

APP_NAME = "Simplex Invoices and Quotes"
FROZEN = getattr(sys, "frozen", False)
BUNDLE = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
APP_DIR = BUNDLE / "app" if FROZEN else Path(__file__).resolve().parent
APP_SCRIPT = APP_DIR / "app.py"
IS_WINDOWS = sys.platform == "win32"
IS_MAC = sys.platform == "darwin"
WEBVIEW2_URL = "https://go.microsoft.com/fwlink/p/?LinkId=2124703"

_LOCAL = urllib.request.build_opener(urllib.request.ProxyHandler({}))


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


def log_path() -> Path:
    return data_dir() / "app.log"


def use_log_when_no_console() -> None:
    if sys.stdout is None or sys.stderr is None:
        stream = open(log_path(), "a", encoding="utf-8", buffering=1)
        sys.stdout = sys.stdout or stream
        sys.stderr = sys.stderr or stream


def tesseract_program() -> Path | None:
    name = "tesseract.exe" if IS_WINDOWS else "tesseract"
    candidates = [BUNDLE / "tesseract" / name, BUNDLE / "tesseract" / "bin" / name]
    if IS_WINDOWS:
        candidates += [
            Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "Tesseract-OCR" / name,
            Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "Tesseract-OCR" / name,
        ]
    else:
        candidates += [Path("/opt/homebrew/bin") / name, Path("/usr/local/bin") / name, Path("/opt/local/bin") / name]
    return next((c for c in candidates if c.is_file()), None)


def configure_runtime() -> None:
    if not IS_WINDOWS and not os.environ.get("TMPDIR"):
        import tempfile

        folder = data_dir() / "tmp"
        folder.mkdir(exist_ok=True)
        os.environ["TMPDIR"] = str(folder)
        tempfile.tempdir = None
    program = tesseract_program()
    if program is not None:
        os.environ["PATH"] = str(program.parent) + os.pathsep + os.environ.get("PATH", "")
        tessdata = BUNDLE / "tesseract" / "tessdata"
        if tessdata.is_dir():
            os.environ["TESSDATA_PREFIX"] = str(tessdata)
        try:
            import pytesseract

            pytesseract.pytesseract.tesseract_cmd = str(program)
        except ImportError:
            pass
    try:
        import certifi

        os.environ.setdefault("SSL_CERT_FILE", certifi.where())
    except ImportError:
        pass


def serve(port: int) -> None:
    use_log_when_no_console()

    def stop_with_parent() -> None:
        try:
            while os.read(0, 1024):
                pass
        except OSError:
            pass
        os._exit(0)

    threading.Thread(target=stop_with_parent, daemon=True).start()
    configure_runtime()
    os.chdir(APP_DIR)
    from streamlit.web import cli

    sys.argv = [
        "streamlit", "run", str(APP_SCRIPT),
        "--global.developmentMode=false",
        "--server.headless=true",
        "--server.address=127.0.0.1",
        f"--server.port={port}",
        "--server.fileWatcherType=none",
        "--server.runOnSave=false",
        "--browser.gatherUsageStats=false",
        "--server.showEmailPrompt=false",
    ]
    sys.exit(cli.main())


def free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return sock.getsockname()[1]


def healthy(port: int) -> bool:
    try:
        with _LOCAL.open(f"http://127.0.0.1:{port}/_stcore/health", timeout=1) as response:
            return response.read().strip() == b"ok"
    except Exception:
        return False


def start_server() -> tuple[subprocess.Popen, int]:
    port = free_port()
    command = [sys.executable, "--serve", str(port)] if FROZEN else [sys.executable, str(Path(__file__).resolve()), "--serve", str(port)]
    log = open(log_path(), "a", encoding="utf-8")
    log.write(f"\n=== {time.strftime('%Y-%m-%d %H:%M:%S')} starting on port {port}\n")
    log.flush()
    flags = subprocess.CREATE_NO_WINDOW if IS_WINDOWS else 0
    process = subprocess.Popen(command, cwd=APP_DIR, stdin=subprocess.PIPE, stdout=log, stderr=subprocess.STDOUT, creationflags=flags)
    return process, port


def wait_until_ready(process: subprocess.Popen, port: int, seconds: float = 180) -> bool:
    deadline = time.monotonic() + seconds
    while time.monotonic() < deadline:
        if process.poll() is not None:
            return False
        if healthy(port):
            return True
        time.sleep(0.25)
    return False


def stop_server(process: subprocess.Popen | None) -> None:
    if process is None:
        return
    try:
        if process.stdin:
            process.stdin.close()
        process.wait(timeout=5)
    except (OSError, subprocess.TimeoutExpired):
        process.terminate()
        try:
            process.wait(timeout=5)
        except subprocess.TimeoutExpired:
            process.kill()


def webview2_installed() -> bool:
    import winreg

    key = r"Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
    for root, path in ((winreg.HKEY_LOCAL_MACHINE, "SOFTWARE\\WOW6432Node\\" + key), (winreg.HKEY_LOCAL_MACHINE, "SOFTWARE\\" + key), (winreg.HKEY_CURRENT_USER, "Software\\" + key)):
        try:
            with winreg.OpenKey(root, path) as handle:
                version = winreg.QueryValueEx(handle, "pv")[0]
                if version and version != "0.0.0.0":
                    return True
        except OSError:
            continue
    return False


def mac_app_folder() -> Path | None:
    if not (IS_MAC and FROZEN):
        return None
    for parent in Path(sys.executable).resolve().parents:
        if parent.suffix == ".app":
            return parent
    return None


PAGE = """<!doctype html><html><head><meta charset="utf-8"><style>
html,body{{height:100%;margin:0;background:#EDEDED;color:#000;font:15px Montserrat,"Avenir Next","Segoe UI",sans-serif;letter-spacing:.03em}}
main{{height:100%;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:14px;text-align:center;padding:0 40px}}
h1{{margin:0;color:#112B74;font-weight:400;font-size:28px}}
p{{margin:0;color:rgba(0,0,0,.6);max-width:520px;line-height:1.6}}
</style></head><body><main><h1>{title}</h1><p>{text}</p></main></body></html>"""


def page(title: str, text: str) -> str:
    return PAGE.format(title=title, text=text)


def open_window() -> int:
    use_log_when_no_console()
    if IS_WINDOWS and not webview2_installed():
        import ctypes
        import webbrowser

        ctypes.windll.user32.MessageBoxW(
            None,
            "Microsoft Edge WebView2 is needed to show this app. The download page opens now. Install it, then open the app again.",
            APP_NAME, 0x40,
        )
        webbrowser.open(WEBVIEW2_URL)
        return 1

    app_folder = mac_app_folder()
    translocated = app_folder is not None and "/AppTranslocation/" in str(app_folder)
    if app_folder is not None and not translocated:
        subprocess.run(["/usr/bin/xattr", "-dr", "com.apple.quarantine", str(app_folder)], capture_output=True)

    import webview

    webview.settings["ALLOW_DOWNLOADS"] = True
    webview.settings["OPEN_EXTERNAL_LINKS_IN_BROWSER"] = True
    window = webview.create_window(
        APP_NAME, html=page(APP_NAME, "Starting. This takes a few seconds."),
        width=1320, height=900, min_size=(960, 640), background_color="#EDEDED", text_select=True,
    )
    server: dict[str, subprocess.Popen | None] = {"process": None}

    def boot() -> None:
        if translocated:
            window.load_html(page("Move the app to Applications", f"Drag {APP_NAME} from this folder to the Applications folder, then open it from there."))
            return
        for _ in range(3):
            process, port = start_server()
            server["process"] = process
            if wait_until_ready(process, port):
                window.load_url(f"http://127.0.0.1:{port}/")
                return
            stop_server(process)
        window.load_html(page("The app could not start", f"Close this window and open the app again. If it keeps happening, send the file {log_path()} to support."))

    try:
        webview.start(boot, private_mode=True)
    finally:
        stop_server(server["process"])
    return 0


def selftest(report: Path) -> int:
    results: dict[str, str] = {}
    configure_runtime()
    sys.path.insert(0, str(APP_DIR))

    def check(name: str, test) -> None:
        try:
            results[name] = "ok: " + str(test())
        except Exception as exc:
            import traceback

            where = " | ".join(line.strip() for line in traceback.format_exc().splitlines()[-8:-1])
            results[name] = f"FAIL: {type(exc).__name__}: {exc} | {where}"

    def tesseract_reads() -> str:
        import pytesseract
        from PIL import Image, ImageDraw, ImageFont

        image = Image.new("L", (1000, 160), 255)
        ImageDraw.Draw(image).text((20, 40), "PURCHASE ORDER 4517204", fill=0, font=ImageFont.load_default(size=64))
        text = pytesseract.image_to_string(image, config="--psm 7").strip()
        if "4517204" not in text:
            raise ValueError(f"read {text!r}")
        return f"{pytesseract.get_tesseract_version()} read {text!r}"

    def quote_document() -> str:
        from decimal import Decimal

        import quotes

        sender = {"company_name": "Simplex Sciences", "company_address": "206 Elm Street\nNew Haven, CT", "company_email": "a@b.c", "company_phone": "1",
                  "bank_name": "B", "account_number": "1", "direct_deposit_routing": "1", "wire_routing": "1", "swift_code": "S"}
        quote = quotes.Quote("Test", "2026-01-01", "Rep", "Company\n1 Road\nTown\nCountry", [quotes.QuoteLine("ss20 DNA Ladder (100µL)", 2, Decimal("169.00"))], "FedEx 2-Day Shipping", Decimal("20.00"))
        return f"{len(quotes.generate_quote_docx((APP_DIR / 'simplex_quote_template.docx').read_bytes(), quote, sender))} bytes"

    def app_runs() -> str:
        from streamlit.testing.v1 import AppTest

        test = AppTest.from_file(str(APP_SCRIPT), default_timeout=180).run()
        if test.exception:
            raise RuntimeError(test.exception[0].message)
        return "passcode screen rendered"

    def server_starts() -> str:
        process, port = start_server()
        try:
            if not wait_until_ready(process, port):
                raise RuntimeError(f"no answer; see {log_path()}")
            return f"answered on 127.0.0.1:{port}"
        finally:
            stop_server(process)

    def imports() -> str:
        import docx
        import docxtpl
        import generic_po
        import numpy
        import pandas
        import pymupdf
        import thermo_qty

        return ", ".join(m.__name__ for m in (docx, docxtpl, generic_po, numpy, pandas, pymupdf, thermo_qty))

    def ocr_word_boxes() -> str:
        import pytesseract
        from PIL import Image, ImageDraw, ImageFont

        image = Image.new("L", (1000, 160), 255)
        ImageDraw.Draw(image).text((20, 40), "PURCHASE ORDER", fill=0, font=ImageFont.load_default(size=64))
        data = pytesseract.image_to_data(image, output_type=pytesseract.Output.DICT, config="--psm 11")
        words = [w for w in data["text"] if w.strip()]
        if not any("PURCHASE" in w for w in words):
            raise ValueError(f"read {words!r}")
        return f"{len(words)} words with positions"

    def other_purchase_order() -> str:
        import generic_po
        import pymupdf

        pdf = pymupdf.open()
        page = pdf.new_page()
        rows = [(72, 72, "Bill To:"), (320, 72, "Ship To:"), (72, 88, "Example Labs"), (320, 88, "Example Labs"), (72, 104, "1 Main Street"),
                (320, 104, "2 Dock Road"), (72, 150, "P.O. #"), (72, 166, "4500123"), (72, 210, "Product"), (300, 210, "Unit Price"),
                (400, 210, "QTY"), (460, 210, "Sub Total"), (72, 228, "ss20 DNA Ladder (100uL)"), (300, 228, "$169.00"), (400, 228, "2"), (460, 228, "$338.00")]
        for x, y, text in rows:
            page.insert_text((x, y), text, fontsize=10)
        fields = generic_po.parse_document(pdf.tobytes(), use_ocr=False)
        item = fields["line_items"][0]
        if fields["customer_po_number"] != "4500123" or item["quantity"] != "2" or item["amount"] != "$338.00":
            raise ValueError(f"got {fields['customer_po_number']!r} {item!r}")
        return f"PO {fields['customer_po_number']}, {item['quantity']} x {item['item']}"

    check("imports", imports)
    check("tesseract", tesseract_reads)
    check("ocr word boxes", ocr_word_boxes)
    check("other purchase order", other_purchase_order)
    check("quote", quote_document)
    check("app", app_runs)
    check("server", server_starts)
    results["platform"] = f"{platform.system()} {platform.machine()} Python {platform.python_version()}"
    report.write_text(json.dumps(results, indent=2), encoding="utf-8")
    print(json.dumps(results, indent=2))
    return 0 if all(not v.startswith("FAIL") for v in results.values()) else 1


def main() -> int:
    if len(sys.argv) >= 3 and sys.argv[1] == "--serve":
        serve(int(sys.argv[2]))
        return 0
    if len(sys.argv) >= 2 and sys.argv[1] == "--selftest":
        return selftest(Path(sys.argv[2]) if len(sys.argv) >= 3 else data_dir() / "selftest.json")
    return open_window()


if __name__ == "__main__":
    sys.exit(main())
