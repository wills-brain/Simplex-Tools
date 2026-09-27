# Simplex Invoices and Quotes

Internal Simplex Sciences app for invoices and quotes.

**Thermo Fisher invoices.** Reads Thermo Fisher purchase order PDFs, including scanned copies, and produces the Word invoice after you review each field.

**Quotes.** Builds a Word quote in the same layout as the quotes we send (for example `2025_10_28_MERCK.docx`). Products and list prices come from simplexsciences.com. The FedEx shipping price is entered by hand.

## Start the app

1. On GitHub, click **Code**, then **Download ZIP**.
2. Unzip the file.
   * **Windows:** first right click the ZIP file, choose **Properties**, tick **Unblock** and click **OK**. Then right click the ZIP file and choose **Extract All**. Do not run anything from inside the ZIP file.
3. Double click the launcher in the unzipped folder.
   * **Mac:** `Start Invoice App (Mac).command`
   * **Windows:** `Start Invoice App (Windows).bat`
4. The app opens in your browser. Keep the launcher window open while you use the app. Close the window to stop the app.

The first start installs everything the app needs. This is about 1 GB and takes 5 to 15 minutes with an internet connection. Later starts take a few seconds.

### First start on a Mac

macOS blocks files downloaded from the internet the first time they are opened (macOS 15 Sequoia and later).

1. Double click the launcher. macOS reports that it was not opened. Click **Done**.
2. Open **System Settings**, then **Privacy & Security**. Under **Security**, click **Open Anyway** next to the launcher's name. The button is shown for about an hour after step 1.
3. Enter your Mac password and confirm with **Open Anyway**.
4. If Terminal asks to access the Downloads folder, click **Allow**. You can also move the unzipped folder to Documents first.

Repeat these steps after downloading a new copy.

### First start on Windows

If **Unblock** was not ticked before unzipping, Windows may show **Open File Security Warning** (click **Run**) or **Windows protected your PC** (click **More info**, then **Run anyway**). With **Smart App Control** turned on, the launcher is blocked. Unzip again after ticking **Unblock** on the ZIP file.

### Requirements

**Python 3.10 to 3.14, 64 bit.** If Python is not found:

* **Mac:** the launcher opens python.org. Install Python 3.13 (Python 3.12.10 on Intel Macs), then double click the launcher again.
* **Windows:** the launcher offers to install Python 3.13 for your account with winget. No administrator rights are needed. You can also install it from python.org.

**Tesseract OCR,** for scanned PDFs only. PDFs with selectable text work without it. The launcher tells you if it is missing.

* **Mac:** run `brew install tesseract` in Terminal. This needs [Homebrew](https://brew.sh).
* **Windows:** run `winget install --id tesseract-ocr.tesseract -e` in Command Prompt, or use the installer from the [Tesseract releases page](https://github.com/tesseract-ocr/tesseract/releases).

The app can only be opened from your own computer.

## Quotes

1. **Customer.** Who the quote is issued to, issue date, sales rep, shipping address, and whether the order ships within the United States.
2. **Products.** One line per product with the dye option, quantity and unit price. Unit prices start at the list price on simplexsciences.com and can be changed. Choose **Custom product** for anything not in the catalog, and **Add product** for more lines.
3. **Shipping.** Choose the FedEx service, then use the FedEx rate calculator link to rate the package from our address to the shipping address. Enter the price, or leave it empty to quote the total without shipping.
4. **Review.** Check the totals and click **Download quote**. The file is named like `2026_09_26_MERCK.docx`.

The bundled `simplex_quote_template.docx` is used unless you upload a different template under **Quote template**. A quote we sent before also works as a template.

## Troubleshooting

* **Start over:** delete the app's setup folder, then start the launcher again.
  * Mac: `~/Library/Application Support/SimplexInvoiceApp`
  * Windows: `%LOCALAPPDATA%\SimplexInvoiceApp`
* **Setup log:** `setup.log` in the same folder.
* **Launcher started twice:** the second window waits for the first one and then opens the browser.
* **Mac reports that the launcher could not be executed:** the ZIP file was unpacked by another tool. Run `chmod +x "Start Invoice App (Mac).command"` in Terminal from the unzipped folder.

## Development

```bash
cd thermofisher_invoice_app_simplex
python3 launcher.py            # same as the double click launchers
python3 launcher.py --reset    # rebuild the app's Python environment
```

To manage your own environment instead, install `requirements.txt` (and `requirements-ocr.txt` for the optional EasyOCR helper; add `-c constraints-intel-mac.txt` on Intel Macs), then run `streamlit run app.py` in `thermofisher_invoice_app_simplex`.

| File | Purpose |
| --- | --- |
| `app.py` | The app, both tabs |
| `quotes.py` | Quote data, the simplexsciences.com catalog, and Word quote output |
| `launcher.py` | Setup and start, used by both launchers |
| `simplex_invoice_template.docx`, `simplex_quote_template.docx` | Word templates |
