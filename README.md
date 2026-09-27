# Simplex Invoices and Quotes

Internal Simplex Sciences app for invoices and quotes. It opens in its own window on Mac and Windows.

**Invoices.** Reads purchase orders as PDFs, including scanned Thermo Fisher POs, and produces the Word invoice after you review each field. Thermo Fisher POs use the Thermo Fisher reader; POs and order letters from other customers are read by their labels (PO number, Bill To, Ship To, and the item table).

**Quotes.** Builds a Word quote in the same layout as the quotes we send (for example `2025_10_28_MERCK.docx`). Products and list prices come from simplexsciences.com. The FedEx shipping price is entered by hand.

## Install

Download the file for your computer from the [latest release](https://github.com/wills-brain/Invoice-Generation/releases/latest):

| Computer | File |
| --- | --- |
| Mac with Apple silicon (M1 or newer) | `Simplex-Invoices-and-Quotes-...-macos-apple-silicon.zip` |
| Mac with an Intel processor | `Simplex-Invoices-and-Quotes-...-macos-intel.zip` |
| Windows 10 or 11 | `Simplex-Invoices-and-Quotes-...-windows-setup.exe` |

Everything the app needs is included, including the OCR program for scanned PDFs. The Mac app needs macOS 14 Sonoma or newer.

### Mac

1. Double click the zip file, then drag **Simplex Invoices and Quotes** into the **Applications** folder.
2. Open it from Applications. macOS reports that it was not opened. Click **Done**.
3. Open **System Settings**, then **Privacy & Security**. Under **Security**, click **Open Anyway** next to the app's name. The button is shown for about an hour after step 2.
4. Enter your Mac password and confirm with **Open Anyway**.

After this the app opens normally. Repeat these steps after installing a new version.

### Windows

1. Run the setup file. If Windows shows **Windows protected your PC**, click **More info**, then **Run anyway**.
2. Setup installs the app for your account (no administrator rights needed), adds it to the Start menu, and installs Microsoft Edge WebView2 if the computer does not have it.

If **Smart App Control** is turned on, Windows blocks the setup file. It can be turned off in **Windows Security**, **App & browser control**, **Smart App Control settings**.

## Quotes

1. **Customer.** Who the quote is issued to, issue date, sales rep, shipping address, and whether the order ships within the United States.
2. **Products.** One line per product with the dye option, quantity and unit price. Unit prices start at the list price on simplexsciences.com and can be changed. Choose **Custom product** for anything not in the catalog, and **Add product** for more lines.
3. **Shipping.** Choose the FedEx service, then use the FedEx rate calculator link to rate the package from our address to the shipping address. Enter the price, or leave it empty to quote the total without shipping.
4. **Review.** Check the totals and click **Download quote**. The file is named like `2026_09_26_MERCK.docx`.

The bundled `simplex_quote_template.docx` is used unless you upload a different template under **Quote template**. A quote we sent before also works as a template.

## Troubleshooting

* **App log:** `~/Library/Application Support/SimplexInvoiceApp/app.log` on a Mac, `%LOCALAPPDATA%\SimplexInvoiceApp\app.log` on Windows.
* **Mac says the app is damaged:** the zip was unpacked by another program. Download it again and double click the zip in Finder.
* **The DR number:** Thermo Fisher order numbers are too faint on the scans to read reliably, so the field starts as "DR". Type the digits from the PO.

## Running from the source code

The launchers in this folder run the app from the source code instead of the installed app. They need Python 3.10 to 3.14 and set everything else up on the first start.

* **Mac:** double click `Start Invoice App (Mac).command`
* **Windows:** double click `Start Invoice App (Windows).bat`

Scanned PDFs then need Tesseract OCR installed: `brew install tesseract` on a Mac, `winget install --id tesseract-ocr.tesseract -e` on Windows.

## Development

```bash
cd thermofisher_invoice_app_simplex
python3 launcher.py --setup-only
python3 desktop.py              # the app window
python3 desktop.py --selftest   # checks OCR, documents and the app server
```

The installers are built by GitHub Actions (`.github/workflows/desktop.yml`) for every pull request, and published as a release when a version tag such as `v1.4.0` is pushed.

| File | Purpose |
| --- | --- |
| `app.py` | The app, both tabs |
| `generic_po.py` | Reads purchase orders from customers other than Thermo Fisher |
| `thermo_qty.py` | Reads the quantity on scanned Thermo Fisher POs when the text is unclear |
| `quotes.py` | Quote data, the simplexsciences.com catalog, and Word quote output |
| `desktop.py` | The app window and the installed app's entry point |
| `launcher.py` | Setup and start from the source code, used by both launchers |
| `packaging/` | App icons, the PyInstaller spec, and the Windows installer script |
