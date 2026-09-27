from __future__ import annotations

import io
import re
from collections import Counter
from decimal import Decimal, InvalidOperation

import numpy as np
import pymupdf
from PIL import Image, ImageFilter

try:
    import pytesseract
except Exception:
    pytesseract = None

MONEY = re.compile(r"\d{1,3}(?:,\d{3})*\.\d{2}")
DIGIT_MAP = str.maketrans({"S": "5", "s": "5", "O": "0", "o": "0", "D": "0", "l": "1", "I": "1", "i": "1", "B": "8", "Z": "2", "z": "2", "g": "9"})
QTY = re.compile(r"[1-9]\d{0,3}")


def item_line(text: str) -> str:
    for line in text.splitlines():
        if re.search(r"ladder|dna|ss50|ss20|s660", line, re.I) and MONEY.search(line):
            return line
    return ""


def qty_from_cell(text: str) -> str:
    cells = [c.strip() for c in item_line(text).split("|")]
    described = [i for i, c in enumerate(cells) if re.search(r"LADDER|DNA", c, re.I)]
    if not described or described[-1] + 1 >= len(cells):
        return ""
    raw = re.sub(r"[^0-9A-Za-z]", "", cells[described[-1] + 1]).translate(DIGIT_MAP)
    return raw if QTY.fullmatch(raw) else ""


def _money_cells(text: str) -> list[Decimal]:
    cells = [c.strip() for c in item_line(text).split("|")]
    return [Decimal(c.replace(",", "")) for c in cells if re.fullmatch(r"\d{1,3}(?:,\d{3})*\.\d{2}", c)]


def qty_from_prices(text: str) -> str:
    money = _money_cells(text)
    if len(money) < 2 or money[-2] <= 0:
        return ""
    quantity = money[-1] / money[-2]
    return str(int(quantity)) if quantity == quantity.to_integral_value() and 1 <= quantity <= 9999 else ""


def _groups(values: np.ndarray, gap: int = 2) -> list[tuple[int, int]]:
    if values.size == 0:
        return []
    out, start, prev = [], int(values[0]), int(values[0])
    for value in values[1:]:
        value = int(value)
        if value - prev > gap:
            out.append((start, prev))
            start = value
        prev = value
    out.append((start, prev))
    return out


def _erase_rows(mask: np.ndarray, min_run: int, gap: int = 3) -> np.ndarray:
    mask = mask.copy()
    for y in range(mask.shape[0]):
        xs = np.flatnonzero(mask[y])
        if xs.size == 0:
            continue
        splits = np.flatnonzero(np.diff(xs) > gap + 1)
        for start, end in zip(np.r_[xs[0], xs[splits + 1]], np.r_[xs[splits], xs[-1]]):
            if end - start + 1 >= min_run:
                mask[y, start : end + 1] = False
    return mask


def _erase_columns(mask: np.ndarray, fraction: float = 0.6) -> np.ndarray:
    mask = mask.copy()
    mask[:, np.flatnonzero(mask.sum(axis=0) >= fraction * mask.shape[0])] = False
    return mask


def _drop_specks(mask: np.ndarray, min_pixels: int) -> np.ndarray:
    height, width = mask.shape
    seen = np.zeros_like(mask)
    out = mask.copy()
    for y0, x0 in zip(*np.nonzero(mask)):
        if seen[y0, x0]:
            continue
        stack, component = [(y0, x0)], []
        seen[y0, x0] = True
        while stack:
            y, x = stack.pop()
            component.append((y, x))
            for yy, xx in ((y - 1, x), (y + 1, x), (y, x - 1), (y, x + 1)):
                if 0 <= yy < height and 0 <= xx < width and mask[yy, xx] and not seen[yy, xx]:
                    seen[yy, xx] = True
                    stack.append((yy, xx))
        if len(component) < min_pixels:
            for y, x in component:
                out[y, x] = False
    return out


def _page_words(image: Image.Image) -> list[dict]:
    data = pytesseract.image_to_data(image, output_type=pytesseract.Output.DICT, config="--psm 11")
    words = []
    for i, text in enumerate(data["text"]):
        text = (text or "").strip()
        if text:
            words.append({"t": text, "x0": data["left"][i], "y0": data["top"][i], "x1": data["left"][i] + data["width"][i], "y1": data["top"][i] + data["height"][i]})
    return words


def _qty_header(words: list[dict]) -> dict | None:
    for word in words:
        text = re.sub(r"[^A-Z0-9?]", "", word["t"].upper())
        if re.fullmatch(r"Q[TI1][YV?]", text) or text in ("QTY", "QTV", "OTY"):
            return word
    for word in words:
        if re.sub(r"[^A-Z]", "", word["t"].upper()) in ("UOM", "JOM", "TOM", "UO", "UOMUNIT"):
            return {**word, "uom": True}
    return None


def _qty_cell(image: Image.Image, words: list[dict]) -> np.ndarray | None:
    header = _qty_header(words)
    if header is None:
        return None
    height = max(10, header["y1"] - header["y0"])
    top, bottom = header["y0"], min(image.height, int(header["y1"] + 9 * height))
    dark = np.asarray(image.crop((0, top, image.width, bottom))) < 128
    rules = np.array([(a + b) // 2 for a, b in _groups(np.flatnonzero(dark.sum(axis=0) / dark.shape[0] > 0.3))])
    if rules.size < 2:
        return None
    if header.get("uom"):
        left_of = rules[rules < header["x0"]]
        if left_of.size < 2:
            return None
        left, right = left_of[-2], left_of[-1]
    else:
        center = (header["x0"] + header["x1"]) // 2
        left_of, right_of = rules[rules < center], rules[rules > center]
        if left_of.size == 0 or right_of.size == 0:
            return None
        left, right = left_of[-1], right_of[0]
    cell = dark[int(header["y1"] + 0.4 * height) - top :, int(left + 4) : int(right - 4)]
    cell = _erase_rows(cell, min_run=int(0.6 * cell.shape[1]))
    cell = _erase_columns(cell)
    cell = _drop_specks(cell, max(4, int(0.02 * height * height)))
    rows = np.flatnonzero(cell.any(axis=1))
    if rows.size == 0:
        return None
    splits = np.flatnonzero(np.diff(rows) > max(3, height // 3))
    return cell[rows[0] : (rows[splits[0]] if splits.size else rows[-1]) + 1]


def _read_digits(mask: np.ndarray) -> str:
    ys, xs = np.flatnonzero(mask.any(axis=1)), np.flatnonzero(mask.any(axis=0))
    if ys.size == 0:
        return ""
    glyphs = mask[ys[0] : ys[-1] + 1, xs[0] : xs[-1] + 1]
    padded = np.zeros((glyphs.shape[0] + 24, glyphs.shape[1] + 24), bool)
    padded[12:-12, 12:-12] = glyphs
    image = Image.fromarray(np.where(padded, 0, 255).astype(np.uint8), "L")
    image = image.resize((max(1, int(image.width * 1.5)), max(1, int(image.height * 1.5))), Image.LANCZOS)
    image = image.filter(ImageFilter.GaussianBlur(2)).point(lambda v: 0 if v < 128 else 255)
    raw = re.sub(r"\s+", "", pytesseract.image_to_string(image, config="--psm 7 -c tessedit_char_whitelist=0123456789"))
    return raw if QTY.fullmatch(raw) else ""


def qty_from_scan(file_bytes: bytes, angle: int) -> str:
    with pymupdf.open(stream=file_bytes, filetype="pdf") as pdf:
        page = pdf[0]
        pix = page.get_pixmap(dpi=300, colorspace=pymupdf.csGRAY)
        rendered = Image.frombytes("L", (pix.width, pix.height), pix.samples)
        images = page.get_images(full=True)
        if len(images) != 1:
            return ""
        native = Image.open(io.BytesIO(pdf.extract_image(images[0][0])["image"])).convert("L")
        page_rotation = page.rotation % 360
    if angle:
        rendered = rendered.rotate(angle, expand=True, fillcolor=255)
    if page_rotation:
        native = native.rotate(360 - page_rotation, expand=True)
    if angle:
        native = native.rotate(angle, expand=True, fillcolor=255)
    native = native.point(lambda v: 0 if v < 128 else 255)
    doubled = native.resize((native.width * 2, native.height * 2), Image.NEAREST)
    fx, fy = doubled.width / rendered.width, doubled.height / rendered.height
    words = [{**w, "x0": int(w["x0"] * fx), "x1": int(w["x1"] * fx), "y0": int(w["y0"] * fy), "y1": int(w["y1"] * fy)} for w in _page_words(rendered)]
    cell = _qty_cell(doubled, words)
    return _read_digits(cell) if cell is not None else ""


def vote_quantity(text: str, file_bytes: bytes, rotation: str) -> str:
    if pytesseract is None:
        return ""
    try:
        angle = int(rotation.split(",")[0]) if rotation[:1].isdigit() else 0
        scan = qty_from_scan(file_bytes, angle)
    except Exception:
        scan = ""
    votes = Counter(v for v in (qty_from_prices(text), qty_from_cell(text), scan) if v)
    top = votes.most_common(1)
    return top[0][0] if top and top[0][1] >= 2 else ""


def unit_price_for(text: str, quantity: str) -> Decimal | None:
    money = _money_cells(text)
    if not money or not quantity:
        return None
    try:
        count = Decimal(quantity)
    except InvalidOperation:
        return None
    extended = money[-1]
    unit = money[-2] if len(money) >= 2 else None
    return (extended / count).quantize(Decimal("0.01")) if unit is None or unit * count != extended else unit
