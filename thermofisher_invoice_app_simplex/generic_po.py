from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation

import pymupdf

try:
    import pytesseract
    from PIL import Image
except Exception:
    pytesseract = None
    Image = None

SELF_NAMES = re.compile(r"\bsimplex\s+sciences?\b", re.I)


@dataclass
class Word:
    x0: float
    y0: float
    x1: float
    y1: float
    text: str
    page: int

    @property
    def cx(self) -> float:
        return (self.x0 + self.x1) / 2

    @property
    def cy(self) -> float:
        return (self.y0 + self.y1) / 2


@dataclass
class Row:
    words: list[Word] = field(default_factory=list)

    @property
    def text(self) -> str:
        return " ".join(w.text for w in self.words)

    @property
    def y0(self) -> float:
        return min(w.y0 for w in self.words)

    @property
    def y1(self) -> float:
        return max(w.y1 for w in self.words)

    @property
    def page(self) -> int:
        return self.words[0].page

    def slice(self, x_from: float, x_to: float) -> list[Word]:
        return [w for w in self.words if x_from <= w.cx < x_to]


def page_words(file_bytes: bytes, use_ocr: bool = True, rotation: str = "") -> list[Word]:
    words: list[Word] = []
    angles = [int(a) for a in rotation.split(",") if a.strip().isdigit()] if rotation else []
    with pymupdf.open(stream=file_bytes, filetype="pdf") as pdf:
        for number, page in enumerate(pdf):
            layer = page.get_text("words")
            if sum(len(w[4]) for w in layer) >= 40:
                words += [Word(w[0], w[1], w[2], w[3], w[4], number) for w in layer if w[4].strip()]
                continue
            if not use_ocr or pytesseract is None or Image is None:
                continue
            words += _ocr_words(page, number, angles[number] if number < len(angles) else 0)
    return words


def _ocr_words(page, number: int, angle: int) -> list[Word]:
    dpi = 300
    image = Image.open(io.BytesIO(page.get_pixmap(dpi=dpi).tobytes("png")))
    if angle:
        image = image.rotate(angle, expand=True)
    try:
        data = pytesseract.image_to_data(image, output_type=pytesseract.Output.DICT, config="--psm 11")
    except Exception:
        return []
    scale = 72 / dpi
    out = []
    for i, text in enumerate(data.get("text") or []):
        text = (text or "").strip()
        if not text or float(data["conf"][i]) < 30:
            continue
        x, y, w, h = (data[k][i] * scale for k in ("left", "top", "width", "height"))
        out.append(Word(x, y, x + w, y + h, text, number))
    return out


def group_rows(words: list[Word]) -> list[Row]:
    rows: list[Row] = []
    for word in sorted(words, key=lambda w: (w.page, w.cy, w.x0)):
        if rows and rows[-1].page == word.page:
            last = rows[-1]
            height = max(1.0, min(last.y1 - last.y0, word.y1 - word.y0))
            if abs(word.cy - (last.y0 + last.y1) / 2) <= height * 0.45:
                last.words.append(word)
                continue
        rows.append(Row([word]))
    for row in rows:
        row.words.sort(key=lambda w: w.x0)
    return rows


_MONEY = re.compile(r"^\(?[$€£]?\(?(\d{1,3}(?:[.,\s]\d{3})*[.,]\d{2}|\d+[.,]\d{2})\)?[$€£]?\)?$")


def parse_amount(text: str) -> Decimal | None:
    token = text.strip().replace("USD", "").replace("EUR", "").strip()
    m = _MONEY.match(token)
    if not m:
        return None
    number = m.group(1).replace(" ", "")
    if re.search(r",\d{2}$", number):
        number = number.replace(".", "").replace(",", ".")
    else:
        number = number.replace(",", "")
    try:
        return Decimal(number).quantize(Decimal("0.01"))
    except InvalidOperation:
        return None


def money_in(words: list[Word]) -> list[Decimal]:
    values = []
    for w in words:
        value = parse_amount(w.text)
        if value is not None:
            values.append(value)
    return values


def fmt(value: Decimal) -> str:
    return f"${value:,.2f}"


def detect_currency(text: str) -> str:
    if "€" in text or re.search(r"\bEUR\b", text):
        return "EUR"
    if "£" in text or re.search(r"\bGBP\b", text):
        return "GBP"
    return "USD"


PO_LABELS = re.compile(
    r"(?:\bcustomer\s+)?\b(?:p\.?\s?o\.?\s*(?:#|no\.?|nr\.?|number)|purchase\s+order\s*(?:#|no\.?|number)?"
    r"|our\s+order\s+(?:no\.?|number)|order\s+(?:#|no\.?|number))\s*:?",
    re.I,
)
SHIP_LABELS = re.compile(r"\b(?:ship\s*-?\s*to|deliver(?:y)?\s+to|delivery\s+address|shipping\s+address|consignee)\b\s*:?", re.I)
BILL_LABELS = re.compile(r"\b(?:bill\s*-?\s*to|invoice\s+to|invoice\s+address|billing\s+address|sold\s+to)\b\s*:?", re.I)
BOTH_LABELS = re.compile(r"\bdelivery\s*/\s*invoice\s+address\b\s*:?", re.I)
ANY_LABEL = re.compile(
    r"^(?:[A-Za-z][\w./#&'-]*\s){0,4}[A-Za-z][\w./#&'-]*\s*:|" + SHIP_LABELS.pattern + "|" + BILL_LABELS.pattern,
    re.I,
)


def _label_span(row: Row, pattern: re.Pattern) -> tuple[float, float, list[Word]] | None:
    text = row.text
    if pattern is not BOTH_LABELS and BOTH_LABELS.search(text):
        return None
    m = pattern.search(text)
    if not m:
        return None
    pos, start_word, end_word = 0, None, None
    for i, w in enumerate(row.words):
        w_start, w_end = pos, pos + len(w.text)
        if start_word is None and w_end > m.start():
            start_word = i
        if w_start < m.end():
            end_word = i
        pos = w_end + 1
    if start_word is None or end_word is None:
        return None
    label = row.words[start_word : end_word + 1]
    return label[0].x0, label[-1].x1, row.words[end_word + 1 :]


def _is_date(text: str) -> bool:
    return bool(re.fullmatch(r"\d{1,2}[./-]\d{1,2}[./-]\d{2,4}|\d{4}-\d{2}-\d{2}", text))


def _plausible_po(text: str) -> bool:
    token = text.strip(" :#")
    return len(token) >= 3 and any(ch.isdigit() for ch in token) and not _is_date(token) and parse_amount(token) is None


def find_po_number(rows: list[Row]) -> str:
    for i, row in enumerate(rows):
        span = _label_span(row, PO_LABELS)
        if not span:
            continue
        x0, x1, after = span
        same_line = [w.text for w in after if abs(w.x0 - x1) < 200]
        if same_line and _plausible_po(same_line[0]):
            return same_line[0].strip(" :#")
        for below in rows[i + 1 : i + 3]:
            if below.page != row.page or below.y0 - row.y1 > 30:
                break
            under = [w for w in below.words if w.x1 > x0 - 8 and w.x0 < x1 + 8]
            if under and _plausible_po(under[0].text):
                return under[0].text.strip(" :#")
    return ""


def _column_limit(rows: list[Row], index: int, x0: float) -> float:
    row = rows[index]
    limit = 10_000.0
    for other in rows[max(0, index - 1) : index + 2]:
        if other.page != row.page:
            continue
        for pattern in (SHIP_LABELS, BILL_LABELS, BOTH_LABELS):
            span = _label_span(other, pattern)
            if span and span[0] > x0 + 40:
                limit = min(limit, span[0] - 4)
    return limit


def find_block(rows: list[Row], pattern: re.Pattern) -> tuple[str, str]:
    for i, row in enumerate(rows):
        span = _label_span(row, pattern)
        if not span:
            continue
        x0, x1, after = span
        limit = _column_limit(rows, i, x0)
        lines: list[str] = []
        same_line = " ".join(w.text for w in after if w.x0 < limit).strip()
        if same_line:
            lines.append(same_line)
        last_y = row.y1
        email = ""
        for below in rows[i + 1 : i + 12]:
            if below.page != row.page:
                break
            words = [w for w in below.words if x0 - 12 <= w.x0 < limit]
            if not words:
                continue
            height = max(6.0, below.y1 - below.y0)
            if below.y0 - last_y > height * 2.2:
                break
            text = " ".join(w.text for w in words).strip()
            found_email = re.search(r"[\w.+-]+@[\w-]+\.[\w.]+", text)
            if found_email and re.search(r"invoice|bill|e-?mail", text, re.I):
                email = found_email.group(0)
                break
            if ANY_LABEL.match(text) and lines:
                break
            if re.fullmatch(r"[\s|/-]*", text):
                continue
            lines.append(text)
            last_y = below.y1
        lines = [re.sub(r"\s+", " ", l).strip(" ,|") for l in lines if l.strip(" ,|")]
        if lines:
            return "\n".join(lines[:8]), email
    return "", ""


def organization(block: str) -> str:
    name = []
    for line in block.splitlines():
        if re.search(r"\d", line):
            break
        name.append(line)
        if len(name) == 2:
            break
    return " ".join(name).strip()


HEADERS = {
    "quantity": re.compile(r"^(?:qty|quantity|qté|menge|anzahl|units?)\.?$", re.I),
    "description": re.compile(r"^(?:description|item|product|article|artikel|bezeichnung|material)s?$", re.I),
    "unit_price": re.compile(r"^(?:unit\s*(?:price|cost)|price|rate|each|preis|einzelpreis)$", re.I),
    "amount": re.compile(r"^(?:amount|total|sub\s*total|extended(?:\s*(?:cost|price))?|ext\.?\s*price|line\s*total|betrag|gesamt)$", re.I),
}
STOP = re.compile(r"\b(?:sub\s*total|order\s+total|invoice\s+total|grand\s+total|total\s*:|tax\s*:|amount\s+due|remit\s+to|bill\s+acct)\b", re.I)
SHIPPING = re.compile(r"\b(?:shipping|freight|delivery\s+charge|handling|postage)\b", re.I)


@dataclass
class Column:
    name: str
    title: str
    x0: float
    x1: float

    @property
    def cx(self) -> float:
        return (self.x0 + self.x1) / 2


def _header_columns(row: Row) -> list[Column]:
    groups: list[list[Word]] = []
    for w in row.words:
        if groups and w.x0 - groups[-1][-1].x1 < 12:
            groups[-1].append(w)
        else:
            groups.append([w])
    columns = []
    for group in groups:
        title = " ".join(w.text for w in group).strip(" :#")
        kind = next((k for k, p in HEADERS.items() if p.match(title)), "other")
        columns.append(Column(kind, title, group[0].x0, group[-1].x1))
    return columns


def _find_header(rows: list[Row]) -> tuple[int, list[Column]] | None:
    for i, row in enumerate(rows):
        height = row.y1 - row.y0
        center = (row.y0 + row.y1) / 2
        nearby = [w for r in rows[i + 1 : i + 2] if r.page == row.page and abs((r.y0 + r.y1) / 2 - center) < height * 0.6 for w in r.words]
        merged = Row(sorted(row.words + nearby, key=lambda w: w.x0))
        columns = _header_columns(merged)
        kinds = {c.name for c in columns}
        if len(kinds & {"quantity", "description", "unit_price", "amount"}) >= 3 or {"description", "amount"} <= kinds:
            return i, columns
    return None


def _assign(row: Row, columns: list[Column]) -> dict[int, list[Word]]:
    numeric = [k for k, c in enumerate(columns) if c.name in ("quantity", "unit_price", "amount")]
    cells: dict[int, list[Word]] = {}
    for w in row.words:
        is_money = parse_amount(w.text) is not None or w.text in ("$", "€", "£", "($")
        if is_money and numeric:
            index = min(numeric, key=lambda k: abs(columns[k].cx - w.cx))
        elif _qty(w.text) and any(columns[k].name == "quantity" for k in numeric) and w.x0 >= min(columns[k].x0 for k in numeric) - 20:
            index = min(numeric, key=lambda k: abs(columns[k].cx - w.cx))
        else:
            starts = [k for k, c in enumerate(columns) if c.x0 - 20 <= w.x0]
            index = starts[-1] if starts else 0
            if columns[index].name in ("quantity", "unit_price", "amount"):
                text_columns = [k for k in range(index) if columns[k].name in ("description", "other")]
                index = text_columns[-1] if text_columns else index
        cells.setdefault(index, []).append(w)
    return cells


def _qty(text: str) -> str:
    m = re.fullmatch(r"(\d{1,5})(?:[.,]0+)?(?:\s*(?:ea|each|pcs?|pc|x|units?))?", text.strip(), re.I)
    return str(int(m.group(1))) if m else ""


def _item(qty: str, description: str, unit: Decimal | None, amount: Decimal | None) -> dict[str, str]:
    q = int(qty) if qty else 0
    if unit is None and amount is not None and q:
        unit = (amount / q).quantize(Decimal("0.01"))
    if amount is None and unit is not None:
        amount = (unit * (q or 1)).quantize(Decimal("0.01"))
    if not qty and unit and amount:
        ratio = amount / unit
        if ratio == ratio.to_integral_value():
            qty = str(int(ratio))
    return {
        "quantity": qty,
        "item": re.sub(r"\s+", " ", description).strip(" ,|"),
        "unit_price": fmt(unit) if unit is not None else "",
        "amount": fmt(amount) if amount is not None else "",
    }


def items_from_table(rows: list[Row]) -> tuple[list[dict[str, str]], Decimal | None]:
    found = _find_header(rows)
    if not found:
        return [], None
    start, columns = found
    header_page = rows[start].page
    items: list[dict[str, str]] = []
    shipping: Decimal | None = None
    last_row: Row | None = None
    for row in rows[start + 1 :]:
        if row.page != header_page and _find_header([row]) is None and not money_in(row.words):
            continue
        if STOP.search(row.text):
            break
        cells = _assign(row, columns)
        by_kind: dict[str, list[Word]] = {}
        text_parts: list[tuple[str, str]] = []
        for index, words in sorted(cells.items()):
            column = columns[index]
            by_kind.setdefault(column.name, []).extend(words)
            if column.name in ("description", "other"):
                text = " ".join(w.text for w in words)
                if not re.fullmatch(r"[\d\s#.-]+", text):
                    text_parts.append((column.name, text))
        unit_values = money_in(by_kind.get("unit_price", []))
        amount_values = money_in(by_kind.get("amount", []))
        if not unit_values and not amount_values:
            words = [t for kind, t in text_parts if kind == "description"]
            close = last_row is not None and row.page == last_row.page and row.y0 - last_row.y1 < (last_row.y1 - last_row.y0) * 1.2
            if items and words and close and not SHIPPING.search(" ".join(words)) and not re.match(r"page\s+\d", row.text, re.I):
                items[-1]["item"] = f"{items[-1]['item']} {' '.join(words)}".strip()
                last_row = row
            continue
        description = " ".join(t for kind, t in text_parts if kind == "description")
        extras = [t for kind, t in text_parts if kind == "other"]
        if extras and description:
            description = f"{description} ({', '.join(extras)})"
        elif extras:
            description = ", ".join(extras)
        qty = _qty(" ".join(w.text for w in by_kind.get("quantity", [])))
        unit = unit_values[0] if unit_values else None
        amount = amount_values[-1] if amount_values else None
        last_row = row
        if SHIPPING.search(description):
            shipping = amount if amount is not None else unit
            continue
        items.append(_item(qty, description, unit, amount))
    return [i for i in items if i["item"] or i["amount"]], shipping


def items_from_rows(rows: list[Row]) -> list[dict[str, str]]:
    items = []
    for row in rows:
        values = [(i, parse_amount(w.text)) for i, w in enumerate(row.words)]
        money = [(i, v) for i, v in values if v is not None]
        if not money or len(money) > 2 or SHIPPING.search(row.text) or STOP.search(row.text):
            continue
        first_money = money[0][0]
        before = row.words[:first_money]
        qty_index = next((i for i in range(len(before) - 1, -1, -1) if _qty(before[i].text)), None)
        if qty_index is None:
            continue
        description_words = [w.text for w in before[:qty_index] if w.text not in ("$", "€", "£")]
        description = " ".join(description_words)
        if len(re.findall(r"[A-Za-z]", description)) < 4:
            continue
        unit = money[0][1]
        amount = money[-1][1] if len(money) == 2 else None
        items.append(_item(_qty(before[qty_index].text), description, unit, amount))
    return items


def parse_document(file_bytes: bytes, use_ocr: bool = True, rotation: str = "") -> dict:
    words = page_words(file_bytes, use_ocr=use_ocr, rotation=rotation)
    rows = group_rows(words)
    text = "\n".join(r.text for r in rows)

    both, both_email = find_block(rows, BOTH_LABELS)
    ship_to, _ = find_block(rows, SHIP_LABELS)
    bill_to, bill_email = find_block(rows, BILL_LABELS)
    ship_to = ship_to or both
    bill_to = bill_to or both
    email = bill_email or both_email
    if bill_to and email and email not in bill_to:
        bill_to += f"\nEmail: {email}"

    items, shipping = items_from_table(rows)
    if not items:
        items = items_from_rows(rows)

    customer = organization(bill_to) or organization(ship_to)
    notes = []
    if customer and SELF_NAMES.search(customer):
        notes.append("This document is billed to Simplex Sciences, so it looks like a supplier invoice rather than a customer purchase order.")
    if not items:
        notes.append("No product lines were found. Add them in the table.")

    return {
        "invoice_issued_to": customer,
        "customer_po_number": find_po_number(rows),
        "bill_to": bill_to,
        "ship_to": ship_to,
        "currency": detect_currency(text),
        "line_items": items or [{"quantity": "", "item": "", "unit_price": "", "amount": ""}],
        "shipping_cost": f"{shipping:.2f}" if shipping else "",
        "notes": notes,
        "raw_text": text,
    }
