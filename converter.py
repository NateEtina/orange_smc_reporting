"""CSV -> XLSX conversion engine.

Reads a CSV "with all details possible":
  * encoding detection (BOM, UTF-16/32, charset-normalizer, safe fallbacks)
  * delimiter sniffing (, ; tab | custom) with quote support
  * ragged-row repair, empty-row removal, header sanitisation
  * per-column type inference: integers, decimals (dot/comma/space grouping),
    currency, percent, booleans, dates (ISO / YMD / DMY / MDY), times
  * leading-zero and big-integer preservation as text

Writes Excel "with all details possible":
  * per-column number formats (#,##0.00, currency, 0.00%, dates, times)
  * header styling, freeze panes, autofilter, auto column widths, borders,
    zebra striping, wrap text
  * 1,048,576-row sheet splitting, 32,767-char cell truncation,
    sheet-name sanitisation, formula-injection escaping
  * optional "Details" summary sheet with per-column statistics
  * optional "monitoring profile": key=value "tags" expanded into columns,
    UNIX-nanosecond "time" -> Date/Timestamp, standard monitoring column
    order, descriptive empty-value labels, colored status cells
"""

from __future__ import annotations

import codecs
import csv
import io
import re
import unicodedata
from collections import Counter
from dataclasses import dataclass, field, fields as dataclass_fields
from datetime import date, datetime, time, timezone
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

from charset_normalizer import from_bytes
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

MAX_ROWS_TOTAL = 1_048_576          # hard Excel limit per sheet
MAX_CELL_CHARS = 32_767             # hard Excel limit per cell
SAMPLE_SIZE = 5_000                 # rows sampled for type inference
MIN_TYPE_RATIO = 0.90               # ratio of values that must match a type
UNIQUE_CAP = 20_000                 # cap for the unique-value counter
MAX_INT = 10**15                    # Excel keeps 15 significant digits

ENCODINGS = [
    "utf-8", "utf-8-sig", "utf-16", "utf-32", "cp1252", "iso-8859-1",
    "iso-8859-15", "cp850", "cp437", "gb18030", "gbk", "big5", "shift_jis",
    "euc-jp", "euc-kr", "koi8-r", "windows-1251", "windows-1254", "mac-roman",
]

_DELIMITERS = [",", ";", "\t", "|"]

_BOMS: Tuple[Tuple[bytes, str], ...] = (
    (codecs.BOM_UTF32_LE, "utf-32-le"),
    (codecs.BOM_UTF32_BE, "utf-32-be"),
    (codecs.BOM_UTF8, "utf-8-sig"),
    (codecs.BOM_UTF16_LE, "utf-16-le"),
    (codecs.BOM_UTF16_BE, "utf-16-be"),
)


class ConversionError(Exception):
    """A user-facing conversion problem."""


# ==========================================================================
# Options
# ==========================================================================

@dataclass
class Options:
    # --- parsing ---
    encoding: str = "auto"            # "auto" or a codec name from ENCODINGS
    delimiter: str = "auto"           # "auto" | "," | ";" | "\\t" | "|" | "custom"
    custom_delimiter: str = ","
    quotechar: str = '"'
    has_header: bool = True
    trim_values: bool = False         # strip surrounding whitespace off values
    drop_empty_rows: bool = True
    max_rows: int = 0                 # 0 = unlimited

    # --- type inference ---
    decimal_separator: str = "auto"   # "auto" | "dot" | "comma"
    date_order: str = "auto"          # "auto" | "YMD" | "DMY" | "MDY"
    detect_numbers: bool = True
    detect_dates: bool = True
    detect_currency: bool = True
    detect_percent: bool = True
    detect_bool: bool = True
    preserve_leading_zeros: bool = True

    # --- safety ---
    escape_formulas: bool = True

    # --- formatting ---
    style_header: bool = True
    header_fill: str = "1F4E78"
    header_font: str = "FFFFFF"
    freeze_panes: bool = True
    auto_filter: bool = True
    auto_width: bool = True
    max_width: int = 60
    borders: bool = False
    zebra: bool = False
    wrap_text: bool = False

    # --- output ---
    single_workbook: bool = True      # all CSVs become sheets of one workbook
    include_summary: bool = True      # trailing "Details" sheet

    # --- monitoring profile (key=value tags -> columns) ---
    monitoring_profile: bool = False  # tags -> columns, ns "time" -> Date/Timestamp,
                                      # standard order, labels, status colors
    zebra_fill: str = "F2F2F2"        # zebra stripe color (F9FAFB in the profile)

    def __post_init__(self) -> None:
        self.zebra_fill = _clean_color(self.zebra_fill, "F2F2F2").upper()
        if self.monitoring_profile:
            # The profile mandates its Excel look (midnight-blue header, zebra,
            # filters, freeze, widths) regardless of the individual switches.
            self.style_header = True
            self.header_fill = "1F4E78"
            self.header_font = "FFFFFF"
            self.freeze_panes = True
            self.auto_filter = True
            self.auto_width = True
            self.zebra = True
            self.zebra_fill = "F9FAFB"

    @classmethod
    def from_dict(cls, data: Any) -> "Options":
        if not isinstance(data, dict):
            return cls()
        kwargs: Dict[str, Any] = {}
        for f in dataclass_fields(cls):
            if f.name not in data:
                continue
            value = data[f.name]
            default = f.default
            if isinstance(default, bool):
                if isinstance(value, bool):
                    kwargs[f.name] = value
                else:
                    kwargs[f.name] = str(value).strip().lower() in ("1", "true", "yes", "on")
            elif isinstance(default, int):
                try:
                    kwargs[f.name] = int(value)
                except (TypeError, ValueError):
                    pass
            elif isinstance(value, str):
                kwargs[f.name] = value
        opts = cls(**kwargs)
        if opts.delimiter == "custom" and not opts.custom_delimiter:
            opts.custom_delimiter = ","
        if opts.delimiter != "auto" and opts.delimiter != "custom" and opts.delimiter not in _DELIMITERS:
            opts.delimiter = "auto"
        if opts.date_order not in ("auto", "YMD", "DMY", "MDY"):
            opts.date_order = "auto"
        if opts.decimal_separator not in ("auto", "dot", "comma"):
            opts.decimal_separator = "auto"
        opts.max_width = min(max(opts.max_width, 8), 255)
        opts.max_rows = max(int(opts.max_rows or 0), 0)
        return opts


# ==========================================================================
# Encoding / delimiter detection
# ==========================================================================

def _charset_detect(raw: bytes) -> Tuple[Optional[str], float]:
    try:
        best = from_bytes(raw[:2_000_000]).best()
    except Exception:
        return None, 0.0
    if best is None or not best.encoding:
        return None, 0.0
    enc = str(best.encoding).replace("_", "-").lower()
    if enc in ("utf8", "utf-8"):
        enc = "utf-8"
    conf = max(0.0, min(1.0, float(getattr(best, "coherence", 0.0) or 0.0) / 100.0))
    return enc, conf


def decode_text(raw: bytes, forced: str = "auto") -> Tuple[str, str, float, List[str]]:
    """Return (text, encoding_used, confidence, warnings). Always succeeds."""
    warnings: List[str] = []

    if forced and forced != "auto":
        try:
            return raw.decode(forced), forced, 1.0, warnings
        except LookupError:
            warnings.append(f"Unknown encoding '{forced}' - falling back to auto-detection.")
        except UnicodeDecodeError as exc:
            warnings.append(f"'{forced}' failed to decode the file ({exc}) - auto-detecting.")

    for bom, enc in _BOMS:
        if raw.startswith(bom):
            return raw.decode(enc, errors="replace"), enc, 1.0, warnings

    # A NUL byte early on usually means UTF-16/32 without a BOM.
    charset_guess: Optional[str] = None
    charset_conf = 0.0
    if b"\x00" in raw[:4096]:
        charset_guess, charset_conf = _charset_detect(raw)
        if charset_guess:
            try:
                return raw.decode(charset_guess), charset_guess, charset_conf, warnings
            except (UnicodeDecodeError, LookupError):
                warnings.append(f"'{charset_guess}' failed to decode the file - trying other encodings.")

    # Valid UTF-8 that is not BOM-tagged is very unlikely to be anything else.
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        pass
    else:
        label = "utf-8-sig" if raw.startswith(codecs.BOM_UTF8) else "utf-8"
        return text, label, 1.0, warnings

    if charset_guess is None:
        charset_guess, charset_conf = _charset_detect(raw)

    for enc in (charset_guess, "cp1252", "iso-8859-1", "latin-1"):
        if not enc:
            continue
        try:
            text = raw.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
        if enc != charset_guess:
            warnings.append(f"Decoded as '{enc}' after other encodings failed.")
        return text, enc, (charset_conf if enc == charset_guess else 0.0), warnings

    return raw.decode("latin-1", errors="replace"), "latin-1", 0.0, warnings


def detect_delimiter(text: str, forced: str = "auto", custom: str = ",") -> str:
    if forced == "custom":
        return (custom or ",")[:1]
    if forced in _DELIMITERS:
        return forced

    sample = text[:100_000]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
        if dialect.delimiter in _DELIMITERS:
            return dialect.delimiter
    except (csv.Error, TypeError, AttributeError):
        pass

    lines = [line for line in sample.splitlines()[:50] if line.strip()]
    scores: Dict[str, int] = {}
    for delim in _DELIMITERS:
        positive = [c for c in (line.count(delim) for line in lines) if c > 0]
        if not positive:
            scores[delim] = 0
            continue
        common, occurrences = Counter(positive).most_common(1)[0]
        scores[delim] = common * occurrences if occurrences >= max(1, len(lines) // 2) else 0
    best = max(_DELIMITERS, key=lambda d: scores[d])
    return best if scores[best] > 0 else ","


# ==========================================================================
# Row reading / sanitising
# ==========================================================================

def read_rows(text: str, delimiter: str, quotechar: str) -> List[List[str]]:
    if quotechar not in ('"', "'"):
        quotechar = '"'
    reader = csv.reader(io.StringIO(text, newline=""), delimiter=delimiter, quotechar=quotechar)
    rows = [[cell if cell is not None else "" for cell in row] for row in reader]
    return _unwrap_wrapped_rows(rows, delimiter, quotechar)


def _unwrap_wrapped_rows(rows: List[List[str]], delimiter: str, quotechar: str) -> List[List[str]]:
    """Repair files whose lines are each one quoted field with doubled quotes.

    A line like ``"a,""b"",1"`` is a single CSV cell whose content is itself
    ``a,"b",1``; such rows collapse to one cell when read. When most data rows
    are single cells that re-parse into the header's column count, the
    re-parsed cells are used instead (well-formed files are left untouched).
    """
    if not rows:
        return rows

    def reparsed(cell: str) -> List[str]:
        try:
            return next(csv.reader(io.StringIO(cell, newline=""),
                                   delimiter=delimiter, quotechar=quotechar))
        except (csv.Error, StopIteration):
            return []

    header = rows[0]
    if len(header) == 1:                       # the header line may be wrapped too
        candidate = reparsed(header[0])
        if len(candidate) > 1:
            header = candidate
            rows = [header] + rows[1:]
    width = len(header)
    if width <= 1:
        return rows

    body = rows[1:]
    single = sum(1 for row in body if len(row) == 1)
    if not body or single * 2 <= len(body):
        return rows                            # not a wrapped file

    fixed: List[List[str]] = [header]
    changed = 0
    for row in body:
        if len(row) == 1 and row[0]:
            cells = reparsed(row[0])
            if len(cells) == width:
                fixed.append(cells)
                changed += 1
                continue
        fixed.append(row)
    return fixed if changed else rows


def sanitize_header(cells: Sequence[str]) -> List[str]:
    used = set()
    out: List[str] = []
    for i, raw in enumerate(cells):
        name = (raw or "").strip().replace("\n", " ").replace("\r", " ")[:500]
        if not name:
            name = f"Column {i + 1}"
        candidate, k = name, 2
        while candidate.lower() in used:
            candidate = f"{name} ({k})"
            k += 1
        used.add(candidate.lower())
        out.append(candidate)
    return out


def safe_sheet_name(base: str, used: set) -> str:
    title = re.sub(r"[\\/*?:\[\]]", "_", (base or "").strip()) or "Sheet"
    title = title.strip("'.").strip() or "Sheet"
    title = title[:31].rstrip() or "Sheet"
    candidate, k = title, 2
    while candidate.lower() in used:
        suffix = f" ({k})"
        candidate = title[: 31 - len(suffix)].rstrip() + suffix
        k += 1
    used.add(candidate.lower())
    return candidate


def file_stem(filename: str) -> str:
    name = (filename or "data").replace("\\", "/").split("/")[-1]
    return name.rsplit(".", 1)[0] if "." in name else name


def display_width(text: str, cap: int = 400) -> int:
    """Column width in Excel units (CJK/wide characters count as 2)."""
    width = 0
    for ch in text[:cap]:
        width += 2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1
    return width


def _clean_color(value: str, fallback: str) -> str:
    value = (value or "").strip().lstrip("#")
    return value if re.fullmatch(r"[0-9A-Fa-f]{6}", value) else fallback


def _argb(color: str) -> str:
    """RRGGBB -> opaque AARRGGBB, the form Excel writes for fills/fonts."""
    color = (color or "").upper()
    return color if len(color) == 8 else "FF" + color


# ==========================================================================
# Numbers
# ==========================================================================

# style key -> (decimal separator, group separator or None)
NUM_STYLES: Dict[str, Tuple[str, Optional[str]]] = {
    "us": (".", ","),
    "eu": (",", "."),
    "space": (".", " "),
    "space_eu": (",", " "),
}
_AUTO_ORDER = {
    "auto": ["us", "eu", "space", "space_eu"],
    "dot": ["us", "space"],
    "comma": ["eu", "space_eu"],
}

_NUM_CHAR = re.compile(r"[+-]?[\d.,\u00a0\u202f ]")
_SPACE_GROUP = re.compile(r"^\d{1,3}(?:[ \u00a0\u202f]\d{3})+(?:[.,]\d+)?$")
_EXPONENT = re.compile(r"^\d+(?:[.,]\d+)?[eE][+-]?\d+$")


@dataclass
class NumResult:
    value: float
    decimals: int
    grouped: bool


def parse_styled(text: str, dec: str, group: Optional[str]) -> Optional[NumResult]:
    """Parse `text` assuming a fixed (decimal, group) separator style."""
    t = text.strip().replace("\u00a0", " ").replace("\u202f", " ")
    if not t or not _NUM_CHAR.match(t):
        return None
    sign = 1
    if t[0] in "+-":
        if t[0] == "-":
            sign = -1
        t = t[1:].strip()
    if not t or not re.fullmatch(r"[\d., ]+", t):
        return None

    if _EXPONENT.match(t.replace(" ", "")):
        norm = t.replace(" ", "")
        if group:
            norm = norm.replace(group, "")
        if dec != ".":
            norm = norm.replace(dec, ".")
        try:
            value = float(norm)
        except ValueError:
            return None
        return NumResult(sign * value, 0, False)

    # Space grouping ("1 234.56" / "1 234,56") is validated and removed here.
    if _SPACE_GROUP.match(t):
        t = t.replace(" ", "")

    if t.count(dec) > 1:
        return None
    has_dec = dec in t
    grouped = False

    if group and group in t:
        int_part, _, frac = t.partition(dec) if has_dec else (t, "", "")
        groups = int_part.split(group)
        if not re.fullmatch(r"\d{1,3}", groups[0] or ""):
            return None
        if any(not re.fullmatch(r"\d{3}", g) for g in groups[1:]):
            return None
        if has_dec and not re.fullmatch(r"\d+", frac or ""):
            return None
        intdigits = "".join(groups)
        grouped = True
    elif has_dec:
        int_part, _, frac = t.partition(dec)
        if not re.fullmatch(r"\d*", int_part) or not re.fullmatch(r"\d+", frac):
            return None
        intdigits = int_part
    else:
        if not re.fullmatch(r"\d+", t):
            return None
        intdigits, frac = t, ""

    decimals = len(frac) if frac else 0
    try:
        value = float(f"{intdigits}.{frac}") if frac else int(intdigits)
    except ValueError:
        return None
    return NumResult(sign * value, decimals, grouped)


def resolve_num_style(values: Sequence[str], decimal_option: str) -> str:
    order = _AUTO_ORDER.get(decimal_option, _AUTO_ORDER["auto"])
    best_style, best_count = order[0], -1
    for style in order:
        dec, group = NUM_STYLES[style]
        count = sum(1 for v in values if parse_styled(v, dec, group) is not None)
        if count > best_count:
            best_style, best_count = style, count
        if count == len(values):
            break
    return best_style


def _ratio(predicate, values: Sequence[str]) -> float:
    if not values:
        return 0.0
    return sum(1 for v in values if predicate(v)) / len(values)


# ==========================================================================
# Currency / percent / boolean
# ==========================================================================

_CURRENCY_SYMBOLS = "$€£¥₹₩₽₪₴₫₺₦฿₲₵₡₮₸₾"
_CURRENCY_CODES = (
    "USD|EUR|GBP|JPY|CNY|RMB|INR|CAD|AUD|NZD|CHF|SEK|NOK|DKK|BRL|MXN|ARS|CLP|COP|PEN|"
    "ZAR|KRW|SGD|HKD|TWD|THB|IDR|MYR|PHP|VND|PLN|CZK|HUF|RON|BGN|UAH|TRY|ILS|AED|SAR|"
    "QAR|EGP|NGN|KES|GHS|PKR|BDT|LKR|KZT|GEL|AMD|AZN|MKD|RSD|ISK|HRK"
)
_CUR_LEAD = re.compile(r"^(?P<tok>[" + _CURRENCY_SYMBOLS + r"])\s*(?P<num>.+)$")
_CUR_CODE_LEAD = re.compile(r"^(?P<tok>(?:" + _CURRENCY_CODES + r"))\s+(?P<num>.+)$")
_CUR_TRAIL = re.compile(r"^(?P<num>.+?)\s*(?P<tok>[" + _CURRENCY_SYMBOLS + r"])$")
_CUR_CODE_TRAIL = re.compile(r"^(?P<num>.+?)\s+(?P<tok>(?:" + _CURRENCY_CODES + r"))$")
_PERCENT = re.compile(r"^(?P<num>[+-]?[\d.,\u00a0\u202f ]+)%$")
_BOOL_SETS = [
    ({"true"}, {"false"}),
    ({"yes"}, {"no"}),
    ({"on"}, {"off"}),
]


def split_currency(text: str) -> Optional[Tuple[str, str]]:
    """Split "€1.234,56" / "1,234.56 USD" into (number, symbol-or-code)."""
    stripped = text.strip()
    for pattern in (_CUR_LEAD, _CUR_CODE_LEAD, _CUR_TRAIL, _CUR_CODE_TRAIL):
        m = pattern.match(stripped)
        if m:
            return m.group("num").strip(), m.group("tok")
    return None


# ==========================================================================
# Dates / times
# ==========================================================================

_SLASH_DATE = re.compile(r"^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})(?:[ T]|$)")
_YMD_DATE = re.compile(r"^\d{4}[/.-]\d{1,2}[/.-]\d{1,2}(?:[ T]|$)")
_TIME_ONLY = re.compile(r"^\d{1,2}:\d{2}(?::\d{2})?(?:[.,]\d+)?\s*(?:[AaPp][Mm])?$")
# Cheap gate: only columns whose values *look* temporal are worth strptime attempts.
_LOOKS_TEMPORAL = re.compile(
    r"^(?:\d{4}[-/.T]\d{1,2}[-/.T]\d{1,2}"
    r"|\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}"
    r"|\d{8}(?!\d)"
    r"|\d{1,2}:\d{2}"
    r"|[A-Za-z]{3,9}\.?\s+\d{1,2},?"
    r"|\d{1,2}\s+[A-Za-z]{3,9})"
)


def _looks_temporal(values: Sequence[str]) -> bool:
    if not values:
        return False
    hits = sum(1 for v in values if _LOOKS_TEMPORAL.match(v))
    return hits / len(values) >= MIN_TYPE_RATIO

_DATE_FMTS = {
    "YMD": ["%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y%m%d"],
    "DMY": ["%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y", "%d-%m-%y",
            "%d %b %Y", "%d %B %Y", "%d.%b.%Y"],
    "MDY": ["%m/%d/%Y", "%m-%d-%Y", "%m.%d.%Y", "%m/%d/%y", "%b %d, %Y",
            "%B %d, %Y", "%b %d %Y", "%B %d %Y", "%b %d, %y"],
}
_TIME_SUFFIXES = [" %H:%M:%S", " %H:%M", " %I:%M:%S %p", " %I:%M %p",
                  "T%H:%M:%S", "T%H:%M", ""]
_TIME_FMTS = ["%H:%M:%S", "%H:%M", "%I:%M:%S %p", "%I:%M %p", "%H:%M:%S.%f"]


def detect_date_order(sample: Sequence[str]) -> str:
    first_max = second_max = 0
    ymd_seen = False
    for value in sample:
        m = _SLASH_DATE.match(value)
        if m:
            first_max = max(first_max, int(m.group(1)))
            second_max = max(second_max, int(m.group(2)))
        elif _YMD_DATE.match(value):
            ymd_seen = True
    if first_max > 12:
        return "DMY"
    if second_max > 12:
        return "MDY"
    if ymd_seen:
        return "YMD"
    return "MDY"  # ambiguous: match the en-US default, the user can override


def _candidate_specs(order: str) -> List[str]:
    if order not in ("YMD", "DMY", "MDY"):
        order = detect_date_order([])
    specs: List[str] = ["ISO"]
    for dfmt in list(_DATE_FMTS["YMD"]) + list(_DATE_FMTS[order]):
        for tsuf in _TIME_SUFFIXES:
            specs.append(dfmt + tsuf)
    seen, unique = set(), []
    for spec in specs:
        if spec not in seen:
            seen.add(spec)
            unique.append(spec)
    return unique


def parse_dt(text: str, spec: str) -> Optional[Any]:
    if spec == "ISO":
        try:
            return datetime.fromisoformat(text.replace("Z", "+00:00"))
        except ValueError:
            pass
        try:
            return date.fromisoformat(text)
        except ValueError:
            return None
    try:
        return datetime.strptime(text, spec)
    except ValueError:
        return None


def spec_is_datetime(spec: str) -> bool:
    return spec == "ISO" or "%H" in spec or "%I" in spec


def _naive(value: Any) -> Any:
    """Excel has no timezone concept: keep the wall clock, drop the offset."""
    if isinstance(value, datetime) and value.tzinfo is not None:
        return value.replace(tzinfo=None)
    return value


# ==========================================================================
# Monitoring profile: key=value tags, ns timestamps, standard layout
# ==========================================================================

MONITORING_SHEET = "Detailed Raw Data"

# Source name (column name or tag key, case-insensitive) -> French report header.
# Both the input CSV columns and the common tag keys are mapped.
MONITORING_RENAMES = {
    # --- input CSV columns ---
    "time": "Horodatage (UTC)",
    "timestamp": "Horodatage (UTC)",
    "count_duration_seconds": "Durée (sec)",
    "duration_seconds": "Durée (sec)",
    "duration (sec)": "Durée (sec)",
    "duration": "Durée (sec)",
    "name": "Type Enregistrement",
    "record type": "Type Enregistrement",
    # --- tag keys (and generic English synonyms) ---
    "zone": "Zone Géographique",
    "geographic_zone": "Zone Géographique",
    "geographic zone": "Zone Géographique",
    "status": "Statut Global",
    "overall_status": "Statut Global",
    "overall status": "Statut Global",
    "global_status": "Statut Global",
    "global status": "Statut Global",
    "monitoring_name": "Nom du Service (Monitoring)",
    "service_name": "Nom du Service (Monitoring)",
    "service name": "Nom du Service (Monitoring)",
    "test_name": "Nom du Test",
    "test name": "Nom du Test",
    "failure_reason": "Raison de l'Échec",
    "failure reason": "Raison de l'Échec",
    "cellular_ussd_status": "Statut USSD",
    "ussd_status": "Statut USSD",
    "ussd status": "Statut USSD",
    "cellular_sms_reception_status": "Statut SMS",
    "sms_status": "Statut SMS",
    "sms status": "Statut SMS",
    "failed_action_kind": "Type d'Action en Échec",
    "failed_action_type": "Type d'Action en Échec",
    "failed action type": "Type d'Action en Échec",
    "monitoring_id": "ID Monitoring",
    "monitoring id": "ID Monitoring",
    "test_id": "ID Test",
    "test id": "ID Test",
    "product_name": "Nom du Produit",
    "product name": "Nom du Produit",
    "product_id": "ID Produit",
    "product id": "ID Produit",
    "platform_id": "ID Plateforme",
    "platform id": "ID Plateforme",
    "type": "Type",
}

MONITORING_COLUMN_ORDER = [
    "Date", "Horodatage (UTC)", "Zone Géographique", "Statut Global",
    "Nom du Service (Monitoring)", "Nom du Test", "Raison de l'Échec",
    "Durée (sec)", "Statut USSD", "Statut SMS", "Type d'Action en Échec",
    "ID Monitoring", "ID Test", "Nom du Produit", "ID Produit",
    "ID Plateforme", "Type", "Type Enregistrement",
]
MONITORING_LABELS = {                      # empty cells get a descriptive label
    "zone géographique": "Non Définie",
    "raison de l'échec": "Aucune (Succès)",
    "statut sms": "N/A",
    "type d'action en échec": "N/A",
}
MONITORING_STATUS_FILLS = {                # Statut Global cell colors
    "success": "E2EFDA",                   # light green
    "failed": "FCE4D6",                    # light red
    "error": "FFF2CC",                     # light yellow
}
_EPOCH_MAX = 253_402_300_799               # 9999-12-31 23:59:59 UTC
_EPOCH_MIN = 1_000_000_000                 # 2001-09-09: guards tiny integers


def _unescape(text: str) -> str:
    """Remove escape characters: '\\ ' -> ' ', '\\,' -> ',', '\\\\' -> '\\'."""
    return re.sub(r"\\(.)", r"\1", text)


def parse_tags(text: str) -> Dict[str, str]:
    """Parse `key1=value1,key2=value2` into an ordered dict (last key wins).

    Commas preceded by a backslash stay inside the value ("c=x\\,y" -> "x,y").
    Fragments without '=' are ignored; keys/values are unescaped and trimmed.
    """
    pairs: Dict[str, str] = {}
    if not text or not text.strip():
        return pairs
    for chunk in re.split(r"(?<!\\),", text):
        key, sep, value = chunk.partition("=")
        if not sep:
            continue                      # not a key=value fragment
        key = _unescape(key).strip()
        if key:
            pairs[key] = _unescape(value).strip()
    return pairs


def expand_tags_columns(header: Sequence[str], data: Sequence[List[str]]):
    """Split the 'tags' column into one independent column per key.

    Returns (header, data, number_of_keys). The raw 'tags' column is dropped
    (its content now lives in the extracted columns); a key that collides
    with an existing column becomes "<key> (2)", "<key> (3)"...
    """
    idx = next((i for i, h in enumerate(header) if h.strip().lower() == "tags"), None)
    if idx is None:
        return header, data, 0

    key_order: List[str] = []
    parsed_rows: List[Dict[str, str]] = []
    for row in data:
        parsed = parse_tags(row[idx] if idx < len(row) else "")
        parsed_rows.append(parsed)
        for key in parsed:
            if key not in key_order:
                key_order.append(key)

    existing = {h.strip().lower() for h in header}
    mapping: List[Tuple[str, str]] = []         # (tag key, final column name)
    for key in key_order:
        base = key or "tag"
        name, k = base, 2
        while name.strip().lower() in existing:
            name = f"{base} ({k})"
            k += 1
        existing.add(name.strip().lower())
        mapping.append((key, name))

    new_header = list(header[:idx]) + [name for _, name in mapping] + list(header[idx + 1:])
    new_data = [list(row[:idx]) + [parsed.get(key, "") for key, _ in mapping]
                + list(row[idx + 1:]) for row, parsed in zip(data, parsed_rows)]
    return new_header, new_data, len(mapping)


def _epoch_to_datetime(text: str) -> Optional[datetime]:
    """UNIX timestamp (ns/us/ms/s, detected by magnitude) -> naive UTC datetime."""
    t = (text or "").strip()
    if not t or not re.fullmatch(r"\d+(?:\.\d+)?", t):
        return None
    try:
        value = float(t)
    except ValueError:
        return None
    if value >= 1e17:                       # nanoseconds
        value /= 1e9
    elif value >= 1e14:                     # microseconds
        value /= 1e6
    elif value >= 1e11:                     # milliseconds
        value /= 1e3
    if not (_EPOCH_MIN <= value <= _EPOCH_MAX):
        return None                         # not a plausible modern timestamp
    try:
        return datetime.fromtimestamp(value, tz=timezone.utc).replace(tzinfo=None)
    except (OverflowError, OSError, ValueError):
        return None


def convert_unix_time_column(header: Sequence[str], data: Sequence[List[str]]):
    """Convert the 'time' column (UNIX ns) to 'Timestamp' + derived 'Date'.

    Returns (header, data, converted, unparsed). The column keeps its name
    when a 'Timestamp' column already exists (nothing is ever shadowed).
    """
    idx = next((i for i, h in enumerate(header) if h.strip().lower() == "time"), None)
    if idx is None:
        return header, data, 0, 0

    converted = unparsed = 0
    stamps: List[str] = []
    dates: List[str] = []
    for row in data:
        text = row[idx] if idx < len(row) else ""
        dt = _epoch_to_datetime(text)
        if dt is None:
            if text.strip():
                unparsed += 1
            stamps.append(text.strip())
            dates.append("")
        else:
            converted += 1
            stamps.append(dt.strftime("%Y-%m-%d %H:%M:%S"))
            dates.append(dt.strftime("%Y-%m-%d"))

    has_timestamp = any(h.strip().lower() == "timestamp" for i, h in enumerate(header) if i != idx)
    ts_name = "time" if has_timestamp else "Timestamp"
    has_date = any(h.strip().lower() == "date" for h in header)

    new_header = list(header[:idx])
    if not has_date:
        new_header.append("Date")
    new_header.append(ts_name)
    new_header += list(header[idx + 1:])

    new_data = []
    for row, stamp, day in zip(data, stamps, dates):
        prefix = list(row[:idx])
        if not has_date:
            prefix.append(day)
        prefix.append(stamp)
        new_data.append(prefix + list(row[idx + 1:]))
    return new_header, new_data, converted, unparsed


def rename_monitoring_columns(header: Sequence[str]) -> List[str]:
    """Map column names / tag keys to the canonical French report headers.

    Matching is case-insensitive; unknown names are kept as-is. Two sources
    mapping to the same header get a " (2)" suffix so nothing is lost.
    """
    out: List[str] = []
    used: Dict[str, int] = {}
    for raw in header:
        name = MONITORING_RENAMES.get((raw or "").strip().lower(), (raw or "").strip())
        key = name.lower()
        used[key] = used.get(key, 0) + 1
        if used[key] > 1:
            name = f"{name} ({used[key]})"
        out.append(name)
    return out


def reorder_monitoring_columns(header: Sequence[str], data: Sequence[List[str]]):
    """Standard priority order first (Date, Timestamp, ...), extras after."""
    wanted = [w.lower() for w in MONITORING_COLUMN_ORDER]
    lower = [h.strip().lower() for h in header]
    chosen: List[int] = []
    for name in wanted:
        for i, header_name in enumerate(lower):
            if header_name == name and i not in chosen:
                chosen.append(i)
                break
    order = chosen + [i for i in range(len(header)) if i not in chosen]
    if order == list(range(len(header))):
        return list(header), [list(r) for r in data]
    return ([header[i] for i in order],
            [[row[i] for i in order] for row in data])


def fill_monitoring_labels(header: Sequence[str], data: Sequence[List[str]]) -> List[List[str]]:
    """Empty cells get a descriptive label ('Non Définie', 'Aucune (Succès)', 'N/A')."""
    for i, name in enumerate(header):
        label = MONITORING_LABELS.get(name.strip().lower())
        if not label:
            continue
        for row in data:
            if i < len(row) and not row[i].strip():
                row[i] = label
    return [list(r) for r in data]


def monitoring_transform(header: Sequence[str], data: Sequence[List[str]]):
    """Full pipeline: ns time -> tags -> reorder -> labels. Returns warnings."""
    warnings: List[str] = []
    header, data, converted, unparsed = convert_unix_time_column(header, data)
    if converted:
        warnings.append(f"'time': {converted:,} UNIX timestamp(s) converted to "
                        f"date-time (+ derived Date column, UTC).")
    if unparsed:
        warnings.append(f"'time': {unparsed:,} value(s) were not UNIX timestamps "
                        f"and were left unchanged.")
    header, data, key_count = expand_tags_columns(header, data)
    if key_count:
        warnings.append(f"'tags': expanded into {key_count} key=value column(s).")
    header = rename_monitoring_columns(header)
    header, data = reorder_monitoring_columns(header, data)
    data = fill_monitoring_labels(header, data)
    return header, data, warnings


# ==========================================================================
# Column metadata + inference
# ==========================================================================

@dataclass
class ColMeta:
    name: str
    opts: Options
    kind: str = "text"                 # text|int|float|bool|date|datetime|time
    number_format: Optional[str] = None
    spec: Optional[str] = None         # date spec | true-words for bool
    num_style: str = "us"
    is_percent: bool = False
    is_currency: bool = False
    currency_token: str = ""
    width: int = 0
    non_empty: int = 0
    empty: int = 0
    kept_text: int = 0
    truncated: int = 0
    examples: List[str] = field(default_factory=list)
    unique: set = field(default_factory=set, repr=False)
    unique_capped: bool = False
    min_v: Any = None
    max_v: Any = None

    # -- label used in reports --------------------------------------------
    @property
    def label(self) -> str:
        if self.is_percent:
            return "percent"
        if self.is_currency:
            return f"currency {self.currency_token}".strip()
        return {
            "text": "text", "int": "integer", "float": "decimal",
            "bool": "boolean", "date": "date", "datetime": "datetime",
            "time": "time",
        }.get(self.kind, self.kind)

    # -- value conversion --------------------------------------------------
    def convert(self, raw: str) -> Any:
        text = raw.strip()
        if not text:
            self.empty += 1
            return None
        self.non_empty += 1
        if len(self.examples) < 3:
            self.examples.append(text[:200])

        fallback: str = text if self.opts.trim_values else raw
        keep_text = False

        if self.kind == "text":
            value: Any = fallback
        elif self.kind == "bool":
            value = text.lower() in self._true_set()
        elif self.kind in ("date", "datetime", "time"):
            parsed = parse_dt(text, self.spec or "ISO")
            if parsed is None:
                keep_text = True
                value = fallback
            else:
                value = self._coerce_dt(parsed)
        else:  # int / float, optionally percent or currency
            number_text = text
            if self.is_percent:
                if number_text.endswith("%"):
                    number_text = number_text[:-1].strip()
            elif self.is_currency:
                parts = split_currency(number_text)
                if parts:
                    number_text = parts[0]
            dec, group = NUM_STYLES.get(self.num_style, NUM_STYLES["us"])
            num = parse_styled(number_text, dec, group)
            if num is None:
                keep_text = True
                value = fallback
            else:
                raw_val = num.value / 100.0 if self.is_percent else num.value
                if abs(raw_val) >= MAX_INT:
                    keep_text = True      # beyond Excel's 15-digit precision
                    value = fallback
                elif self.kind == "int" and float(raw_val).is_integer():
                    value = int(raw_val)
                else:
                    value = raw_val

        if keep_text:
            self.kept_text += 1
        self._update_stats(value)
        return self._cap(value)

    def _true_set(self) -> set:
        return set((self.spec or "true").split("|"))

    def _coerce_dt(self, parsed: Any) -> Any:
        parsed = _naive(parsed)
        if self.kind == "time":
            return parsed.time() if isinstance(parsed, datetime) else parsed
        if self.kind == "date":
            return parsed.date() if isinstance(parsed, datetime) else parsed
        # datetime
        if isinstance(parsed, datetime):
            return parsed.replace(microsecond=0)
        return datetime.combine(parsed, time.min)

    def _cap(self, value: Any) -> Any:
        if isinstance(value, str) and len(value) > MAX_CELL_CHARS:
            self.truncated += 1
            return value[:MAX_CELL_CHARS]
        return value

    def _update_stats(self, value: Any) -> None:
        if value is None or isinstance(value, bool):
            return
        if isinstance(value, (int, float, datetime, date, time)):
            try:
                self.min_v = value if self.min_v is None else min(self.min_v, value)
                self.max_v = value if self.max_v is None else max(self.max_v, value)
            except TypeError:
                pass
        if not self.unique_capped:
            try:
                self.unique.add(value)
                if len(self.unique) >= UNIQUE_CAP:
                    self.unique_capped = True
            except TypeError:
                pass


def infer_column(name: str, sample: Sequence[str], opts: Options) -> ColMeta:
    """Pick the richest type that at least 90% of the sample can hold."""
    meta = ColMeta(name=name, opts=opts)
    values = [v.strip() for v in sample if v.strip()]
    if not values:
        return meta

    # 1. booleans --------------------------------------------------------
    if opts.detect_bool:
        lowered = {v.lower() for v in values}
        for trues, falses in _BOOL_SETS:
            allowed = trues | falses
            if lowered and lowered <= allowed:
                meta.kind = "bool"
                meta.spec = "|".join(sorted(trues))
                return meta

    # 2. dates / datetimes / times --------------------------------------
    if opts.detect_dates and _looks_temporal(values):
        order = opts.date_order if opts.date_order != "auto" else detect_date_order(values)

        if all(_TIME_ONLY.match(v) for v in values):
            for fmt in _TIME_FMTS:
                if _ratio(lambda v, f=fmt: parse_dt(v, f) is not None, values) >= MIN_TYPE_RATIO:
                    meta.kind = "time"
                    meta.spec = fmt
                    meta.number_format = "hh:mm:ss"
                    return meta

        best_spec, best_count = None, 0
        for spec in _candidate_specs(order):
            count = sum(1 for v in values if parse_dt(v, spec) is not None)
            if count > best_count:
                best_spec, best_count = spec, count
            if count == len(values):
                break
        if best_spec and best_count / len(values) >= MIN_TYPE_RATIO:
            is_dt = spec_is_datetime(best_spec)
            if best_spec == "ISO":
                is_dt = any("T" in v or ":" in v for v in values)
            meta.kind = "datetime" if is_dt else "date"
            meta.spec = best_spec
            meta.number_format = "yyyy-mm-dd hh:mm:ss" if is_dt else "yyyy-mm-dd"
            return meta

    # 3. percent ---------------------------------------------------------
    if opts.detect_percent:
        naked = [m.group("num").strip() if (m := _PERCENT.match(v)) else "" for v in values]
        ok_naked = [n for n in naked if n]
        if len(ok_naked) / len(values) >= MIN_TYPE_RATIO:
            style = resolve_num_style(ok_naked, opts.decimal_separator)
            dec, group = NUM_STYLES[style]
            parsed = [parse_styled(n, dec, group) for n in ok_naked]
            ok = [p for p in parsed if p is not None]
            if len(ok) / len(values) >= MIN_TYPE_RATIO:
                decimals = max((p.decimals for p in ok), default=0)
                meta.kind = "float"
                meta.is_percent = True
                meta.num_style = style
                meta.number_format = f"0.{'0' * max(decimals, 2)}%"
                return meta

    # 4. currency --------------------------------------------------------
    if opts.detect_currency:
        parts = [p for p in (split_currency(v) for v in values) if p]
        if len(parts) / len(values) >= MIN_TYPE_RATIO:
            token = Counter(tok for _, tok in parts).most_common(1)[0][0]
            naked = [num for num, tok in parts if tok == token]
            style = resolve_num_style(naked, opts.decimal_separator)
            dec, group = NUM_STYLES[style]
            parsed = [p for p in (parse_styled(n, dec, group) for n in naked) if p is not None]
            if len(parsed) / len(values) >= MIN_TYPE_RATIO:
                decimals = max((p.decimals for p in parsed), default=0)
                all_int = not any(p.decimals for p in parsed)
                body = "#,##0" if all_int else "#,##0." + "0" * max(decimals, 2)
                sep = "" if len(token) == 1 else " "
                meta.kind = "int" if all_int else "float"
                meta.is_currency = True
                meta.currency_token = token
                meta.num_style = style
                meta.number_format = f'"{token}"{sep}{body}'
                return meta

    # 5. plain numbers ---------------------------------------------------
    if opts.detect_numbers:
        style = resolve_num_style(values, opts.decimal_separator)
        dec, group = NUM_STYLES[style]
        parsed = [parse_styled(v, dec, group) for v in values]
        ok = [p for p in parsed if p is not None]
        if ok and len(ok) / len(values) >= MIN_TYPE_RATIO:
            looks_padded = any(re.fullmatch(r"[+-]?0\d+", v) for v in values)
            if opts.preserve_leading_zeros and looks_padded:
                return meta  # keep "007", "0123" ... exactly as written
            has_decimal = any(p.decimals for p in ok)
            grouped = any(p.grouped for p in ok)
            meta.num_style = style
            if has_decimal:
                meta.kind = "float"
                decimals = max(p.decimals for p in ok)
                meta.number_format = ("#,##0." if grouped else "0.") + "0" * decimals
            else:
                meta.kind = "int"
                if grouped or any(abs(p.value) >= 1000 for p in ok):
                    meta.number_format = "#,##0"
            return meta

    return meta


# ==========================================================================
# Workbook writing
# ==========================================================================

_THIN = Side(style="thin", color="D9D9D9")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_WRAP = Alignment(wrap_text=True, vertical="top")

# Monitoring profile fonts (the reference report uses Segoe UI, bold statuses)
_MONITOR_BODY_FONT = Font(name="Segoe UI", size=9)
_MONITOR_STATUS_FONT = Font(name="Segoe UI", size=9, bold=True)

_ZEBRA_FILLS: Dict[str, PatternFill] = {}


def _zebra_fill(color: str) -> PatternFill:
    """Cached zebra fill for the configured stripe color."""
    key = (color or "F2F2F2").upper()
    fill = _ZEBRA_FILLS.get(key)
    if fill is None:
        fill = PatternFill("solid", fgColor=_argb(key))
        _ZEBRA_FILLS[key] = fill
    return fill


def _write_header(ws, header: Sequence[str], opts: Options) -> None:
    ws.append(list(header))
    fill = PatternFill("solid", fgColor=_argb(_clean_color(opts.header_fill, "1F4E78"))) if opts.style_header else None
    color = _argb(_clean_color(opts.header_font, "FFFFFF")) if opts.style_header else "FF000000"
    if opts.monitoring_profile:
        font = Font(name="Segoe UI", size=11, bold=True, color=color)
    else:
        font = Font(bold=True, color=color)
    for cell in ws[1]:
        cell.font = font
        if fill is not None:
            cell.fill = fill
        cell.alignment = Alignment(vertical="center", wrap_text=opts.wrap_text)
        if opts.borders:
            cell.border = _BORDER
    ws.row_dimensions[1].height = 20


def _finish_sheet(ws, header: Sequence[str], metas: Sequence[ColMeta], opts: Options,
                  last_row: int) -> None:
    if opts.has_header:
        if opts.freeze_panes:
            ws.freeze_panes = "A2"
        if opts.auto_filter and metas:
            ws.auto_filter.ref = f"A1:{get_column_letter(len(metas))}{max(last_row, 1)}"
    if opts.auto_width:
        for i, meta in enumerate(metas):
            width = max(meta.width, display_width(header[i]) if i < len(header) else 0)
            ws.column_dimensions[get_column_letter(i + 1)].width = min(max(width + 2, 8), opts.max_width)


def write_table(
    wb: Workbook,
    base_title: str,
    header: Sequence[str],
    data_rows: Iterable[List[str]],
    metas: Sequence[ColMeta],
    opts: Options,
    used_titles: set,
) -> List[str]:
    """Write the table, splitting into continuation sheets at Excel's row limit."""
    titles: List[str] = []
    esc_cols = [i for i, m in enumerate(metas) if m.kind == "text"]
    any_fmt = any(m.number_format for m in metas)
    status_idx: Optional[int] = None
    status_fills: Dict[str, PatternFill] = {}
    if opts.monitoring_profile:
        for i, name in enumerate(header):
            if name.strip().lower() in ("statut global", "overall status", "global status"):
                status_idx = i
                status_fills = {value: PatternFill("solid", fgColor=_argb(color))
                                for value, color in MONITORING_STATUS_FILLS.items()}
                break
    cell_pass = (bool(esc_cols) or any_fmt or opts.borders or opts.zebra
                 or opts.wrap_text or status_idx is not None)

    budget = MAX_ROWS_TOTAL - (1 if opts.has_header else 0)
    split_index = 0

    def open_sheet() -> Any:
        nonlocal split_index
        split_index += 1
        suffix = "" if split_index == 1 else f" ({split_index})"
        title = safe_sheet_name(f"{base_title}{suffix}", used_titles)
        titles.append(title)
        sheet = wb.create_sheet(title)
        if opts.has_header:
            _write_header(sheet, header, opts)
        return sheet

    ws = open_sheet()
    written = 0
    data_index = 0
    row_base = 1 if opts.has_header else 0  # NOTE: never use ws.max_row per row,
    # openpyxl computes it with max(self._cells) -> O(all cells), quadratic per row.

    for row in data_rows:
        if written >= budget:
            _finish_sheet(ws, header, metas, opts, row_base + written)
            ws = open_sheet()
            written = 0

        converted = [meta.convert(cell) for meta, cell in zip(metas, row)]

        fix_eq: set = set()
        if opts.escape_formulas and esc_cols:
            for i in esc_cols:
                v = converted[i]
                if isinstance(v, str) and v[:1] == "=":
                    fix_eq.add(i)

        r = row_base + written + 1
        for c, value in enumerate(converted, start=1):
            # NOTE: ws.cell() is a plain dict lookup; ws.append() + ws[row] must be
            # avoided here because openpyxl resolves ws[row] via max_column, which
            # scans every cell in the sheet (O(n) per row -> O(n^2) overall).
            cell = ws.cell(row=r, column=c, value=value)
            if cell_pass:
                fmt = metas[c - 1].number_format
                if fmt:
                    cell.number_format = fmt
                if opts.wrap_text:
                    cell.alignment = _WRAP
                if opts.borders:
                    cell.border = _BORDER
                if opts.zebra and data_index % 2 == 1:
                    cell.fill = _zebra_fill(opts.zebra_fill)
                if opts.monitoring_profile:
                    cell.font = _MONITOR_BODY_FONT
                if status_idx is not None and c - 1 == status_idx and isinstance(value, str):
                    status_fill = status_fills.get(value.strip().lower())
                    if status_fill is not None:
                        cell.fill = status_fill
                        cell.font = _MONITOR_STATUS_FONT
            if c - 1 in fix_eq:
                cell.data_type = "s"

        written += 1
        data_index += 1

    _finish_sheet(ws, header, metas, opts, row_base + written)
    return titles


# ==========================================================================
# Summary ("Details") sheet
# ==========================================================================

def _fmt_stat(meta: ColMeta, value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, float) and meta.is_percent:
        return f"{value * 100:g}%"
    if isinstance(value, (int, float)) and meta.is_currency:
        return f"{meta.currency_token} {value:,.2f}"
    if isinstance(value, float):
        return f"{value:g}"
    if isinstance(value, datetime):
        return value.isoformat(sep=" ")
    if isinstance(value, (date, time)):
        return value.isoformat()
    return str(value)


def build_summary(wb: Workbook, reports: Sequence[Dict[str, Any]], metas_per_sheet) -> None:
    ws = wb.create_sheet(safe_sheet_name("Details", {s.title.lower() for s in wb.worksheets}))
    row = 1
    ws.cell(row, 1, "CSV -> Excel conversion details").font = Font(bold=True, size=14)
    row += 2

    for rep in reports:
        info = [
            ("File", rep["file"]),
            ("Sheets", ", ".join(rep["sheets"])),
            ("Rows", rep["rows"]),
            ("Columns", rep["columns"]),
            ("Encoding", f"{rep['encoding']} (confidence {rep['encoding_confidence']:.0%})"),
            ("Delimiter", repr(rep["delimiter"])),
            ("Generated", rep["generated"]),
        ]
        if rep.get("warnings"):
            info.append(("Warnings", " | ".join(rep["warnings"])))
        for key, value in info:
            ws.cell(row, 1, key).font = Font(bold=True)
            ws.cell(row, 2, value)
            row += 1
        row += 1

    headers = ["Sheet", "#", "Column", "Type", "Number format", "Non-empty", "Empty",
               "Unique", "Min", "Max", "Examples"]
    for col, title in enumerate(headers, start=1):
        cell = ws.cell(row, col, title)
        cell.font = Font(bold=True, color="FFFFFFFF")
        cell.fill = PatternFill("solid", fgColor="FF1F4E78")
    header_row = row
    row += 1

    for sheet_title, metas in metas_per_sheet:
        for i, meta in enumerate(metas, start=1):
            values = [
                sheet_title, i, meta.name, meta.label,
                meta.number_format or "General",
                meta.non_empty, meta.empty,
                f"{len(meta.unique)}{'+' if meta.unique_capped else ''}",
                _fmt_stat(meta, meta.min_v), _fmt_stat(meta, meta.max_v),
                ", ".join(meta.examples),
            ]
            for col, value in enumerate(values, start=1):
                ws.cell(row, col, value)
            row += 1

    widths = {"A": 24, "B": 5, "C": 30, "D": 16, "E": 18, "F": 11, "G": 8,
              "H": 10, "I": 20, "J": 20, "K": 60}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    ws.freeze_panes = f"A{header_row + 1}"


# ==========================================================================
# Public API
# ==========================================================================

def _prepare(filename: str, raw: bytes, opts: Options) -> Dict[str, Any]:
    if not raw:
        raise ConversionError(f"'{filename}' is empty.")

    text, encoding, confidence, warnings = decode_text(raw, opts.encoding)
    if "\x00" in text:
        warnings.append("File contains NUL bytes; it may be a binary file, not a CSV.")

    delimiter = detect_delimiter(text, opts.delimiter, opts.custom_delimiter)
    rows = read_rows(text, delimiter, opts.quotechar)

    if opts.drop_empty_rows:
        kept = [r for r in rows if any(cell.strip() for cell in r)]
        if len(kept) != len(rows):
            warnings.append(f"Skipped {len(rows) - len(kept)} empty row(s).")
        rows = kept
    if not rows:
        raise ConversionError(f"'{filename}' contains no data rows.")

    width = max(len(r) for r in rows)
    if any(len(r) != width for r in rows):
        warnings.append(f"Ragged rows padded to {width} columns.")
    for r in rows:
        if len(r) < width:
            r.extend([""] * (width - len(r)))

    if opts.has_header:
        header = sanitize_header(rows[0])
        data = rows[1:]
    else:
        header = [f"Column {i + 1}" for i in range(width)]
        data = rows

    if opts.max_rows and len(data) > opts.max_rows:
        data = data[: opts.max_rows]
        warnings.append(f"Truncated to the first {opts.max_rows:,} data rows (per your limit).")
    if not data:
        raise ConversionError(f"'{filename}' has a header but no data rows.")

    if opts.monitoring_profile:
        header, data, monitor_warns = monitoring_transform(header, data)
        warnings.extend(monitor_warns)

    return {
        "encoding": encoding, "confidence": confidence, "delimiter": delimiter,
        "header": header, "data": data, "warnings": warnings,
    }


def _measure_widths(header: Sequence[str], data: Sequence[List[str]], metas: List[ColMeta],
                    max_width: int) -> None:
    widths = [display_width(h) for h in header]
    for row in data:
        for i, cell in enumerate(row):
            if cell and len(cell) > widths[i]:
                widths[i] = display_width(cell)
                if widths[i] >= max_width:
                    widths[i] = max_width + 1  # stop measuring this column
                    break
    for i, meta in enumerate(metas):
        if i < len(widths):
            meta.width = min(widths[i], max_width + 1)


def _prepare_table(filename: str, raw: bytes, opts: Options) -> Dict[str, Any]:
    """Parse one CSV and infer its column types (no Excel writing yet)."""
    prep = _prepare(filename, raw, opts)
    header, data = prep["header"], prep["data"]

    metas: List[ColMeta] = []
    for i in range(len(header)):
        column_sample = [row[i] for row in data[:SAMPLE_SIZE]]
        metas.append(infer_column(header[i], column_sample, opts))
    _measure_widths(header, data, metas, opts.max_width)

    return {
        "filename": filename,
        "base_title": (MONITORING_SHEET if opts.monitoring_profile
                       else file_stem(filename) or "Sheet1"),
        "header": header,
        "data": data,
        "metas": metas,
        "prep": prep,
    }


def _save_workbook(tables: Sequence[Dict[str, Any]], opts: Options) -> Tuple[bytes, List[Dict[str, Any]]]:
    """Write one workbook containing every prepared table + optional summary."""
    wb = Workbook()
    wb.remove(wb.active)
    wb.properties.title = "CSV to Excel conversion"
    wb.properties.creator = "CSV -> Excel Converter"
    wb.properties.created = datetime.now()

    used: set = set()
    reports: List[Dict[str, Any]] = []
    metas_per_sheet: List[Tuple[str, List[ColMeta]]] = []

    for table in tables:
        prep = table["prep"]
        metas = table["metas"]
        data = table["data"]
        titles = write_table(wb, table["base_title"], table["header"], data, metas, opts, used)

        warnings = list(prep["warnings"])
        if len(titles) > 1:
            warnings.append(
                f"{len(data):,} rows exceed one sheet; data split across {len(titles)} sheets."
            )
        # kept_text / truncated are only known once the rows have been written
        for meta in metas:
            if meta.kept_text:
                warnings.append(
                    f"'{meta.name}': {meta.kept_text:,} value(s) did not fit the inferred "
                    f"type and were kept as text."
                )
            if meta.truncated:
                warnings.append(
                    f"'{meta.name}': {meta.truncated:,} cell(s) hit Excel's "
                    f"32,767-character limit and were truncated."
                )
        metas_per_sheet.extend((title, metas) for title in titles)
        reports.append({
            "file": table["filename"],
            "sheets": titles,
            "rows": len(data),
            "columns": len(table["header"]),
            "encoding": prep["encoding"],
            "encoding_confidence": prep["confidence"],
            "delimiter": prep["delimiter"],
            "generated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "warnings": warnings,
            "types": [
                {
                    "column": meta.name,
                    "type": meta.label,
                    "format": meta.number_format or "",
                    "non_empty": meta.non_empty,
                    "empty": meta.empty,
                }
                for meta in metas
            ],
        })

    if opts.include_summary and reports:
        build_summary(wb, reports, metas_per_sheet)

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue(), reports


def convert_file(filename: str, raw: bytes, opts: Options) -> Tuple[bytes, Dict[str, Any]]:
    """Convert one CSV into (xlsx bytes, report dict) - its own workbook."""
    table = _prepare_table(filename, raw, opts)
    data, reports = _save_workbook([table], opts)
    return data, reports[0]


def convert_many(
    items: Sequence[Tuple[str, bytes]], opts: Options
) -> Tuple[List[Tuple[str, bytes]], List[Dict[str, Any]]]:
    """Convert several CSVs.

    single_workbook=True  -> [(name.xlsx, bytes)] with one sheet set
    single_workbook=False -> one workbook per CSV
    """
    if opts.single_workbook:
        tables = [_prepare_table(name, raw, opts) for name, raw in items]
        data, reports = _save_workbook(tables, opts)
        if len(items) == 1:
            out_name = f"{file_stem(items[0][0]) or 'converted'}.xlsx"
        else:
            out_name = "converted.xlsx"
        return [(out_name, data)], reports

    outputs: List[Tuple[str, bytes]] = []
    reports: List[Dict[str, Any]] = []
    used_names: set = set()
    for name, raw in items:
        table = _prepare_table(name, raw, opts)
        data, reps = _save_workbook([table], opts)
        base = file_stem(name) or "converted"
        candidate, k = f"{base}.xlsx", 2
        while candidate.lower() in used_names:
            candidate = f"{base} ({k}).xlsx"
            k += 1
        used_names.add(candidate.lower())
        outputs.append((candidate, data))
        reports.extend(reps)
    return outputs, reports


def preview_file(filename: str, raw: bytes, opts: Options, max_rows: int = 40) -> Dict[str, Any]:
    """Detect encoding/delimiter/types and return the first rows for the UI."""
    sample_raw = raw if len(raw) <= 8_000_000 else raw[:8_000_000]
    prep = _prepare(filename, sample_raw, opts)
    header, data = prep["header"], prep["data"]

    column_sample = data[:SAMPLE_SIZE]
    types = []
    for i in range(len(header)):
        meta = infer_column(header[i], [row[i] for row in column_sample], opts)
        types.append({
            "column": meta.name,
            "type": meta.label,
            "format": meta.number_format or "",
            "examples": meta.examples,
        })

    return {
        "file": filename,
        "encoding": prep["encoding"],
        "encoding_confidence": prep["confidence"],
        "delimiter": prep["delimiter"],
        "truncated_file": len(raw) > 8_000_000,
        "rows": len(data),
        "columns": len(header),
        "header": header,
        "preview": [[cell[:120] for cell in row] for row in data[:max_rows]],
        "types": types,
        "warnings": prep["warnings"],
    }
