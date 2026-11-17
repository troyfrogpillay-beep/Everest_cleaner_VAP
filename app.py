import io
import re
import difflib
from copy import copy
from dataclasses import dataclass, field
from datetime import datetime
from typing import Tuple, List, Dict, Any, Optional, NamedTuple

import openpyxl
from openpyxl.utils import get_column_letter
import pandas as pd
from email_validator import validate_email, EmailNotValidError
import streamlit as st

st.set_page_config(page_title="Everest Data Cleaner", layout="wide")

# ==============================================================================
# PRE-CLEANING STEPS: CLIENT SECTION EXTRACTION (ADDITIVE - NEVER TOUCHES THE
# SOURCE WORKSHEET OR THE DATAFRAME USED BY THE CLEANER)
# ==============================================================================

SUMMARY_SHEET_NAME = "Client Premium Summary"
DETAIL_SHEET_NAME = "Client Detailed Schedule"

SUMMARY_HEADERS = ["Payment Date", "Month End", "Policy Name", "Client Name", "Risk Payment"]
DETAIL_HEADERS = ["Client Name", "Member ID", "Surname", "First Name", "Policy Description", "Risk Premium"]


class ExtractionError(Exception):
    """Raised with a user-readable message when an extraction cannot proceed."""


class CellSnapshot(NamedTuple):
    """Independent copy of one source cell (value + type + formatting)."""
    value: Any
    data_type: str
    number_format: str
    font: Any
    fill: Any
    border: Any
    alignment: Any
    protection: Any


@dataclass
class ExtractedSection:
    section_label: str
    default_sheet_name: str
    source_sheet: str
    header_row: int
    first_row: int
    last_row: int
    first_col: int
    last_col: int
    header: List[CellSnapshot] = field(default_factory=list)
    rows: List[List[CellSnapshot]] = field(default_factory=list)
    source_rows: List[int] = field(default_factory=list)
    col_widths: Dict[int, float] = field(default_factory=dict)

    @property
    def section_range(self) -> str:
        return (f"'{self.source_sheet}'!{get_column_letter(self.first_col)}{self.header_row}:"
                f"{get_column_letter(self.last_col)}{self.last_row}")


def _norm_header(val: Any) -> str:
    if not isinstance(val, str):
        return ""
    return re.sub(r"\s+", " ", val.strip()).casefold()


def _norm_client(val: Any) -> str:
    if val is None:
        return ""
    return str(val).strip().casefold()


def _is_blank(val: Any) -> bool:
    return val is None or (isinstance(val, str) and not val.strip())


def _sheet_grid(ws) -> List[tuple]:
    return list(ws.iter_rows(values_only=True))


def _header_positions(row: tuple) -> Dict[str, List[int]]:
    pos: Dict[str, List[int]] = {}
    for idx, v in enumerate(row, start=1):
        key = _norm_header(v)
        if key:
            pos.setdefault(key, []).append(idx)
    return pos


def _matching_header_rows(grid: List[tuple], required: List[str]) -> List[Tuple[int, Dict[str, List[int]]]]:
    req = [_norm_header(h) for h in required]
    hits = []
    for r_idx, row in enumerate(grid, start=1):
        pos = _header_positions(row)
        if all(h in pos for h in req):
            hits.append((r_idx, pos))
    return hits


def _best_partial(grid: List[tuple], required: List[str]) -> Tuple[int, int, List[str]]:
    req = [(h, _norm_header(h)) for h in required]
    best = (0, 0, [h for h, _ in req])
    for r_idx, row in enumerate(grid, start=1):
        pos = _header_positions(row)
        found = [h for h, n in req if n in pos]
        if len(found) > best[1]:
            best = (r_idx, len(found), [h for h, n in req if n not in pos])
    return best


def _extract_section(wb, client_name: str, required: List[str], other_required: List[str],
                     label: str, default_sheet_name: str, stop_at_first_blank_row: bool) -> ExtractedSection:
    if client_name is None or not str(client_name).strip():
        raise ExtractionError(f"Client name is blank. Enter a client name to extract the {label}.")
    target = _norm_client(client_name)

    found: List[Tuple[Any, List[tuple], int, Dict[str, List[int]]]] = []
    best_overall = (None, 0, 0, list(required))
    for ws in wb.worksheets:
        grid = _sheet_grid(ws)
        for r_idx, pos in _matching_header_rows(grid, required):
            found.append((ws, grid, r_idx, pos))
        r, cnt, missing = _best_partial(grid, required)
        if cnt > best_overall[2]:
            best_overall = (ws.title, r, cnt, missing)

    if not found:
        if best_overall[2] >= 2:
            raise ExtractionError(
                f"Could not find the {label} header. The closest row is '{best_overall[0]}' row "
                f"{best_overall[1]}, which is missing: {', '.join(best_overall[3])}. "
                f"Required headers: {', '.join(required)}.")
        raise ExtractionError(
            f"Could not find the {label} header. No row contains the required headers: {', '.join(required)}.")

    if len(found) > 1:
        where = ", ".join(f"'{ws.title}' row {r}" for ws, _, r, _ in found)
        raise ExtractionError(
            f"Ambiguous {label} header: the required headers appear in more than one place ({where}). "
            f"Cannot tell which one to use.")

    ws, grid, header_row, pos = found[0]
    dupes = [h for h in required if len(pos[_norm_header(h)]) > 1]
    if dupes:
        raise ExtractionError(
            f"Ambiguous {label} header in '{ws.title}' row {header_row}: "
            f"duplicate column(s) {', '.join(dupes)}.")

    header_row_vals = grid[header_row - 1]
    used_cols = [i for i, v in enumerate(header_row_vals, start=1) if not _is_blank(v)]
    first_col, last_col = min(used_cols), max(used_cols)
    client_col = pos[_norm_header("Client Name")][0]

    other_headers = [r for r, _ in _matching_header_rows(grid, other_required) if r > header_row]
    limit = (min(other_headers) - 1) if other_headers else len(grid)
    first_row = header_row + 1

    def row_blank(r: int) -> bool:
        return all(_is_blank(v) for v in grid[r - 1][first_col - 1:last_col])

    if stop_at_first_blank_row:
        end = first_row - 1
        for r in range(first_row, limit + 1):
            if row_blank(r):
                break
            end = r
    else:
        end = limit
        while end >= first_row and row_blank(end):
            end -= 1

    section = ExtractedSection(
        section_label=label, default_sheet_name=default_sheet_name, source_sheet=ws.title,
        header_row=header_row, first_row=first_row, last_row=max(end, header_row),
        first_col=first_col, last_col=last_col)

    def snap(r: int, c: int) -> CellSnapshot:
        cell = ws.cell(row=r, column=c)
        return CellSnapshot(cell.value, cell.data_type, cell.number_format, copy(cell.font),
                            copy(cell.fill), copy(cell.border), copy(cell.alignment), copy(cell.protection))

    section.header = [snap(header_row, c) for c in range(first_col, last_col + 1)]
    seen_clients: Dict[str, str] = {}
    for r in range(first_row, end + 1):
        row = grid[r - 1]
        raw = row[client_col - 1] if client_col - 1 < len(row) else None
        key = _norm_client(raw)
        if key:
            seen_clients.setdefault(key, str(raw).strip())
        if key == target:
            section.rows.append([snap(r, c) for c in range(first_col, last_col + 1)])
            section.source_rows.append(r)

    if not section.rows:
        close = difflib.get_close_matches(target, list(seen_clients.keys()), n=3, cutoff=0.6)
        hint = f" Did you mean: {'; '.join(seen_clients[k] for k in close)}?" if close else ""
        raise ExtractionError(
            f"Client '{str(client_name).strip()}' was not found in the {label} "
            f"({section.section_range}, {end - first_row + 1} rows scanned).{hint}")

    for c in range(first_col, last_col + 1):
        dim = ws.column_dimensions.get(get_column_letter(c))
        if dim is not None and dim.width:
            section.col_widths[c - first_col + 1] = dim.width
    return section


def _open_source(source_bytes: bytes):
    try:
        return openpyxl.load_workbook(io.BytesIO(source_bytes), data_only=True)
    except Exception as e:
        raise ExtractionError(f"Could not open the file for extraction (.xlsx required): {e}")


def extract_client_premium_summary(source_bytes: bytes, client_name: str, workbook=None) -> ExtractedSection:
    wb = workbook if workbook is not None else _open_source(source_bytes)
    return _extract_section(wb, client_name, SUMMARY_HEADERS, DETAIL_HEADERS,
                            "premium summary", SUMMARY_SHEET_NAME, stop_at_first_blank_row=True)


def extract_client_detailed_schedule(source_bytes: bytes, client_name: str, workbook=None) -> ExtractedSection:
    wb = workbook if workbook is not None else _open_source(source_bytes)
    return _extract_section(wb, client_name, DETAIL_HEADERS, SUMMARY_HEADERS,
                            "detailed schedule", DETAIL_SHEET_NAME, stop_at_first_blank_row=False)


def run_client_extractions(source_bytes: bytes, client_name: str) -> Tuple[List[ExtractedSection], List[str]]:
    errors: List[str] = []
    if client_name is None or not str(client_name).strip():
        return [], ["Client name is blank. Enter a client name before running."]
    try:
        wb = _open_source(source_bytes)
    except ExtractionError as e:
        return [], [str(e)]
    sections: List[ExtractedSection] = []
    for fn in (extract_client_premium_summary, extract_client_detailed_schedule):
        try:
            sections.append(fn(source_bytes, client_name, workbook=wb))
        except ExtractionError as e:
            errors.append(str(e))
    return sections, errors


def _unique_sheet_title(book, desired: str) -> str:
    existing = {t.casefold() for t in book.sheetnames}
    if desired.casefold() not in existing:
        return desired
    n = 2
    while True:
        suffix = f" ({n})"
        title = desired[:31 - len(suffix)] + suffix
        if title.casefold() not in existing:
            return title
        n += 1


def write_section_to_workbook(book, section: ExtractedSection) -> str:
    ws = book.create_sheet(_unique_sheet_title(book, section.default_sheet_name))
    for r_i, cells in enumerate([section.header] + section.rows, start=1):
        for c_i, s in enumerate(cells, start=1):
            cell = ws.cell(row=r_i, column=c_i)
            cell.value = s.value
            if s.data_type == "s" and isinstance(s.value, str):
                cell.data_type = "s"
            cell.number_format = s.number_format
            cell.font, cell.fill, cell.border = copy(s.font), copy(s.fill), copy(s.border)
            cell.alignment, cell.protection = copy(s.alignment), copy(s.protection)
    for c_i, width in section.col_widths.items():
        ws.column_dimensions[get_column_letter(c_i)].width = width
    return ws.title


# ==============================================================================
# NAME HANDLING (RULE 5): SA-AWARE NAME / TITLE / INITIALS / SURNAME SPLITTING
# ==============================================================================

SA_SURNAME_PARTICLES = {
    "van", "von", "der", "den", "de", "du", "di", "le", "la", "ten", "ter", "te",
    "op", "het", "da", "dos", "das", "do", "del", "della", "el", "al", "bin",
}

NAME_TITLES = {
    "mr": "Mr", "mrs": "Mrs", "ms": "Ms", "miss": "Miss", "mx": "Mx",
    "mnr": "Mnr", "mev": "Mev", "me": "Me", "dr": "Dr", "prof": "Prof",
    "professor": "Prof", "adv": "Adv", "advocate": "Adv", "rev": "Rev",
    "pastor": "Pastor", "hon": "Hon", "sir": "Sir", "ds": "Ds",
}

_NAME_CHARS = r"A-Za-z\u00C0-\u00FF"
_GLUED_PARTICLE = re.compile(r"^(van|von|du|de|le|la|di|ten|ter)([A-Z][" + _NAME_CHARS + r"'\-]+)$")

_FIRST_NAME_HINTS = ("first name", "firstname", "given name", "forename", "middle name",
                     "second name", "other name", "name")
_FULL_NAME_EXACT = {"name", "full name", "fullname", "member name", "employee name",
                    "person name", "member full name"}
_NON_PERSON_HINTS = ("client", "policy", "category", "paypoint", "short", "file", "fast", "scheme", "fund")


def _split_glued(word: str) -> List[str]:
    m = _GLUED_PARTICLE.match(word)
    return [m.group(1), m.group(2)] if m else [word]


def _cap_segment(seg: str) -> str:
    if "'" in seg:
        return "'".join(x.capitalize() for x in seg.split("'"))
    if seg.startswith("mc") and len(seg) > 2:
        return "Mc" + seg[2:].capitalize()
    return seg.capitalize()


def format_person_name(val: Any, surname: bool = False) -> str:
    if val is None or (not isinstance(val, str) and pd.isna(val)):
        return ""
    text = re.sub(r"[^" + _NAME_CHARS + r"\s\-']", "", str(val))
    words: List[str] = []
    for w in text.split():
        words.extend(_split_glued(w) if surname else [w])
    out = []
    for i, w in enumerate(words):
        lw = w.lower()
        if surname and i > 0 and lw in SA_SURNAME_PARTICLES:
            out.append(lw)
        else:
            out.append("-".join(_cap_segment(seg) for seg in lw.split("-")))
    return " ".join(out)


def _title_of(tok: str) -> Optional[str]:
    return NAME_TITLES.get(tok.strip(".,()").casefold())


def _is_initials(tok: str) -> bool:
    core = tok.strip(",")
    letters = core.replace(".", "")
    if not letters or not letters.isalpha() or core.casefold().strip(".") in SA_SURNAME_PARTICLES:
        return False
    if len(letters) == 1:
        return True
    return "." in core or (letters.isupper() and len(letters) <= 2)


def _tokens(text: str) -> List[str]:
    toks: List[str] = []
    for t in re.split(r"\s+", text.replace("(", " ").replace(")", " ")):
        if t:
            toks.extend(_split_glued(t.rstrip(",")))
    return toks


def parse_person_name(raw: Any) -> Dict[str, str]:
    """Splits a single name string into Title, Name (all given names), Initials, Surname.
    Handles 'Zulu Mr AA', 'Mr J Smith', 'Dr. J.P. van der Merwe', 'Smith, J', 'Johan Pieter du Preez'."""
    text = re.sub(r"\s+", " ", str(raw)).strip()
    result = {"Title": "", "Name": "", "Initials": "", "Surname": ""}
    if not text:
        return result

    toks = text.split()
    
    # 1. Check for "Surname Title Initials" layout (e.g., "Zulu Mr AA") where a title appears after the first token
    title = ""
    title_idx = -1
    for idx, t in enumerate(toks):
        k = _title_of(t)
        if k and idx > 0:
            title = k
            title_idx = idx
            break

    if title_idx > 0:
        surname_toks = toks[:title_idx]
        rest = toks[title_idx + 1:]
        initials_parts = [r for r in rest if _is_initials(r) or len(r.replace(".", "")) <= 2]
        name_parts = [r for r in rest if not (_is_initials(r) or len(r.replace(".", "")) <= 2)]
        
        result["Title"] = title
        result["Surname"] = format_person_name(" ".join(surname_toks), surname=True)
        result["Initials"] = "".join(c for t in initials_parts for c in t if c.isalpha()).upper()
        result["Name"] = format_person_name(" ".join(name_parts))
        return result

    # 2. Standard parsing for other formats
    def strip_titles(toks: List[str]) -> List[str]:
        nonlocal title
        out = []
        for t in toks:
            k = _title_of(t)
            if k and len(toks) > 1:
                title = title or k
            else:
                out.append(t)
        return out

    names: List[str] = []
    initials: List[str] = []
    if "," in text:
        left, right = text.split(",", 1)
        surname_toks = strip_titles(_tokens(left))
        for t in strip_titles(_tokens(right)):
            (initials if _is_initials(t) else names).append(t)
    else:
        toks = strip_titles(_tokens(text))
        lead: List[str] = []
        trail: List[str] = []
        while len(toks) > 1 and _is_initials(toks[0]):
            lead.append(toks.pop(0))
        while len(toks) > 1 and _is_initials(toks[-1]):
            trail.insert(0, toks.pop())
        if trail and not lead:
            surname_toks, initials = toks, trail
        else:
            i = len(toks) - 1
            while i > 0 and toks[i - 1].casefold() in SA_SURNAME_PARTICLES:
                i -= 1
            surname_toks, names = toks[i:], toks[:i]
            initials = lead + trail
        if len(toks) == 1 and not (title or initials):
            surname_toks, names = [], toks

    result["Title"] = title
    result["Name"] = format_person_name(" ".join(names))
    result["Initials"] = "".join(c for t in initials for c in t if c.isalpha()).upper()
    result["Surname"] = format_person_name(" ".join(surname_toks), surname=True)
    return result


def find_name_columns(columns: List[str]) -> Dict[str, Any]:
    low = {c: c.lower().strip() for c in columns}
    ok = lambda c: not any(h in low[c] for h in _NON_PERSON_HINTS)
    return {
        "surname": next((c for c in columns if "surname" in low[c] or "last name" in low[c] or "lastname" in low[c]), None),
        "first": [c for c in columns if ok(c) and any(h in low[c] for h in _FIRST_NAME_HINTS)],
        "full": next((c for c in columns if low[c] in _FULL_NAME_EXACT or "full" in low[c] or "name" in low[c]), None),
        "initials": next((c for c in columns if "initial" in low[c]), None),
        "title": next((c for c in columns if low[c] in ("title", "salutation")), None),
    }


def resolve_name_fields(row: pd.Series, cols: Dict[str, Any], row_num: int, logs: List[Dict[str, Any]]) -> Dict[str, str]:
    def get(c: Optional[str]) -> str:
        if not c or c not in row.index:
            return ""
        v = row[c]
        return "" if v is None or pd.isna(v) else re.sub(r"\s+", " ", str(v)).strip()

    out = {"Title": "", "Name": "", "Initials": "", "Surname": ""}
    sur_raw = get(cols["surname"])
    given_raw = " ".join(x for x in (get(c) for c in cols["first"]) if x)

    if sur_raw:
        toks = _tokens(sur_raw)
        composite = len(toks) > 1 and any(_title_of(t) or _is_initials(t) for t in toks)
        if composite:
            parsed = parse_person_name(sur_raw)
            out.update(parsed)
            logs.append({"Row": row_num, "Field": "Name", "Action": f"Split '{sur_raw}' into Title/Initials/Surname"})
        else:
            out["Surname"] = format_person_name(sur_raw, surname=True)
    elif get(cols["full"]):
        full_raw = get(cols["full"])
        parsed = parse_person_name(full_raw)
        out.update(parsed)
        logs.append({"Row": row_num, "Field": "Name/Surname", "Action": f"Split '{full_raw}' into Name/Title/Initials & Surname"})
        if not parsed["Surname"]:
            logs.append({"Row": row_num, "Field": "Name/Surname", "Action": f"Only one name found in '{full_raw}'; no surname split"})

    if given_raw and given_raw != sur_raw:
        out["Name"] = format_person_name(given_raw)
    if get(cols["initials"]):
        out["Initials"] = "".join(c for c in get(cols["initials"]) if c.isalpha()).upper()
    if get(cols["title"]):
        out["Title"] = NAME_TITLES.get(get(cols["title"]).strip(".").casefold(), format_person_name(get(cols["title"])))
    return out


# ==============================================================================
# HELPER FUNCTIONS (IMPLEMENTING THE RULES)
# ==============================================================================

def clean_text_only(val: Any) -> str:
    if pd.isna(val):
        return ""
    text = str(val).strip()
    text = re.sub(r'[^a-zA-Z\s\-]', '', text)
    return text.title()


def format_date_of_birth(val: Any) -> str:
    if pd.isna(val) or not str(val).strip():
        return ""
    val_str = str(val).strip()
    try:
        dt = pd.to_datetime(val_str, dayfirst=True)
        return dt.strftime("%d/%m/%Y")
    except Exception:
        return val_str


def process_sa_id_or_passport(id_val: Any, raw_dob_val: Any) -> Tuple[str, str, str, str, str]:
    if pd.isna(id_val):
        formatted_dob = format_date_of_birth(raw_dob_val)
        return "", "", formatted_dob, "", ""
    
    clean_id = re.sub(r'\D', '', str(id_val).strip())

    if len(clean_id) == 13:
        try:
            yy, mm, dd = int(clean_id[0:2]), int(clean_id[2:4]), int(clean_id[4:6])
            current_yy = int(datetime.now().strftime("%y"))
            yyyy = 2000 + yy if yy <= current_yy else 1900 + yy

            dob_dt = datetime(yyyy, mm, dd)
            derived_dob = dob_dt.strftime("%d/%m/%Y")

            gender_digit = int(clean_id[6:10])
            gender = "Female" if gender_digit < 5000 else "Male"

            return clean_id, "", derived_dob, gender, "Valid SA ID"
        except ValueError:
            pass

    raw_str = str(id_val).strip()
    formatted_dob = format_date_of_birth(raw_dob_val)
    return "", raw_str, formatted_dob, "", f"Moved '{raw_str}' to Passport (Invalid SA ID)"


def process_phone(val: Any) -> Tuple[str, bool]:
    if pd.isna(val):
        return "", False
    num = re.sub(r'\D', '', str(val).strip())
    if num.startswith("27") and len(num) == 11:
        num = "0" + num[2:]
    if len(num) == 10 and num.startswith("0"):
        return num, True
    return str(val).strip(), False


def process_email(val: Any) -> Tuple[str, bool]:
    if pd.isna(val):
        return "", False
    raw = str(val).strip()
    try:
        valid = validate_email(raw, check_deliverability=False)
        return valid.normalized, True
    except EmailNotValidError:
        return raw, False


def format_currency(val: Any) -> Tuple[float, str]:
    if pd.isna(val):
        return 0.0, "0.00"
    num_str = re.sub(r'[^0-9.]', '', str(val).replace(',', ''))
    try:
        num_float = float(num_str)
        return num_float, f"{num_float:,.2f}"
    except ValueError:
        return 0.0, "0.00"


def run_everest_pipeline(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Executes transformation rules, splits names, corrects DOB via ID or Passport, and compiles an audit report."""
    logs: List[Dict[str, Any]] = []

    df.columns = [re.sub(r'[^a-zA-Z0-9_\s]', '', str(c)).strip() for c in df.columns]

    name_cols = find_name_columns(list(df.columns))
    id_col = next((c for c in df.columns if 'id' in c.lower() or 'identity' in c.lower()), None)
    dob_col = next((c for c in df.columns if 'dob' in c.lower() or 'birth' in c.lower()), None)
    gender_col = next((c for c in df.columns if 'gender' in c.lower() or 'sex' in c.lower()), None)
    email_col = next((c for c in df.columns if 'email' in c.lower() or 'mail' in c.lower()), None)
    cell_col = next((c for c in df.columns if 'cell' in c.lower() or 'phone' in c.lower() or 'mobile' in c.lower()), None)
    prem_cols = [c for c in df.columns if 'prem' in c.lower() or 'amount' in c.lower() or 'price' in c.lower()]

    replaced_cols = set()
    for c in [name_cols["surname"], name_cols["full"], name_cols["initials"], name_cols["title"],
              id_col, dob_col, gender_col, email_col, cell_col]:
        if c:
            replaced_cols.add(c)
    replaced_cols.update(name_cols["first"])

    cleaned_rows = []

    for idx, row in df.iterrows():
        row_num = idx + 2
        out_row: Dict[str, Any] = {}

        out_row.update(resolve_name_fields(row, name_cols, row_num, logs))

        raw_id = row.get(id_col, "") if id_col else ""
        raw_dob = row.get(dob_col, "") if dob_col else ""
        sa_id, passport, derived_dob, derived_gender, id_log = process_sa_id_or_passport(raw_id, raw_dob)

        out_row["SA ID Number"] = str(sa_id)
        out_row["Passport Number"] = str(passport)
        if id_log.startswith("Moved"):
            logs.append({"Row": row_num, "Field": "ID/Passport", "Action": id_log})

        out_row["Date of Birth"] = derived_dob
        if raw_id and len(re.sub(r'\D', '', str(raw_id))) == 13 and raw_dob:
            formatted_raw_dob = format_date_of_birth(raw_dob)
            if formatted_raw_dob and formatted_raw_dob != derived_dob:
                logs.append({
                    "Row": row_num,
                    "Field": "DOB",
                    "Action": f"Corrected provided DOB '{formatted_raw_dob}' to ID-derived DOB '{derived_dob}'"
                })

        raw_gender = clean_text_only(row.get(gender_col, "")) if gender_col else ""
        out_row["Gender"] = derived_gender if derived_gender else raw_gender

        email, valid_email = process_email(row.get(email_col, "")) if email_col else ("", False)
        out_row["Email"] = email
        if not valid_email and email:
            logs.append({"Row": row_num, "Field": "Email", "Action": f"Flagged invalid email format: '{email}'"})

        cell, valid_cell = process_phone(row.get(cell_col, "")) if cell_col else ("", False)
        out_row["Cell Number"] = str(cell)
        if not valid_cell and cell:
            logs.append({"Row": row_num, "Field": "Cell Number", "Action": f"Flagged non-standard SA cell number: '{cell}'"})

        if prem_cols:
            for p_col in prem_cols:
                f_val, formatted_str = format_currency(row[p_col])
                out_row[p_col] = formatted_str
                out_row[f"_raw_num_{p_col}"] = f_val
        elif "Premium" in df.columns:
            f_val, formatted_str = format_currency(row.get("Premium", 0))
            out_row["Premium"] = formatted_str
            out_row["_raw_num_Premium"] = f_val

        for col in df.columns:
            if col not in replaced_cols and col not in out_row:
                val = row[col]
                out_row[col] = "" if pd.isna(val) else str(val).strip()

        cleaned_rows.append(out_row)

    cleaned_df = pd.DataFrame(cleaned_rows)

    raw_num_cols = [c for c in cleaned_df.columns if c.startswith("_raw_num_")]
    if raw_num_cols:
        total_sum = cleaned_df[raw_num_cols].sum().sum()
        total_row = {col: "" for col in cleaned_df.columns if not col.startswith("_raw_num_")}
        total_row["Surname"] = "TOTAL"
        target_prem_col = prem_cols[0] if prem_cols else "Premium"
        if target_prem_col in total_row:
            total_row[target_prem_col] = f"{total_sum:,.2f}"
        cleaned_df = pd.concat([cleaned_df, pd.DataFrame([total_row])], ignore_index=True)

    cleaned_df = cleaned_df.drop(columns=raw_num_cols, errors='ignore')
    audit_df = pd.DataFrame(logs) if logs else pd.DataFrame(columns=["Row", "Field", "Action"])
    return cleaned_df, audit_df


# ==============================================================================
# STREAMLIT USER INTERFACE
# ==============================================================================

st.title("Everest Migration Data Cleaner")
st.markdown("Upload raw client files (CSV or Excel) below to automatically clean and transform them according to Everest specifications.")

uploaded_file = st.file_uploader("Upload client data file", type=["csv", "xlsx", "xls"])

if uploaded_file is not None:
    source_bytes = uploaded_file.getvalue()

    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, dtype=str)
        else:
            temp_df = pd.read_excel(uploaded_file, header=None, dtype=str)
            header_row_idx = 0
            for r_idx, row_vals in temp_df.iterrows():
                if any(str(v).strip().upper() in ["NAME", "CLIENT NAME", "NEW EMPLOYEE NUMBER"] for v in row_vals.values if pd.notna(v)):
                    header_row_idx = r_idx
                    break
            df = pd.read_excel(uploaded_file, header=header_row_idx, dtype=str)
    except Exception as e:
        st.error(f"Error reading file: {e}")
        st.stop()

    st.subheader("Raw Data Preview")
    st.dataframe(df.head(5))

    can_extract = uploaded_file.name.lower().endswith(".xlsx")
    client_name = ""
    if can_extract:
        client_name = st.text_input(
            "Client name (optional)",
            help="Leave blank to clean the entire upload. If you enter a name, the Client Detailed "
                 "Schedule and Premium Summary are extracted, and cleaning is run on that client's schedule.")
    else:
        st.info("Client section extraction needs an .xlsx file; it will be skipped for this upload.")

    if st.button("Run Cleaning Pipeline"):
        extracted_sections: List[ExtractedSection] = []
        run_extraction = can_extract and bool(client_name.strip())

        if run_extraction:
            extracted_sections, extraction_errors = run_client_extractions(source_bytes, client_name)
            if extraction_errors:
                for msg in extraction_errors:
                    st.error(msg)
                st.stop()
            for sec in extracted_sections:
                st.info(f"Extracted {len(sec.rows)} {sec.section_label} row(s) for "
                        f"'{client_name.strip()}' from {sec.section_range}.")

            detailed_sec = next((sec for sec in extracted_sections if "detailed" in sec.section_label.lower()), None)
            if detailed_sec and detailed_sec.rows:
                h_vals = [str(c.value) if c.value is not None else "" for c in detailed_sec.header]
                r_vals = [[(c.value if c.value is not None else "") for c in r] for r in detailed_sec.rows]
                df = pd.DataFrame(r_vals, columns=h_vals, dtype=str)
            else:
                st.warning("Client found in summary, but detailed schedule rows were empty. Cleaning entire upload as fallback.")
        elif can_extract:
            st.info("No client name entered, so the entire upload will be cleaned.")

        with st.spinner("Processing records and validating data..."):
            cleaned_df, audit_df = run_everest_pipeline(df)

        st.success("Cleaning completed successfully!")

        total_records = len(cleaned_df)
        if "Surname" in cleaned_df.columns and "TOTAL" in cleaned_df["Surname"].values:
            total_records -= 1

        total_audits = len(audit_df)

        col1, col2 = st.columns(2)
        with col1:
            st.metric(label="Total Records Processed", value=total_records)
        with col2:
            st.metric(
                label="Flagged Audit Actions",
                value=total_audits,
                delta=-total_audits if total_audits > 0 else 0,
                delta_color="inverse"
            )

        st.subheader("Cleaned Data Preview")
        st.dataframe(cleaned_df.head(10))

        st.subheader("Audit Report Preview")
        st.dataframe(audit_df.head(10))

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            cleaned_df.to_excel(writer, index=False, sheet_name="Everest_Cleaned_Data")
            audit_df.to_excel(writer, index=False, sheet_name="Audit_Report")
            for sec in extracted_sections:
                write_section_to_workbook(writer.book, sec)
        output.seek(0)

        base_name = uploaded_file.name.rsplit('.', 1)[0]
        suffix = f"_{client_name.strip().replace(' ', '_')}" if client_name.strip() else ""
        st.download_button(
            label="📥 Download Everest-Ready File (With Audit Report)",
            data=output.getvalue(),
            file_name=f"everest_ready_{base_name}{suffix}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )