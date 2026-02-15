import re
from typing import Optional, Tuple

FILE_NUMBER_RE = re.compile(r"^\d{2}-\d{3}$")
SAMPLE_NAME_RE = re.compile(r"^(B|T|HA|C)-\d{1,3}$", re.IGNORECASE)


def is_valid_file_number(value: str) -> bool:
    return bool(FILE_NUMBER_RE.match(value.strip()))


def is_valid_sample_name(value: str) -> bool:
    return bool(SAMPLE_NAME_RE.match(value.strip()))


def parse_depth(depth_raw: str) -> Tuple[Optional[float], Optional[float], Optional[str]]:
    """
    Accepts strings like 1.0'-2.0' or 40.0''-50.0''.
    Returns (from, to, unit) where unit is 'ft' or 'in'.
    """
    if not depth_raw:
        return None, None, None

    raw = depth_raw.strip().replace(" ", "")
    # Detect inches (double apostrophe) or feet (single)
    unit = None
    if "''" in raw:
        unit = "in"
        raw = raw.replace("''", "")
    elif "'" in raw:
        unit = "ft"
        raw = raw.replace("'", "")

    if "-" not in raw:
        return None, None, unit

    left, right = raw.split("-", 1)
    try:
        return float(left), float(right), unit
    except ValueError:
        return None, None, unit