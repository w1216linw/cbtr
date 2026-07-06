import re

STREET_ABBREVS = [
    (r'\bstreet\b',    'st'),
    (r'\bavenue\b',    'ave'),
    (r'\bboulevard\b', 'blvd'),
    (r'\broad\b',      'rd'),
    (r'\bdrive\b',     'dr'),
    (r'\bparkway\b',   'pkwy'),
    (r'\bcourt\b',     'ct'),
    (r'\blane\b',      'ln'),
    (r'\bplace\b',     'pl'),
    (r'\bcircle\b',    'cir'),
    (r'\bhighway\b',   'hwy'),
]

UNIT_RE = re.compile(
    r'\s+(#.*|unit\s+.*|suites?\s+.*|ste\s+.*|ste$|apt\.?\s+.*|loading\b.*|dock\b.*)$',
    re.IGNORECASE,
)


def _norm(s: str) -> str:
    s = s.lower().strip()
    s = s.replace('.', '')
    s = re.sub(r'(\d)([a-z])(\s|$)', r'\1 \2\3', s)
    s = re.sub(r'#\s*(\w+)', r'unit \1', s)
    for pat, abbr in STREET_ABBREVS:
        s = re.sub(pat, abbr, s)
    return re.sub(r'\s+', ' ', s).strip()


def norm_full(s: str) -> str:
    return _norm(s) if s else ''


def norm_base(s: str) -> str:
    return _norm(UNIT_RE.sub('', s).strip()) if s else ''


def parse_db_addr(addr: str):
    """Returns (base, full, city) for one address_db entry."""
    parts = [p.strip() for p in str(addr).split(',')]
    raw_street = parts[0]
    base = norm_base(raw_street)
    full = norm_full(raw_street)

    city = ''
    for part in parts[1:]:
        p = part.strip()
        if not p:
            continue
        if re.fullmatch(r'[A-Z]{2}', p):
            continue
        if re.match(r'^[A-Z]{2}\s+\d{5}', p):
            continue
        if re.fullmatch(r'\d{5}(-\d{4})?', p):
            continue
        if p.upper() in ('USA', 'UNITED STATES'):
            continue
        city = p.lower()
        break
    return base, full, city
