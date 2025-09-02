import csv
import sys
import re
import json
from pathlib import Path
from urllib.parse import urlsplit, urlunsplit, parse_qsl, urlencode

try:
    import segno
except ImportError:
    print("Missing dependency: segno. Install with: py -m pip install segno")
    sys.exit(1)


def find_column(fieldnames, target_name):
    """Find a column by case-insensitive, trimmed match."""
    norm_map = { (f or "").strip().lower(): f for f in (fieldnames or []) }
    return norm_map.get(target_name.strip().lower())


INVALID_CHARS_RE = re.compile(r'[<>:"/\\|?*\x00-\x1F]')
RESERVED_NAMES = {
    "CON","PRN","AUX","NUL",
    *(f"COM{i}" for i in range(1,10)),
    *(f"LPT{i}" for i in range(1,10)),
}


def sanitize_filename(name: str, max_len: int = 150) -> str:
    """Sanitize for Windows filenames while keeping name recognizable."""
    if not name:
        return "untitled"
    name = name.strip().rstrip(". ")  # no trailing dot/space
    name = INVALID_CHARS_RE.sub("", name)
    if name.upper() in RESERVED_NAMES:
        name = f"{name}_file"
    # collapse whitespace
    name = re.sub(r"\s+", " ", name).strip()
    if not name:
        name = "untitled"
    # limit length (leave space for extension and possible suffix)
    return name[:max_len]


def unique_path(directory: Path, base: str, ext: str) -> Path:
    """Ensure unique filename by appending (n) if needed."""
    candidate = directory / f"{base}{ext}"
    if not candidate.exists():
        return candidate
    n = 2
    while True:
        cand = directory / f"{base} ({n}){ext}"
        if not cand.exists():
            return cand
        n += 1


# Where your redirector is hosted (change to your domain)
BASE_REDIRECT = "https://your-domain.example"  # e.g., https://lookbook.example.com

# UTM defaults (edit as needed)
UTM_SOURCE = "qr"
UTM_MEDIUM = "offline"
UTM_CAMPAIGN = "diploma_lookbook"

def slugify(text: str) -> str:
    t = (text or "").lower().strip()
    t = re.sub(r"[^a-z0-9]+", "-", t)
    t = re.sub(r"-{2,}", "-", t).strip("-")
    return t or "item"

def add_utm(url: str, params: dict) -> str:
    """Append UTM params if not already present."""
    parts = urlsplit(url)
    pairs = parse_qsl(parts.query, keep_blank_values=True)
    existing = {k.lower() for k, _ in pairs}
    for k, v in params.items():
        if v and k.lower() not in existing:
            pairs.append((k, v))
    new_query = urlencode(pairs)
    return urlunsplit((parts.scheme, parts.netloc, parts.path, new_query, parts.fragment))


def main():
    # CSV path: argument or default in project root
    csv_path = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).with_name("DRNItemDiplomaFramesResults569.csv")
    if not csv_path.exists():
        print(f"CSV not found: {csv_path}")
        sys.exit(1)

    out_dir = Path(__file__).with_name("qr-codes")
    out_dir.mkdir(parents=True, exist_ok=True)

    # mappings.json will live under redirector/
    redirector_dir = Path(__file__).parent / "redirector"
    redirector_dir.mkdir(parents=True, exist_ok=True)
    mappings_path = redirector_dir / "mappings.json"
    mappings: dict[str, str] = {}
    used_codes: set[str] = set()

    total = 0
    created = 0
    skipped = 0

    # Read with utf-8-sig to handle BOM from Excel exports
    with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        if not reader.fieldnames:
            print("No headers found in CSV.")
            sys.exit(1)

        url_col = find_column(reader.fieldnames, "URL Absolute")
        name_col = find_column(reader.fieldnames, "Store Display Name")

        if not url_col or not name_col:
            print("Required columns not found.")
            print(f"Found headers: {reader.fieldnames}")
            print("Need: 'URL Absolute' and 'Store Display Name'")
            sys.exit(1)

        for row in reader:
            total += 1
            url = (row.get(url_col) or "").strip()
            display_name = (row.get(name_col) or "").strip()

            if not url or not display_name:
                skipped += 1
                continue

            # Destination with UTM parameters
            tracked_url = add_utm(url, {
                "utm_source": UTM_SOURCE,
                "utm_medium": UTM_MEDIUM,
                "utm_campaign": UTM_CAMPAIGN,
                "utm_content": slugify(display_name),
            })

            # Create a stable code and ensure uniqueness in this run
            base_code = slugify(display_name)
            code = base_code or "item"
            i = 2
            while code in used_codes:
                code = f"{base_code}-{i}"
                i += 1
            used_codes.add(code)
            mappings[code] = tracked_url

            safe_base = sanitize_filename(display_name)
            out_path = unique_path(out_dir, safe_base, ".svg")

            try:
                # The QR encodes your redirect URL, not the final URL
                qr_url = f"{BASE_REDIRECT}/r/{code}"
                qr = segno.make(qr_url)
                qr.save(out_path, border=2, scale=8)
                created += 1
            except Exception as e:
                print(f"Failed to create QR for '{display_name}': {e}")
                skipped += 1

    # Write mappings.json for the redirector
    try:
        with mappings_path.open("w", encoding="utf-8") as mf:
            json.dump(mappings, mf, indent=2, ensure_ascii=False)
    except Exception as e:
        print(f"Failed to write mappings.json: {e}")
        sys.exit(1)

    print(f"Done. Rows: {total}, Created: {created}, Skipped: {skipped}")
    print(f"QR output: {out_dir.resolve()}")
    print(f"Mappings: {mappings_path.resolve()}")


if __name__ == "__main__":
    main()