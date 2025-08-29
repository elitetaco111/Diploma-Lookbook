import csv
import sys
import re
from pathlib import Path

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


def main():
    # CSV path: argument or default in project root
    csv_path = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).with_name("DRNItemDiplomaFramesResults569.csv")
    if not csv_path.exists():
        print(f"CSV not found: {csv_path}")
        sys.exit(1)

    out_dir = Path(__file__).with_name("qr-codes")
    out_dir.mkdir(parents=True, exist_ok=True)

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

            safe_base = sanitize_filename(display_name)
            out_path = unique_path(out_dir, safe_base, ".svg")

            try:
                qr = segno.make(url)  # auto error correction
                # Save a crisp SVG suitable for print/display
                qr.save(out_path, border=2, scale=8)  # adjust scale as needed
                created += 1
            except Exception as e:
                print(f"Failed to create QR for '{display_name}': {e}")
                skipped += 1

    print(f"Done. Rows: {total}, Created: {created}, Skipped: {skipped}")
    print(f"Output folder: {out_dir.resolve()}")


if __name__ == "__main__":
    main()