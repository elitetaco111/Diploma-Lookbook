import csv
import sys
import re
from pathlib import Path
from urllib.parse import urlsplit, urlunsplit, parse_qsl, urlencode
from io import BytesIO

try:
    import segno
except ImportError:
    print("Missing dependency: segno. Install with: py -m pip install segno")
    sys.exit(1)

try:
    from PIL import Image
except ImportError:
    print("Missing dependency: pillow. Install with: py -m pip install pillow")
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

    svg_dir = Path(__file__).with_name("qr-codes-svg")
    jpg_dir = Path(__file__).with_name("qr-codes-jpg")
    svg_dir.mkdir(parents=True, exist_ok=True)
    jpg_dir.mkdir(parents=True, exist_ok=True)

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

            # Build tracked URL with UTM parameters
            tracked_url = add_utm(url, {
                "utm_source": UTM_SOURCE,
                "utm_medium": UTM_MEDIUM,
                "utm_campaign": UTM_CAMPAIGN,
                "utm_content": slugify(display_name),
            })

            safe_base = sanitize_filename(display_name)
            out_svg_path = unique_path(svg_dir, safe_base, ".svg")
            out_jpg_path = unique_path(jpg_dir, safe_base, ".jpg")

            try:
                qr = segno.make(tracked_url)

                # Save SVG
                qr.save(out_svg_path, border=2, scale=8)

                # Rasterize to JPG (via in-memory PNG -> Pillow -> JPG)
                buf = BytesIO()
                qr.save(buf, kind="png", border=2, scale=8)
                buf.seek(0)
                img = Image.open(buf)

                # Ensure opaque white background for JPG
                if img.mode in ("RGBA", "LA"):
                    bg = Image.new("RGB", img.size, (255, 255, 255))
                    alpha = img.getchannel("A") if "A" in img.getbands() else img.split()[-1]
                    bg.paste(img, mask=alpha)
                    img = bg
                else:
                    img = img.convert("RGB")

                img.save(out_jpg_path, "JPEG", quality=95, optimize=True)

                created += 1
            except Exception as e:
                print(f"Failed to create QR for '{display_name}': {e}")
                skipped += 1

    print(f"Done. Rows: {total}, Created: {created}, Skipped: {skipped}")
    print(f"SVG folder: {svg_dir.resolve()}")
    print(f"JPG folder: {jpg_dir.resolve()}")

if __name__ == "__main__":
    main()