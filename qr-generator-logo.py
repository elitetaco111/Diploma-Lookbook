import csv
import sys
import re
from pathlib import Path
from urllib.parse import urlsplit, urlunsplit, parse_qsl, urlencode
from io import BytesIO
import base64

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

# --- QR + Logo configuration ---
ERROR_CORRECTION = "h"   # Use high error correction to tolerate center logo
BORDER = 4               # Quiet zone modules (2–4 typical)
SCALE = 12                # Pixel scale per module (increase for higher resolution)

# Logo configuration
LOGO_PATH = Path(__file__).with_name("logo.png")  # Put your logo image here
LOGO_SCALE = 0.4          # Logo width as fraction of QR width (0.15–0.25 recommended)
LOGO_PAD = 1              # White box padding (pixels) around the logo in raster/JPG
# -------------------------------

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

# UTM vars
UTM_SOURCE = "qr"
UTM_MEDIUM = "direct"
UTM_CAMPAIGN = "diploma_insert"

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


def add_logo_on_pil(img: Image.Image, logo_path: Path, logo_scale: float, pad: int) -> Image.Image:
    """Overlay a centered logo with a white background box on a PIL image of the QR."""
    if not logo_path or not logo_path.exists():
        return img

    base = img.convert("RGBA")
    W, H = base.size

    # Load and scale the logo
    logo = Image.open(logo_path).convert("RGBA")
    target_w = max(1, int(W * logo_scale))
    # Keep aspect ratio
    lw, lh = logo.size
    aspect = lw / lh if lh else 1
    new_w = target_w
    new_h = max(1, int(new_w / aspect))
    logo = logo.resize((new_w, new_h), Image.LANCZOS)

    # White background box behind the logo
    bg_w, bg_h = new_w + 2 * pad, new_h + 2 * pad
    bg = Image.new("RGBA", (bg_w, bg_h), (255, 255, 255, 255))

    # Positions
    x0 = (W - bg_w) // 2
    y0 = (H - bg_h) // 2

    # Composite background and logo
    base.alpha_composite(bg, dest=(x0, y0))
    base.alpha_composite(logo, dest=(x0 + pad, y0 + pad))

    return base


def embed_logo_in_svg(svg_path: Path, logo_path: Path, total_w: int, total_h: int,
                      logo_scale: float, pad: int):
    """Embed a raster logo into the SVG as base64 at the center, with a white box behind."""
    if not logo_path or not logo_path.exists():
        return

    # Compute logo and box sizes in SVG user units (same units as width/height)
    Lw = max(1, int(total_w * logo_scale))
    # Compute logo height from actual logo aspect
    with Image.open(logo_path) as im:
        lw, lh = im.size
        aspect = lw / lh if lh else 1
    Lh = max(1, int(Lw / aspect))
    BgW, BgH = Lw + 2 * pad, Lh + 2 * pad

    X = (total_w - BgW) // 2
    Y = (total_h - BgH) // 2

    # Encode the logo as base64 (PNG recommended)
    mime = "image/png"
    with open(logo_path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("ascii")

    overlay = f'''
  <g id="qr-logo">
    <rect x="{X}" y="{Y}" width="{BgW}" height="{BgH}" fill="#ffffff"/>
    <image href="data:{mime};base64,{b64}" x="{X + pad}" y="{Y + pad}" width="{Lw}" height="{Lh}" />
  </g>
</svg>'''
    # Insert overlay before closing </svg>
    svg_text = svg_path.read_text(encoding="utf-8")
    if "</svg>" in svg_text:
        svg_text = svg_text.replace("</svg>", overlay)
        svg_path.write_text(svg_text, encoding="utf-8")


def main():
    # CSV path: argument or default in project root
    csv_path = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).with_name("diploma_qr_codes.csv")
    if not csv_path.exists():
        print(f"CSV not found: {csv_path}")
        sys.exit(1)

    svg_dir = Path(__file__).with_name("qr-codes-svg-logo")
    jpg_dir = Path(__file__).with_name("qr-codes-jpg-logo")
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
        file_name_col = find_column(reader.fieldnames, "Name")
        display_col = find_column(reader.fieldnames, "Store Display Name")  # optional

        if not url_col or not file_name_col:
            print("Required columns not found.")
            print(f"Found headers: {reader.fieldnames}")
            print("Need: 'URL Absolute' and 'Name'")
            sys.exit(1)

        for row in reader:
            total += 1
            url = (row.get(url_col) or "").strip()
            file_name = (row.get(file_name_col) or "").strip()
            display_name = (row.get(display_col) or "").strip() if display_col else ""

            if not url or not file_name:
                skipped += 1
                continue

            # Build tracked URL with UTM parameters
            utm_content = slugify(display_name or file_name)
            tracked_url = add_utm(url, {
                "utm_source": UTM_SOURCE,
                "utm_medium": UTM_MEDIUM,
                "utm_campaign": UTM_CAMPAIGN,
                "utm_content": utm_content,
            })

            safe_base = sanitize_filename(file_name)
            out_svg_path = unique_path(svg_dir, safe_base, ".svg")
            out_jpg_path = unique_path(jpg_dir, safe_base, ".jpg")

            try:
                # Generate QR with high error correction
                qr = segno.make(tracked_url, error=ERROR_CORRECTION)

                # Save SVG
                qr.save(out_svg_path, border=BORDER, scale=SCALE)

                # Embed logo into SVG (uses same total size as rendered)
                total_w, total_h = qr.symbol_size(scale=SCALE, border=BORDER)
                embed_logo_in_svg(out_svg_path, LOGO_PATH, total_w, total_h, LOGO_SCALE, LOGO_PAD)

                # Rasterize to PNG in-memory
                buf = BytesIO()
                qr.save(buf, kind="png", border=BORDER, scale=SCALE)
                buf.seek(0)
                img = Image.open(buf)

                # Overlay logo for raster/JPG
                img = add_logo_on_pil(img, LOGO_PATH, LOGO_SCALE, LOGO_PAD)

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
                print(f"Failed to create QR for '{file_name}': {e}")
                skipped += 1

    print(f"Done. Rows: {total}, Created: {created}, Skipped: {skipped}")
    print(f"SVG folder: {svg_dir.resolve()}")
    print(f"JPG folder: {jpg_dir.resolve()}")

if __name__ == "__main__":
    main()