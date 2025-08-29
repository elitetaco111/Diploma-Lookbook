from __future__ import annotations

import csv
import re
from collections import OrderedDict, defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, landscape

# TODO
# Add style number
# Potential for QR Code

# Config
CSV_FILE = Path("DRNItemDiplomaFramesResults569.csv")
IMAGES_DIR = Path("diploma_images")
OUTPUT_DIR = Path("lookbooks")

# Page/layout settings (mm)
LEFT_MARGIN = 15
RIGHT_MARGIN = 15
TOP_MARGIN = 12
BOTTOM_MARGIN = 12
FOOTER_H = 16  # space reserved at bottom for text under the image
FONT_NAME = "Helvetica"
FONT_SIZE = 18

# Convert mm to points for ReportLab
MM_TO_PT = 72.0 / 25.4
LEFT_MARGIN_PT = LEFT_MARGIN * MM_TO_PT
RIGHT_MARGIN_PT = RIGHT_MARGIN * MM_TO_PT
TOP_MARGIN_PT = TOP_MARGIN * MM_TO_PT
BOTTOM_MARGIN_PT = BOTTOM_MARGIN * MM_TO_PT
FOOTER_H_PT = FOOTER_H * MM_TO_PT

PAGE_SIZE = landscape(letter)  # Landscape US Letter

# Column names to match (case/punctuation-insensitive)
CANDIDATES = {
    "team": [
        "team league data - lookbook",
        "team league data lookbook",
        "team league data (lookbook)",
        "team league data",
        "lookbook team",
        "team lookbook",
    ],
    "name": ["name", "product name", "sku name"],
    # Use only "Store Display Name"
    "display_name": ["store display name"],
    "price": ["original price", "price", "msrp", "list price"],
}

def normalize(s: str) -> str:
    return re.sub(r"[\W_]+", " ", (s or "").strip().lower())

def find_column(headers: List[str], keys: List[str], contains_all: Optional[List[str]] = None) -> Optional[str]:
    norm_headers = {normalize(h): h for h in headers}
    # Try exact matches first
    for key in keys:
        nk = normalize(key)
        if nk in norm_headers:
            return norm_headers[nk]
    # Try contains-all heuristic
    if contains_all:
        tokens = [normalize(t) for t in contains_all]
        for nh, original in norm_headers.items():
            if all(t in nh for t in tokens):
                return original
    return None


def resolve_columns(headers: List[str]) -> Dict[str, str]:
    team_col = find_column(headers, CANDIDATES["team"], contains_all=["team", "league", "lookbook"])
    name_col = find_column(headers, CANDIDATES["name"])
    display_name_col = find_column(headers, CANDIDATES["display_name"])
    price_col = find_column(headers, CANDIDATES["price"])

    missing = []
    if not team_col:
        missing.append("Team League Data - lookbook")
    if not name_col:
        missing.append("Name")
    if not display_name_col:
        missing.append("Store Display Name")
    if not price_col:
        missing.append("Original Price")

    if missing:
        raise ValueError(
            "Missing required columns in CSV: "
            + ", ".join(missing)
            + f". Found headers: {headers}"
        )

    return {
        "team": team_col,
        "name": name_col,
        "display_name": display_name_col,
        "price": price_col,
    }


def sanitize_filename(name: str) -> str:
    name = name.strip() or "untitled"
    # Replace invalid Windows filename chars
    name = re.sub(r'[<>:"/\\|?*]+', "_", name)
    # Strip trailing dots/spaces
    return name.rstrip(" .")[:150]


@dataclass
class ItemRow:
    team: str
    name: str
    display_name: str
    price: str


def load_rows(csv_path: Path) -> Tuple[List[ItemRow], Dict[str, str]]:
    if not csv_path.exists():
        raise FileNotFoundError(f"CSV not found: {csv_path}")

    with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        if not reader.fieldnames:
            raise ValueError("CSV has no headers/columns.")
        cols = resolve_columns(reader.fieldnames)

        items: List[ItemRow] = []
        for row in reader:
            team = (row.get(cols["team"]) or "").strip()
            name = (row.get(cols["name"]) or "").strip()
            display_name = (row.get(cols["display_name"]) or "").strip()
            price = (row.get(cols["price"]) or "").strip()

            # Skip rows missing essential fields
            if not team or not name:
                continue

            items.append(ItemRow(team=team, name=name, display_name=display_name, price=price))

    return items, cols


def add_item_page(c: canvas.Canvas, image_path: Path, display_name: str, price: str) -> None:
    page_w, page_h = PAGE_SIZE
    # Content area
    epw = page_w - LEFT_MARGIN_PT - RIGHT_MARGIN_PT
    eph = page_h - TOP_MARGIN_PT - BOTTOM_MARGIN_PT
    max_img_w = max(10.0, epw)
    max_img_h = max(10.0, eph - FOOTER_H_PT)

    # Compute scaled size preserving aspect ratio
    try:
        with Image.open(image_path) as im:
            w_px, h_px = im.size
    except Exception:
        w_px, h_px = 1000, 750

    ratio = w_px / h_px if h_px else 1.3333
    target_w = max_img_w
    target_h = target_w / ratio
    if target_h > max_img_h:
        target_h = max_img_h
        target_w = target_h * ratio

    # Position image centered in the image area (above the footer)
    image_area_h = eph - FOOTER_H_PT
    x = LEFT_MARGIN_PT + (epw - target_w) / 2
    y = BOTTOM_MARGIN_PT + FOOTER_H_PT + (image_area_h - target_h) / 2

    try:
        c.drawImage(str(image_path), x, y, width=target_w, height=target_h, preserveAspectRatio=False, anchor='sw')
    except Exception as e:
        print(f"[WARN] Failed to embed image '{image_path.name}': {e}")

    # Footer text (centered)
    c.setFont(FONT_NAME, FONT_SIZE)
    line_h = FONT_SIZE + 2
    footer_bottom = BOTTOM_MARGIN_PT
    start_y = footer_bottom + (FOOTER_H_PT - 2 * line_h) / 2
    cx = page_w / 2.0
    c.drawCentredString(cx, start_y + line_h, f"Store Display Name: {display_name or ''}")
    c.drawCentredString(cx, start_y, f"Original Price: {price or ''}")

    c.showPage()


def generate_lookbooks(items: List[ItemRow]) -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # Preserve first-seen order of teams
    team_order = OrderedDict()
    grouped: Dict[str, List[ItemRow]] = defaultdict(list)
    for it in items:
        if it.team not in team_order:
            team_order[it.team] = None
        grouped[it.team].append(it)

    missing_images_count = 0
    total_pages = 0

    for team in team_order.keys():
        rows = grouped[team]
        if not rows:
            continue

        # Filter to rows that have existing images first (to avoid empty PDFs)
        valid_rows: List[Tuple[ItemRow, Path]] = []
        for row in rows:
            image_path = IMAGES_DIR / f"{row.name}.jpg"
            if not image_path.exists():
                print(f"[WARN] Missing image: {image_path}")
                missing_images_count += 1
                continue
            valid_rows.append((row, image_path))

        if not valid_rows:
            print(f"[INFO] No pages for team '{team}' (no images found). Skipped PDF.")
            continue

        out_name = sanitize_filename(team) + ".pdf"
        out_path = OUTPUT_DIR / out_name
        c = canvas.Canvas(str(out_path), pagesize=PAGE_SIZE)

        for row, image_path in valid_rows:
            add_item_page(c, image_path, row.display_name, row.price)

        c.save()
        added_pages = len(valid_rows)
        total_pages += added_pages
        print(f"[OK] Wrote {added_pages} page(s): {out_path}")

    print(f"[DONE] Total pages: {total_pages}. Missing images: {missing_images_count}. Output dir: {OUTPUT_DIR.resolve()}")


def main() -> None:
    items, _ = load_rows(CSV_FILE)
    if not items:
        print("[INFO] No items found. Nothing to do.")
        return
    generate_lookbooks(items)


if __name__ == "__main__":
    main()