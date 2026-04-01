"""
Yearbook PDF Generator
======================
Reads student data from an Excel sheet and generates a PDF yearbook grid
suitable for upload to Canva.

CONFIGURATION — edit the values in this section to customise the output.
"""

# ── Third-party imports ────────────────────────────────────────────────────
import os
import re
import requests
from io import BytesIO

import pandas as pd
from PIL import Image
from pillow_heif import register_heif_opener
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle, Paragraph, Spacer, Image as RLImage
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas as rl_canvas  # ← replaces SimpleDocTemplate

register_heif_opener()

# ╔══════════════════════════════════════════════════════════════════════════╗
# ║                        CONFIGURATION BLOCK                              ║
# ╚══════════════════════════════════════════════════════════════════════════╝

# --- File paths ---
INPUT_EXCEL       = 'data.xlsx'         # Path to the input spreadsheet
DEFAULT_IMAGE     = 'default.png'       # Fallback image if a student has none
OUTPUT_PDF        = 'yearbook.pdf'      # Output PDF filename

# --- Column names (must match the header row in your Excel file exactly) ---
COL_NAME   = 'Name'
COL_ID     = 'BITS ID (this form is only for students enrolled in the year 2023)'
COL_QUOTE  = 'Submit a clean, creative yearbook quote (under 100 characters) to be printed under your image. Violation of the length limit may lead to omission of your quote.'
COL_PHOTO  = 'Upload a clear, well-lit, decent photo (1:1 ratio or passport size). Editing is not allowed, and you can only upload once. Ensure View Permissions are set to "Anyone with the Link"'

# --- Default fallback values ---
DEFAULT_QUOTE     = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit'
DEFAULT_PHOTO_URL = 'https://drive.google.com/file/d/null/view?usp=sharing'

# --- Grid layout ---
COLS_PER_PAGE = 3       # Number of columns in the grid
ROWS_PER_PAGE = 3       # Number of rows per page

# --- Image dimensions ---
# Aspect ratio expressed as width:height 3→  e.g. 1:1 = square, 3:4 = portrait
IMAGE_RATIO_W = 1       # Width part of the ratio
IMAGE_RATIO_H = 1       # Height part of the ratio

# Photo width in the PDF (in cm).  Height is calculated from the ratio above.
PHOTO_WIDTH_CM = 4.5

# --- Network ---
REQUEST_TIMEOUT = 15    # Seconds before a download is abandoned

# --- Font paths (Noto Sans — install with: sudo apt install fonts-noto) ---
# Update these paths if Noto Sans is installed somewhere else on your system.
FONT_REGULAR_PATH = '/usr/share/fonts/noto/NotoSans-Regular.ttf'
FONT_BOLD_PATH    = '/usr/share/fonts/noto/NotoSans-Bold.ttf'

# ╚══════════════════════════════════════════════════════════════════════════╝


# ── Derived constants (do not edit) ───────────────────────────────────────
STUDENTS_PER_PAGE = COLS_PER_PAGE * ROWS_PER_PAGE
PHOTO_RATIO       = IMAGE_RATIO_H / IMAGE_RATIO_W
PHOTO_WIDTH       = PHOTO_WIDTH_CM * cm
PHOTO_HEIGHT      = PHOTO_WIDTH * PHOTO_RATIO


# ── Font registration ─────────────────────────────────────────────────────
def _register_noto() -> tuple[str, str]:
    """
    Register NotoSans-Regular and NotoSans-Bold with ReportLab.
    Returns the font names to use for (regular, bold).
    Falls back to Helvetica with a clear error message if the files are missing.
    """
    reg_ok = bold_ok = False

    if os.path.exists(FONT_REGULAR_PATH):
        try:
            pdfmetrics.registerFont(TTFont('NotoSans', FONT_REGULAR_PATH))
            reg_ok = True
        except Exception as e:
            print(f"[WARN] Could not register NotoSans regular: {e}")
    else:
        print(f"[WARN] Font not found: {FONT_REGULAR_PATH}")

    if os.path.exists(FONT_BOLD_PATH):
        try:
            pdfmetrics.registerFont(TTFont('NotoSans-Bold', FONT_BOLD_PATH))
            bold_ok = True
        except Exception as e:
            print(f"[WARN] Could not register NotoSans bold: {e}")
    else:
        print(f"[WARN] Font not found: {FONT_BOLD_PATH}")

    if reg_ok and bold_ok:
        print("[INFO] Noto Sans loaded — full Unicode support enabled.")
        return 'NotoSans', 'NotoSans-Bold'
    else:
        print("[WARN] Falling back to Helvetica (Hindi/Chinese/emoji will not render correctly).")
        print("[WARN] Install Noto Sans and update FONT_REGULAR_PATH / FONT_BOLD_PATH above.")
        return 'Helvetica', 'Helvetica-Bold'

FONT_REGULAR, FONT_BOLD = _register_noto()


# ── Paragraph styles ──────────────────────────────────────────────────────
ID_STYLE = ParagraphStyle(
    'StudentID',
    fontSize=8,
    leading=10,
    alignment=TA_CENTER,
    textColor=colors.HexColor('#444444'),
    fontName=FONT_BOLD,
)
NAME_STYLE = ParagraphStyle(
    'StudentName',
    fontSize=9,
    leading=11,
    alignment=TA_CENTER,
    fontName=FONT_BOLD,
)
QUOTE_STYLE = ParagraphStyle(
    'Quote',
    fontSize=7,
    leading=9,
    alignment=TA_CENTER,
    textColor=colors.HexColor('#555555'),
    fontName=FONT_REGULAR,
)


# ── Helpers ───────────────────────────────────────────────────────────────

def extract_drive_file_id(url: str) -> str | None:
    m = re.search(r'/file/d/([a-zA-Z0-9_-]+)', url)
    if m:
        return m.group(1)
    m = re.search(r'[?&]id=([a-zA-Z0-9_-]+)', url)
    if m:
        return m.group(1)
    return None


def _try_open_image(data: bytes) -> Image.Image | None:
    try:
        img = Image.open(BytesIO(data))
        img.verify()
        img = Image.open(BytesIO(data))
        return img
    except Exception:
        return None


def download_drive_image(url: str) -> Image.Image | None:
    """
    Download an image from a Google Drive sharing link.
    Tries four URLs in order; returns None on total failure.
    """
    file_id = extract_drive_file_id(url)
    if not file_id:
        print(f"  [WARN] Could not parse file ID from URL: {url}")
        return None

    HEADERS = {
        'User-Agent': (
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
            'AppleWebKit/537.36 (KHTML, like Gecko) '
            'Chrome/122.0.0.0 Safari/537.36'
        )
    }
    candidates = [
        f'https://drive.usercontent.google.com/download?id={file_id}&export=view&authuser=0',
        f'https://drive.usercontent.google.com/download?id={file_id}&export=download&confirm=t',
        f'https://drive.google.com/uc?export=download&id={file_id}&confirm=t',
        f'https://drive.google.com/uc?export=download&id={file_id}',
    ]

    session = requests.Session()
    for attempt_url in candidates:
        try:
            response = session.get(
                attempt_url, headers=HEADERS,
                timeout=REQUEST_TIMEOUT, allow_redirects=True,
            )
            response.raise_for_status()
            if 'text/html' in response.headers.get('Content-Type', ''):
                continue
            img = _try_open_image(response.content)
            if img is not None:
                return img
            print(f"  [WARN] Received non-image bytes from {attempt_url[:60]}…")
        except requests.RequestException as e:
            print(f"  [WARN] Request error for file_id={file_id}: {e}")
            continue

    print(f"  [WARN] All download attempts failed for file_id={file_id}. Using default image.")
    return None


def crop_and_resize(img: Image.Image, ratio_w: int, ratio_h: int) -> Image.Image:
    target_ratio = ratio_w / ratio_h
    img_ratio = img.width / img.height
    if img_ratio > target_ratio:
        new_width = int(img.height * target_ratio)
        left = (img.width - new_width) // 2
        crop_box = (left, 0, left + new_width, img.height)
    else:
        new_height = int(img.width / target_ratio)
        top = (img.height - new_height) // 4
        crop_box = (0, top, img.width, top + new_height)
    return img.crop(crop_box)


def load_image(url: str) -> Image.Image | None:
    if not url or pd.isna(url):
        return None
    return download_drive_image(str(url).strip())


def image_to_rl(pil_img: Image.Image, col_width: float) -> Table | None:
    """Convert a PIL Image to a centred ReportLab flowable (1-cell wrapper Table)."""
    try:
        pil_img = pil_img.convert('RGB')
        buf = BytesIO()
        pil_img.save(buf, format='JPEG', quality=85)
        buf.seek(0)
        rl_img = RLImage(buf, width=PHOTO_WIDTH, height=PHOTO_HEIGHT)
        wrapper = Table([[rl_img]], colWidths=[col_width])
        wrapper.setStyle(TableStyle([
            ('ALIGN',         (0, 0), (0, 0), 'CENTER'),
            ('LEFTPADDING',   (0, 0), (0, 0), 0),
            ('RIGHTPADDING',  (0, 0), (0, 0), 0),
            ('TOPPADDING',    (0, 0), (0, 0), 0),
            ('BOTTOMPADDING', (0, 0), (0, 0), 0),
            ('INNERGRID',     (0, 0), (-1, -1), 0, colors.white),
            ('BOX',           (0, 0), (-1, -1), 0, colors.white),
        ]))
        wrapper.hAlign = 'CENTER'
        return wrapper
    except Exception as e:
        print(f"  [WARN] Could not convert image to ReportLab format: {e}")
        return None


def load_default_image(col_width: float) -> Table | None:
    if not os.path.exists(DEFAULT_IMAGE):
        print(f"[INFO] {DEFAULT_IMAGE} not found — creating grey placeholder.")
        w = int(300 * IMAGE_RATIO_W)
        h = int(300 * IMAGE_RATIO_H)
        Image.new('RGB', (w, h), color=(220, 220, 220)).save(DEFAULT_IMAGE)
    try:
        img = Image.open(DEFAULT_IMAGE)
        img = crop_and_resize(img, IMAGE_RATIO_W, IMAGE_RATIO_H)
        return image_to_rl(img, col_width)
    except Exception as e:
        print(f"  [WARN] Could not load default image: {e}")
        return None


def safe_str(value) -> str:
    """Coerce any pandas/Excel cell value to a plain Python str."""
    try:
        if pd.isna(value):
            return ''
    except (TypeError, ValueError):
        pass
    return str(value)


def build_cell(name: str, student_id: str, quote: str, photo_url: str,
               default_rl_img: Table | None, col_width: float) -> list:
    """Build a list of ReportLab flowables for one student cell: photo → ID → name → quote."""
    rl_img = None
    if photo_url and photo_url.strip() != DEFAULT_PHOTO_URL:
        pil_img = load_image(photo_url)
        if pil_img:
            pil_img = crop_and_resize(pil_img, IMAGE_RATIO_W, IMAGE_RATIO_H)
            rl_img = image_to_rl(pil_img, col_width)

    if rl_img is None:
        rl_img = default_rl_img

    cell = []
    if rl_img:
        cell.append(rl_img)
    cell.append(Spacer(1, 3))
    cell.append(Paragraph(student_id, ID_STYLE))
    cell.append(Paragraph(name,       NAME_STYLE))
    cell.append(Paragraph(quote,      QUOTE_STYLE))
    return cell


def build_page_table(page_students: list, page_start: int, total: int,
                     col_width: float, usable_w: float, usable_h: float,
                     default_rl_img: Table | None) -> tuple[Table, float, float]:
    """
    Build the grid Table for one page and return (table, rendered_w, rendered_h).

    Padding to a full grid is done here so the last page always has
    COLS_PER_PAGE columns even when students don't divide evenly.

    wrap() is called against the full usable area (not just col_width *
    COLS_PER_PAGE) so the reported height matches what will actually be
    drawn — preventing any off-by-a-few-points overflow.
    """
    padded = list(page_students)
    while len(padded) % COLS_PER_PAGE != 0:
        padded.append(None)

    table_data = []
    for row_idx in range(0, len(padded), COLS_PER_PAGE):
        row_students = padded[row_idx: row_idx + COLS_PER_PAGE]
        row_cells = []
        for i, s in enumerate(row_students):
            if s is None:
                row_cells.append('')
            else:
                name       = safe_str(s.get(COL_NAME, ''))
                student_id = safe_str(s.get(COL_ID, ''))
                quote      = safe_str(s.get(COL_QUOTE, DEFAULT_QUOTE))
                photo_url  = safe_str(s.get(COL_PHOTO, DEFAULT_PHOTO_URL))
                abs_idx    = page_start + row_idx + i + 1
                print(f"  [{abs_idx}/{total}] Processing {name} …")
                row_cells.append(
                    build_cell(name, student_id, quote, photo_url,
                               default_rl_img, col_width)
                )
        table_data.append(row_cells)

    tbl = Table(
        table_data,
        colWidths=[col_width] * COLS_PER_PAGE,
        rowHeights=None,
    )
    tbl.setStyle(TableStyle([
        ('VALIGN',        (0, 0), (-1, -1), 'TOP'),
        ('ALIGN',         (0, 0), (-1, -1), 'CENTER'),
        ('TOPPADDING',    (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('LEFTPADDING',   (0, 0), (-1, -1), 4),
        ('RIGHTPADDING',  (0, 0), (-1, -1), 4),
        ('INNERGRID',     (0, 0), (-1, -1), 0, colors.white),
        ('BOX',           (0, 0), (-1, -1), 0, colors.white),
    ]))

    tbl_w, tbl_h = tbl.wrap(usable_w, usable_h)
    return tbl, tbl_w, tbl_h


# ── Main ──────────────────────────────────────────────────────────────────

def main():
    # ── 1. Load data ──────────────────────────────────────────────────────
    print(f"[INFO] Reading {INPUT_EXCEL} …")
    df = pd.read_excel(INPUT_EXCEL)
    df.columns = df.columns.str.strip()

    required_cols = [COL_NAME, COL_ID, COL_QUOTE, COL_PHOTO]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        print("[ERROR] The following columns were not found in the spreadsheet:")
        for c in missing:
            print(f"  - '{c}'")
        print("\nAvailable columns:")
        for c in df.columns:
            print(f"  - '{c}'")
        return

    df[COL_QUOTE] = df[COL_QUOTE].fillna(DEFAULT_QUOTE)
    df[COL_PHOTO] = df[COL_PHOTO].fillna(DEFAULT_PHOTO_URL)
    print(f"[INFO] Loaded {len(df)} students.")

    # ── 2. Compute layout ─────────────────────────────────────────────────
    page_w, page_h = A4
    margin    = 1.5 * cm
    col_width = (page_w - 2 * margin) / COLS_PER_PAGE
    usable_w  = page_w - 2 * margin
    usable_h  = page_h - 2 * margin

    # ── 3. Pre-load default image once ────────────────────────────────────
    default_rl_img = load_default_image(col_width)

    # ── 4. Render each page directly onto a Canvas ────────────────────────
    #
    # Why canvas.Canvas instead of SimpleDocTemplate?
    #
    # SimpleDocTemplate is a *flowing* layout engine. Even with explicit
    # PageBreaks, if Spacer + Table height exceeds usable_h by even one point,
    # ReportLab spills the last row onto the next page.
    #
    # With canvas.Canvas the contract is simple:
    #   tbl.wrap()   — measure the table in isolation
    #   tbl.drawOn() — stamp it at an exact (x, y) coordinate
    #   c.showPage() — hard page commit; nothing can ever bleed past this line
    #
    # This guarantees exactly ROWS_PER_PAGE × COLS_PER_PAGE students per page
    # regardless of content height or font metrics.

    c        = rl_canvas.Canvas(OUTPUT_PDF, pagesize=A4)
    students = df.to_dict('records')
    total    = len(students)
    page_num = 0

    for page_start in range(0, total, STUDENTS_PER_PAGE):
        page_num += 1
        page_students = students[page_start: page_start + STUDENTS_PER_PAGE]
        print(f"\n[INFO] Building page {page_num} "
              f"(students {page_start + 1}–{min(page_start + STUDENTS_PER_PAGE, total)}) …")

        tbl, tbl_w, tbl_h = build_page_table(
            page_students, page_start, total,
            col_width, usable_w, usable_h, default_rl_img,
        )

        # Centre the table horizontally and vertically.
        # Canvas y=0 is the bottom of the page, so the bottom-left corner of a
        # vertically-centred block sits at  margin + (usable_h - tbl_h) / 2.
        x = margin + max(0.0, (usable_w - tbl_w) / 2)
        y = margin + max(0.0, (usable_h - tbl_h) / 2)

        tbl.drawOn(c, x, y)
        c.showPage()   # ← hard page boundary — nothing can overflow past here

    # ── 5. Save PDF ───────────────────────────────────────────────────────
    c.save()
    print(f"\n[DONE] {page_num} page(s) saved to {OUTPUT_PDF}")


if __name__ == '__main__':
    main()