import os
import re
import requests
import hashlib
from io import BytesIO

import pandas as pd
from PIL import Image
from pillow_heif import register_heif_opener

# Allow high-res photos (silences DecompressionBombWarning)
Image.MAX_IMAGE_PIXELS = None

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle, Paragraph, Spacer, Image as RLImage
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.fonts import addMapping

register_heif_opener()

INPUT_EXCEL       = 'data.xlsx'
DEFAULT_IMAGE     = 'default.png'
OUTPUT_PDF        = 'yearbook_4x4.pdf'

COL_NAME   = 'Name'
COL_ID     = 'BITS ID (this form is only for students enrolled in the year 2023)'
COL_QUOTE  = 'Submit a clean, creative yearbook quote (under 100 characters) to be printed under your image. Violation of the length limit may lead to omission of your quote.'
COL_PHOTO  = 'Upload a clear, well-lit, decent photo (1:1 ratio or passport size). Editing is not allowed, and you can only upload once. Ensure View Permissions are set to "Anyone with the Link"'

DEFAULT_QUOTE     = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit'
DEFAULT_PHOTO_URL = 'https://drive.google.com/file/d/null/view?usp=sharing'

COLS_PER_PAGE = 4
ROWS_PER_PAGE = 4
IMAGE_RATIO_W = 1
IMAGE_RATIO_H = 1
PHOTO_WIDTH_CM = 3.375
REQUEST_TIMEOUT = 15

FONT_REGULAR_PATH = 'assets/NotoSans-Regular.ttf'
FONT_BOLD_PATH    = 'assets/NotoSans-Bold.ttf'

STUDENTS_PER_PAGE = COLS_PER_PAGE * ROWS_PER_PAGE
PHOTO_RATIO       = IMAGE_RATIO_H / IMAGE_RATIO_W
PHOTO_WIDTH       = PHOTO_WIDTH_CM * cm
PHOTO_HEIGHT      = PHOTO_WIDTH * PHOTO_RATIO

# ── Dynamic Font Registry ───────────────────────────────────────────────────
def ensure_font(path, url):
    if not os.path.exists(path):
        os.makedirs(os.path.dirname(path), exist_ok=True)
        print(f"[INFO] Downloading font {path}...")
        try:
            r = requests.get(url, timeout=15)
            r.raise_for_status()
            with open(path, 'wb') as f:
                f.write(r.content)
        except Exception as e:
            print(f"[WARN] Failed to download {path}: {e}")

def _register_noto() -> tuple[str, str]:
    ensure_font(FONT_REGULAR_PATH, 'https://github.com/googlefonts/noto-fonts/raw/main/hinted/ttf/NotoSans/NotoSans-Regular.ttf')
    ensure_font(FONT_BOLD_PATH, 'https://github.com/googlefonts/noto-fonts/raw/main/hinted/ttf/NotoSans/NotoSans-Bold.ttf')
    
    global_font_configs = [
        ('NotoSansJP', 'assets/NotoSansJP.ttf', 'https://github.com/googlefonts/noto-cjk/raw/main/Sans/Variable/TTF/NotoSansCJKjp-VF.ttf'),
        ('NotoSansSC', 'assets/NotoSansSC.ttf', 'https://github.com/googlefonts/noto-cjk/raw/main/Sans/Variable/TTF/NotoSansCJKsc-VF.ttf'),
        ('NotoSansDevanagari', 'assets/NotoSansDevanagari.ttf', 'https://github.com/googlefonts/noto-fonts/raw/main/hinted/ttf/NotoSansDevanagari/NotoSansDevanagari-Regular.ttf'),
    ]
    
    try:
        pdfmetrics.registerFont(TTFont('NotoSans', FONT_REGULAR_PATH))
        pdfmetrics.registerFont(TTFont('NotoSans-Bold', FONT_BOLD_PATH))
    except Exception as e:
        print("[WARN] Base font registration failed:", e)
        return 'Helvetica', 'Helvetica-Bold'
        
    for font_name, font_path, url in global_font_configs:
        ensure_font(font_path, url)
        if os.path.exists(font_path):
            try:
                pdfmetrics.registerFont(TTFont(font_name, font_path))
                # Add explicit mapping to prevent ReportLab's "Can't map determine family/bold/italic" error
                addMapping(font_name, 0, 0, font_name) # Plain
                addMapping(font_name, 1, 0, font_name) # Bold (fallback to plain)
                addMapping(font_name, 0, 1, font_name) # Italic (fallback to plain)
                addMapping(font_name, 1, 1, font_name) # BoldItalic (fallback to plain)
                print(f"[INFO] Successfully registered CJK font: {font_name}")
            except Exception as e:
                print(f"[ERROR] Font registration failed for {font_name}: {e}")
                
    return 'NotoSans', 'NotoSans-Bold'

FONT_REGULAR, FONT_BOLD = _register_noto()

ID_STYLE = ParagraphStyle('StudentID', fontSize=7, leading=9, alignment=TA_CENTER, textColor=colors.HexColor('#444444'), fontName=FONT_BOLD)
NAME_STYLE = ParagraphStyle('StudentName', fontSize=8, leading=10, alignment=TA_CENTER, fontName=FONT_BOLD)
QUOTE_STYLE = ParagraphStyle('Quote', fontSize=6, leading=8, alignment=TA_CENTER, textColor=colors.HexColor('#555555'), fontName=FONT_REGULAR)

# ── Feature: Emoji & Unicode Processing ─────────────────────────────────────
def format_unicode_for_reportlab(text: str) -> str:
    text = text.replace('“', '"').replace('”', '"').replace("‘", "'").replace("’", "'")
    
    try:
        import emoji
        emojis_found = emoji.emoji_list(text)
    except ImportError:
        emojis_found = []
        
    result_chunks = []
    offset = 0
    for e_info in emojis_found:
        start, end, emj = e_info['match_start'], e_info['match_end'], e_info['emoji']
        if start > offset: result_chunks.append(text[offset:start])
        
        hex_code = '-'.join(f'{ord(c):x}' for c in emj if ord(c) != 0xfe0f)
        path = f"assets/emoji_{hex_code}.png"
        
        if not os.path.exists(path):
            os.makedirs(os.path.dirname(path), exist_ok=True)
            try:
                r = requests.get(f"https://cdn.jsdelivr.net/gh/twitter/twemoji@14.0.2/assets/72x72/{hex_code}.png", timeout=3)
                if r.status_code == 200:
                    with open(path, 'wb') as f: f.write(r.content)
            except Exception: pass
            
        if os.path.exists(path):
            result_chunks.append(f'<img src="{path}" width="7" height="7" valign="middle"/>')
            
        offset = end

    if offset < len(text): result_chunks.append(text[offset:])
    if not result_chunks:
        result_chunks = [text] # fallback if empty
        
    final_output = ""
    for chunk in result_chunks:
        if chunk.startswith('<img '):
            final_output += chunk
            continue
            
        current_font = None
        current_text = ""
        
        for char in chunk:
            code = ord(char)
            # Failsafe drop for emojis that slipped through
            if 0x1F600 <= code <= 0x1F64F or 0x1F300 <= code <= 0x1F5FF or 0x1F680 <= code <= 0x1F6FF or 0x2600 <= code <= 0x26FF or 0x2700 <= code <= 0x27BF: continue
            
            font_needed = None
            if 0x3040 <= code <= 0x309F or 0x30A0 <= code <= 0x30FF or 0x4E00 <= code <= 0x9FFF:
                font_needed = "NotoSansJP"
            elif 0x0900 <= code <= 0x097F:
                font_needed = "NotoSansDevanagari"
            
            if font_needed != current_font:
                if current_text:
                    if current_font: final_output += f'<font name="{current_font}">{current_text}</font>'
                    else: final_output += current_text
                current_font = font_needed
                current_text = ""
            
            if char == "<": current_text += "&lt;"
            elif char == ">": current_text += "&gt;"
            elif char == "&": current_text += "&amp;"
            else: current_text += char
            
        if current_text:
            if current_font: final_output += f'<font name="{current_font}">{current_text}</font>'
            else: final_output += current_text
                
    return final_output

# ── Feature: Borders and Watermark ──────────────────────────────────────────

def draw_background(c: rl_canvas.Canvas, page_w: float, page_h: float):
    c.saveState()
    
    margin = 0.8 * cm
    c.setLineWidth(2.0)
    # Top border - Yellow
    c.setStrokeColor(colors.HexColor('#F2C047'))
    c.line(margin, page_h - margin, page_w - margin, page_h - margin)
    # Bottom border - Red
    c.setStrokeColor(colors.HexColor('#C84433'))
    c.line(margin, margin, page_w - margin, margin)
    # Left & Right borders - Light Blue
    c.setStrokeColor(colors.HexColor('#60ABC9'))
    c.line(margin, margin, margin, page_h - margin)
    c.line(page_w - margin, margin, page_w - margin, page_h - margin)
    
    if os.path.exists('emblem.png'):
        try:
            from reportlab.lib.utils import ImageReader
            # Open as RGBA and extract original alpha
            original = Image.open('emblem.png').convert('RGBA')
            alpha = original.getchannel('A')
            # Convert to grayscale but keep it as RGBA for merging
            gray_rgb = original.convert('L').convert('RGBA')
            r, g, b, _ = gray_rgb.split()
            
            # Dim the original alpha to 12%
            dimmed_alpha = alpha.point(lambda p: int(p * 0.12))
            emblem_img = Image.merge('RGBA', (r, g, b, dimmed_alpha))
            
            # Larger centered watermark (75% of page width)
            wm_width = page_w * 0.75
            wm_height = (emblem_img.height / emblem_img.width) * wm_width
            x_pos = (page_w - wm_width) / 2
            y_pos = (page_h - wm_height) / 2
            c.drawImage(ImageReader(emblem_img), x_pos, y_pos, width=wm_width, height=wm_height, mask='auto')
        except Exception as e:
            pass
    c.restoreState()

# ── Core Image Logic ────────────────────────────────────────────────────────
def extract_drive_file_id(url: str):
    m = re.search(r'/file/d/([a-zA-Z0-9_-]+)', url)
    if m: return m.group(1)
    m = re.search(r'[?&]id=([a-zA-Z0-9_-]+)', url)
    if m: return m.group(1)
    return None

def _try_open_image(data: bytes):
    try:
        img = Image.open(BytesIO(data))
        img.verify()
        return Image.open(BytesIO(data))
    except Exception:
        return None

def download_drive_image(url: str):
    file_id = extract_drive_file_id(url)
    if not file_id: return None
    HEADERS = {'User-Agent': 'Mozilla/5.0 Chrome/122.0.0.0 Safari/537.36'}
    candidates = [
        f'https://drive.usercontent.google.com/download?id={file_id}&export=view&authuser=0',
        f'https://drive.usercontent.google.com/download?id={file_id}&export=download&confirm=t',
        f'https://drive.google.com/uc?export=download&id={file_id}&confirm=t',
        f'https://drive.google.com/uc?export=download&id={file_id}',
    ]
    session = requests.Session()
    for attempt_url in candidates:
        try:
            response = session.get(attempt_url, headers=HEADERS, timeout=REQUEST_TIMEOUT, allow_redirects=True)
            response.raise_for_status()
            if 'text/html' in response.headers.get('Content-Type', ''): continue
            img = _try_open_image(response.content)
            if img: return img
        except requests.RequestException:
            continue
    return None

def crop_and_resize(img: Image.Image, ratio_w: int, ratio_h: int):
    target_ratio = ratio_w / ratio_h
    img_ratio = img.width / img.height
    if img_ratio > target_ratio:
        new_width = int(img.height * target_ratio)
        left = (img.width - new_width) // 2
        return img.crop((left, 0, left + new_width, img.height))
    else:
        new_height = int(img.width / target_ratio)
        top = (img.height - new_height) // 4
        return img.crop((0, top, img.width, top + new_height))

def load_image(url: str):
    if not url or pd.isna(url): return None
    url_str = str(url).strip()
    if not url_str or "drive.google.com" not in url_str: return None
    
    # NEW CACHING LOGIC
    os.makedirs('photo_cache', exist_ok=True)
    url_hash = hashlib.md5(url_str.encode()).hexdigest()
    cache_path = os.path.join('photo_cache', f"{url_hash}.jpg")
    
    if os.path.exists(cache_path):
        try: return Image.open(cache_path)
        except Exception: pass
        
    img = download_drive_image(url_str)
    if img:
        try:
            # Save a compressed version to cache
            img.convert('RGB').save(cache_path, quality=85)
        except Exception: pass
    return img

def image_to_rl(pil_img: Image.Image, col_width: float):
    try:
        pil_img = pil_img.convert('RGB')
        buf = BytesIO()
        pil_img.save(buf, format='JPEG', quality=85)
        buf.seek(0)
        wrapper = Table([[RLImage(buf, width=PHOTO_WIDTH, height=PHOTO_HEIGHT)]], colWidths=[col_width])
        wrapper.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER')]))
        return wrapper
    except Exception: return None

def load_default_image():
    if not os.path.exists(DEFAULT_IMAGE):
        Image.new('RGB', (int(300 * IMAGE_RATIO_W), int(300 * IMAGE_RATIO_H)), color=(220, 220, 220)).save(DEFAULT_IMAGE)
    try: 
        return crop_and_resize(Image.open(DEFAULT_IMAGE), IMAGE_RATIO_W, IMAGE_RATIO_H)
    except Exception: return None

def build_cell(name, student_id, quote, photo_url, default_pil_img, col_width):
    rl_img = None
    # Smarter detection of missing/invalid photo submissions
    is_placeholder = (
        not photo_url or 
        pd.isna(photo_url) or 
        "null" in str(photo_url).lower() or 
        str(photo_url).strip() == "" or
        "drive.google.com" not in str(photo_url)
    )
    
    if not is_placeholder:
        pil_img = load_image(photo_url)
        if pil_img: 
            rl_img = image_to_rl(crop_and_resize(pil_img, IMAGE_RATIO_W, IMAGE_RATIO_H), col_width)
    
    if not rl_img and default_pil_img: 
        # Create a fresh ReportLab instance for the default image to allow reuse
        rl_img = image_to_rl(default_pil_img, col_width)
    
    quote = format_unicode_for_reportlab(quote)
    cell = []
    if rl_img: cell.append(rl_img)
    cell.append(Spacer(1, 3))
    cell.append(Paragraph(student_id, ID_STYLE))
    cell.append(Paragraph(name, NAME_STYLE))
    cell.append(Paragraph(quote, QUOTE_STYLE))
    return cell

def safe_str(val): return '' if pd.isna(val) else str(val)

def build_page_table(page_students, page_start, total, col_width, usable_w, usable_h, default_pil_img):
    padded = list(page_students)
    while len(padded) % COLS_PER_PAGE != 0: padded.append(None)

    table_data = []
    for row_idx in range(0, len(padded), COLS_PER_PAGE):
        row_students = padded[row_idx: row_idx + COLS_PER_PAGE]
        row_cells = []
        for i, s in enumerate(row_students):
            if s is None: row_cells.append('')
            else:
                name, student_id = safe_str(s.get(COL_NAME, '')), safe_str(s.get(COL_ID, ''))
                quote, photo_url = safe_str(s.get(COL_QUOTE, DEFAULT_QUOTE)), safe_str(s.get(COL_PHOTO, DEFAULT_PHOTO_URL))
                print(f"  [{page_start + row_idx + i + 1}/{total}] Processing {name} …")
                row_cells.append(build_cell(name, student_id, quote, photo_url, default_pil_img, col_width))
        table_data.append(row_cells)

    tbl = Table(table_data, colWidths=[col_width] * COLS_PER_PAGE)
    tbl.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP'), ('ALIGN', (0,0), (-1,-1), 'CENTER')]))
    tbl_w, tbl_h = tbl.wrap(usable_w, usable_h)
    return tbl, tbl_w, tbl_h

def render_page(c, page_students, page_start, total, page_w, page_h, margin, col_width, usable_w, usable_h, default_pil_img):
    draw_background(c, page_w, page_h)
    tbl, tbl_w, tbl_h = build_page_table(page_students, page_start, total, col_width, usable_w, usable_h, default_pil_img)
    x_pos = margin + max(0.0, (usable_w - tbl_w) / 2)
    y_pos = margin + usable_h - tbl_h
    tbl.drawOn(c, x_pos, y_pos)
    c.showPage()

def main():
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument('--update', type=str, help='Update only this student ID (e.g. 2023xxxxG)')
    parser.add_argument('--sample', action='store_true', help='Run using sample_data.xlsx and generate _sample PDF')
    args = parser.parse_args()

    global INPUT_EXCEL, OUTPUT_PDF
    if args.sample:
        INPUT_EXCEL = 'sample_data.xlsx'
        OUTPUT_PDF = OUTPUT_PDF.replace('.pdf', '_sample.pdf')
        print(f"[SAMPLE MODE] Using {INPUT_EXCEL} and outputting to {OUTPUT_PDF}")

    df = pd.read_excel(INPUT_EXCEL, engine='openpyxl')
    df.columns = df.columns.str.strip()
    df[COL_QUOTE] = df[COL_QUOTE].fillna(DEFAULT_QUOTE)
    df[COL_PHOTO] = df[COL_PHOTO].fillna(DEFAULT_PHOTO_URL)
    students = df.to_dict('records')

    def get_sort_key(s):
        uid = str(s.get(COL_ID, '')).strip()
        if len(uid) < 12: return uid
        # 202n AB CD xxxx G
        # 0123 45 67 8901 2
        n = uid[3]
        ab = uid[4:6]
        xxxx = uid[8:12]
        return (n, ab, xxxx)

    students.sort(key=get_sort_key)
    total = len(students)

    page_w, page_h = A4
    margin = 1.0 * cm
    col_width = (page_w - 2 * margin) / COLS_PER_PAGE
    usable_w, usable_h = page_w - 2 * margin, page_h - 2 * margin
    default_pil_img = load_default_image()

    if args.update:
        from pypdf import PdfReader, PdfWriter
        target_id = args.update.strip()
        student_idx = next((i for i, s in enumerate(students) if str(s.get(COL_ID, '')).strip() == target_id), None)
        if student_idx is None:
            print(f"[ERROR] Student ID {target_id} not found.")
            return

        target_page_idx = student_idx // STUDENTS_PER_PAGE
        page_start = target_page_idx * STUDENTS_PER_PAGE
        page_students = students[page_start: page_start + STUDENTS_PER_PAGE]

        tmp_pdf = "tmp_single_page.pdf"
        c_tmp = rl_canvas.Canvas(tmp_pdf, pagesize=A4)
        print(f"\n[INFO] Updating page {target_page_idx + 1} for student {target_id} ...")
        render_page(c_tmp, page_students, page_start, total, page_w, page_h, margin, col_width, usable_w, usable_h, default_pil_img)
        c_tmp.save()

        if not os.path.exists(OUTPUT_PDF):
            os.rename(tmp_pdf, OUTPUT_PDF)
            print(f"[INFO] {OUTPUT_PDF} created with updated page.")
        else:
            reader = PdfReader(OUTPUT_PDF)
            writer = PdfWriter()
            new_page_reader = PdfReader(tmp_pdf)
            new_page = new_page_reader.pages[0]

            for i in range(len(reader.pages)):
                if i == target_page_idx: writer.add_page(new_page)
                else: writer.add_page(reader.pages[i])

            with open(OUTPUT_PDF, "wb") as f: writer.write(f)
            os.remove(tmp_pdf)
            print(f"\n[DONE] Page {target_page_idx + 1} updated in {OUTPUT_PDF}")
    else:
        c = rl_canvas.Canvas(OUTPUT_PDF, pagesize=A4)
        page_num = 0
        for page_start in range(0, total, STUDENTS_PER_PAGE):
            page_num += 1
            print(f"\n[INFO] Building page {page_num} ...")
            render_page(c, students[page_start: page_start + STUDENTS_PER_PAGE], page_start, total, page_w, page_h, margin, col_width, usable_w, usable_h, default_pil_img)
        c.save()
        print(f"\n[DONE] {page_num} page(s) saved to {OUTPUT_PDF}")

if __name__ == '__main__':
    main()
