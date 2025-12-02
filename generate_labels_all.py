# generate_labels_all.py — v2.12.5
import os, io, re, sys, glob
import fitz  # PyMuPDF
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.graphics.barcode import code39

def _force_utf8():
    try:
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        else:
            import io as _io
            sys.stdout = _io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
            sys.stderr = _io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
    except Exception:
        pass
_force_utf8()
os.environ["PYTHONUTF8"] = "1"
os.environ["PYTHONIOENCODING"] = "utf-8"

OUTPUT_DIR = "output_pdfs"
DEFAULT_SPLIT = 500
DEFAULT_DPI = 150

CROP_RED = dict(top=65.5, bottom=132, left=43, right=17)  # mm
REF_W, REF_H = 1155, 768
BOX_REF = (760, 585)
QTY_REF = (1030, 585)
LABEL_W_MM, LABEL_H_MM = 141, 97
MARGIN_MM = 4.5
RED_FONT_FILE = "Sansation_Regular.ttf"
RED_FONT_NAME = "SansationRegular"
RED_TEXT_SIZE_PT = 14
RED_SHIFT_MM = 5

BLUE_BORDER_PT   = 0.5
CODE_FONTSIZE_PT = 7
WEEK_FONTSIZE_PT = 7
BARSCALE_W_MM    = 44
BARSCALE_H_MM    = 7

def ensure_dir(path: str):
    os.makedirs(path, exist_ok=True)

def code_outdir(code: str) -> str:
    folder = os.path.join(OUTPUT_DIR, re.sub(r'[^A-Z0-9]', '', str(code).upper()))
    ensure_dir(folder)
    return folder

def sanitize_for_filename(s: str) -> str:
    s = str(s or '').strip()
    s = re.sub(r'\s+', '_', s)
    s = re.sub(r'[^A-Za-z0-9._-]+', '-', s)
    return s if s else 'LSX'

def render_full_page(page, dpi=DEFAULT_DPI):
    import PIL.Image as PilImage
    mat = fitz.Matrix(dpi/72, dpi/72)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    return PilImage.open(io.BytesIO(pix.tobytes('png'))).convert('RGB')

def crop_region(img, crop_conf, dpi):
    px_per_mm = dpi / 25.4
    w, h = img.size
    return img.crop((
        int(crop_conf['left']*px_per_mm),
        int(crop_conf['top']*px_per_mm),
        w - int(crop_conf['right']*px_per_mm),
        h - int(crop_conf['bottom']*px_per_mm)
    ))

def fmt_min2(n: int) -> str:
    return f"{n:02d}" if n < 100 else str(n)

def columnize_rows(start_n: int, end_n: int):
    total = max(0, end_n - start_n + 1)
    if total <= 0: return []
    base, rem = divmod(total, 4)
    sizes = [base + (1 if i < rem else 0) for i in range(4)]
    cols = []; cur = start_n
    for s in sizes:
        cols.append(list(range(cur, cur + s))); cur += s
    maxlen = max(len(c) for c in cols) if cols else 0
    rows = []
    for r in range(maxlen):
        rows.append([
            cols[0][r] if r < len(cols[0]) else None,
            cols[1][r] if r < len(cols[1]) else None,
            cols[2][r] if r < len(cols[2]) else None,
            cols[3][r] if r < len(cols[3]) else None,
        ])
    return rows

def try_register_font(ttf_path: str, name: str = 'SansationBold') -> str:
    try:
        if os.path.exists(ttf_path):
            pdfmetrics.registerFont(TTFont(name, ttf_path))
            return name
    except Exception:
        pass
    return 'Helvetica-Bold'

def draw_code39_on_canvas(c, value, x, y, cell_w, cell_h):
    from reportlab.lib.units import mm as _mm
    clean_val = re.sub(r'[^A-Z0-9]', '', str(value).upper())
    desired_w, desired_h = BARSCALE_W_MM*_mm, BARSCALE_H_MM*_mm
    bc = code39.Standard39(clean_val, stop=1, checksum=0, barHeight=desired_h, barWidth=0.25*_mm)
    scale_x = desired_w / bc.width
    bx = x + (cell_w - desired_w)/2.0
    by = y + (cell_h - desired_h)/2.0 - 5*_mm + 4*_mm
    c.saveState(); c.translate(bx, by); c.scale(scale_x, 1.0)
    bc.drawOn(c, 0, 0); c.restoreState()

def export_hangtag_generated(code_text, week_text, out_dir):
    ensure_dir(out_dir)
    out_path = os.path.join(out_dir, f"{re.sub(r'[^A-Z0-9]', '', str(code_text).upper())}_Hangtag.pdf")
    c = canvas.Canvas(out_path, pagesize=landscape(A4))

    page_w, page_h = landscape(A4)
    cell_w, cell_h = 40 * mm, 30 * mm
    spacing = 4 * mm
    total_w = 6 * cell_w + 5 * spacing
    total_h = 6 * cell_h + 5 * spacing
    offset_x = (page_w - total_w) / 2
    offset_y = (page_h - total_h) / 2

    font_name = try_register_font('Sansation_Bold.ttf', 'SansationBold')

    for r in range(6):
        for cidx in range(6):
            x = offset_x + cidx * (cell_w + spacing)
            y = offset_y + r * (cell_h + spacing)
            c.setLineWidth(BLUE_BORDER_PT)
            c.setStrokeColorRGB(0, 204/255, 1)
            c.rect(x, y, cell_w, cell_h)
            c.setFillColorRGB(0, 0, 0)
            c.setFont(font_name, CODE_FONTSIZE_PT)
            c.drawString(x + 3*mm, y + 4*mm, re.sub(r'[^A-Z0-9]', '', str(code_text).upper()))
            c.setFont(font_name, WEEK_FONTSIZE_PT)
            c.drawString(x + 3*mm, y + 1*mm, str(week_text))
            draw_code39_on_canvas(c, code_text, x, y, cell_w, cell_h)

    c.showPage(); c.save()
    print(f"[OK] Hangtag -> {os.path.basename(out_path)}")

def extract_week_text(doc, page_idx):
    try:
        mm2pt = 72/25.4
        rect = fitz.Rect(
            CROP_RED['left']*mm2pt,
            CROP_RED['top']*mm2pt,
            doc[page_idx].rect.width - CROP_RED['right']*mm2pt,
            doc[page_idx].rect.height - CROP_RED['bottom']*mm2pt
        )
        text = doc[page_idx].get_text('text', clip=rect)
        for ln in (ln.strip() for ln in text.splitlines() if ln.strip()):
            if 'MER-' in ln or re.search(r'W\d{1,2}', ln):
                return ln
    except Exception:
        pass
    return 'WEEK'

def export_chunk_colorlabel(code, qty_val, base_img, dpi,
                            start_n, end_n, total_display, out_suffix, out_dir):
    global RED_FONT_NAME
    try:
        pdfmetrics.getFont(RED_FONT_NAME)
    except:
        try:
            if os.path.exists(RED_FONT_FILE):
                pdfmetrics.registerFont(TTFont(RED_FONT_NAME, RED_FONT_FILE))
            else:
                RED_FONT_NAME = 'Helvetica'
        except:
            RED_FONT_NAME = 'Helvetica'

    ensure_dir(out_dir)
    buf_img = io.BytesIO()
    base_img.save(buf_img, format='JPEG', quality=85, optimize=True)
    buf_img.seek(0)
    img_reader = ImageReader(buf_img)

    page_w, page_h = landscape(A4)
    label_w_pt = LABEL_W_MM * mm
    label_h_pt = LABEL_H_MM * mm
    positions = [
        (MARGIN_MM*mm, page_h - MARGIN_MM*mm - label_h_pt),
        (page_w - MARGIN_MM*mm - label_w_pt, page_h - MARGIN_MM*mm - label_h_pt),
        (MARGIN_MM*mm, MARGIN_MM*mm),
        (page_w - MARGIN_MM*mm - label_w_pt, MARGIN_MM*mm),
    ]

    bx_norm, by_norm = BOX_REF[0]/REF_W, BOX_REF[1]/REF_H
    qx_norm, qy_norm = QTY_REF[0]/REF_W, QTY_REF[1]/REF_H

    ranged = f"{fmt_min2(start_n)}-{fmt_min2(end_n)}"
    out_name = f"{re.sub(r'[^A-Z0-9]', '', str(code).upper())}_ColorLabel_{ranged}_{out_suffix}.pdf"
    out_path = os.path.join(out_dir, out_name)
    c = canvas.Canvas(out_path, pagesize=landscape(A4))

    rows = columnize_rows(start_n, end_n)
    printed = 0
    for row in rows:
        for pos_idx, cur in enumerate(row):
            if cur is None: continue
            x, y = positions[pos_idx]
            c.drawImage(img_reader, x, y, width=label_w_pt, height=label_h_pt)
            c.setStrokeColorRGB(1, 0, 0); c.setLineWidth(1); c.rect(x, y, label_w_pt, label_h_pt, stroke=1, fill=0)
            c.setFillColorRGB(0, 0, 0); c.setFont(RED_FONT_NAME, RED_TEXT_SIZE_PT)
            bx_pt = x + bx_norm*label_w_pt - RED_SHIFT_MM*mm
            by_pt = y + (1 - by_norm)*label_h_pt
            qx_pt = x + qx_norm*label_w_pt - RED_SHIFT_MM*mm
            qy_pt = y + (1 - qy_norm)*label_h_pt
            c.drawCentredString(bx_pt, by_pt, f"{fmt_min2(cur)}/{total_display}")
            c.drawCentredString(qx_pt, qy_pt, f"{int(qty_val):02d}")
            printed += 1
        c.showPage()
    c.save()
    print(f"[OK] ColorLabel -> {os.path.basename(out_path)} ({printed} nhãn; {ranged}/{total_display})")

def to_int_safe(x):
    try:
        s = str(x).strip()
        if s == '' or s.lower() in ('nan','none','null'):
            return 0
        return int(float(s))
    except:
        return 0

def find_page_by_code(doc, code: str):
    for i, page in enumerate(doc):
        if code in page.get_text('text'):
            return i
    if code.startswith('C') and not code.startswith('CC'):
        alt = 'CC' + code[1:]
    elif code.startswith('CC'):
        alt = 'C' + code[2:]
    else:
        alt = 'CC' + code
    for i, page in enumerate(doc):
        if alt in page.get_text('text'):
            return i
    pattern = r''.join([re.escape(ch) + r'[\s\-]*' for ch in code])
    rx = re.compile(pattern)
    for i, page in enumerate(doc):
        if rx.search(page.get_text('text')):
            return i
    return None

def pick_first_existing(files, prefer_keywords=None):
    if not files: return None
    if prefer_keywords:
        for k in prefer_keywords:
            for f in files:
                if k.lower() in f.lower():
                    return f
    files_sorted = sorted(files, key=lambda p: os.path.getmtime(p), reverse=True)
    return files_sorted[0]

def resolve_args(argv):
    excel_file = None; pdf_file = None; dpi = DEFAULT_DPI; selected = 'all'
    args = argv[1:]
    positionals = [a for a in args if not a.startswith('--')]
    if (len(positionals) >= 2 and positionals[0].lower().endswith('.xlsx') and positionals[1].lower().endswith('.pdf')):
        excel_file, pdf_file = positionals[0], positionals[1]
        if len(positionals) >= 3:
            p3 = positionals[2].strip()
            try:
                dpi = int(p3)
                if len(positionals) >= 4:
                    selected = positionals[3]
            except ValueError:
                selected = p3
    else:
        if len(positionals) >= 1:
            try: dpi = int(positionals[0])
            except ValueError: pass
        if len(positionals) >= 2:
            selected = positionals[1]
        excel_file = os.environ.get('EXCEL_FILE')
        pdf_file   = os.environ.get('PDF_FILE')
        if not excel_file or not os.path.exists(excel_file):
            xlxs = glob.glob('*.xlsx'); excel_file = pick_first_existing(xlxs, prefer_keywords=['W','week'])
        if not pdf_file or not os.path.exists(pdf_file):
            pdfs = glob.glob('*.pdf'); pdf_file = pick_first_existing(pdfs, prefer_keywords=['label','labels'])
    if not excel_file or not os.path.exists(excel_file):
        print('Khong tim thay file Excel.'); sys.exit(1)
    if not pdf_file or not os.path.exists(pdf_file):
        print('Khong tim thay file PDF.'); sys.exit(1)
    return excel_file, pdf_file, dpi, selected, args

def parse_manual_range(args):
    mf = None; mt = None
    for a in args:
        if a.startswith('--manual-range='):
            s = a.split('=',1)[1].strip()
            if '-' in s:
                p1, p2 = s.split('-',1)
                try: mf = int(float(p1)); mt = int(float(p2))
                except: pass
        elif a.startswith('--manual-from='):
            try: mf = int(float(a.split('=',1)[1].strip()))
            except: pass
        elif a.startswith('--manual-to='):
            try: mt = int(float(a.split('=',1)[1].strip()))
            except: pass
    return mf, mt

def process_group(doc, df_group, export_mode, dpi, manual_range=None, mode_tag='default'):
    code = re.sub(r'[^A-Z0-9]', '', str(df_group.iloc[0]['code_norm']).upper())
    out_dir = code_outdir(code)
    page_idx = find_page_by_code(doc, code)
    if page_idx is None:
        print(f'BO QUA: Khong thay ma {code} trong PDF'); return
    page_img = render_full_page(doc[page_idx], dpi=dpi)
    base_img = crop_region(page_img, CROP_RED, dpi)

    if manual_range:
        row0 = df_group.iloc[0]
        qty_val  = to_int_safe(row0.get('qty_col_val', 0))
        sl_tong_val = to_int_safe(row0.get('sltong_val', 0))
        denom = (sl_tong_val // max(qty_val,1)) if (sl_tong_val > 0 and qty_val > 0) else 0
        pf, pt = manual_range
        if denom and pt > denom: pt = denom
        if pf < 1 or (denom and pf > denom) or pf > pt:
            print(f'BO QUA: Khoang khong hop le cho {code}: {pf}-{pt} / {denom or "N/A"}')
            return
        lsx_tag = sanitize_for_filename(row0.get('lsx_val', ''))
        export_chunk_colorlabel(code, qty_val, base_img, dpi,
                                start_n=int(pf), end_n=int(pt), total_display=int(denom or pt),
                                out_suffix=lsx_tag or 'LSX', out_dir=out_dir)
        if export_mode in ('hangtag','both'):
            wtxt = extract_week_text(doc, page_idx)
            export_hangtag_generated(code, wtxt, out_dir=out_dir)
        return

    if mode_tag == 'from_excel':
        cum_start = 1
        for _, row in df_group.iterrows():
            so_luong = to_int_safe(row.get('sl_col_val', 0))
            qty_val  = to_int_safe(row.get('qty_col_val', 0))
            labels_this_row = (so_luong // qty_val) if qty_val > 0 else 0
            if row.get('from_val') is not None and row.get('to_val') is not None:
                pf = int(row['from_val']); pt = int(row['to_val'])
            else:
                pf = cum_start; pt = cum_start + max(labels_this_row, 0) - 1
            sl_tong_val = to_int_safe(row.get('sltong_val', 0))
            denom = (sl_tong_val // max(qty_val,1)) if (sl_tong_val > 0 and qty_val > 0) else max(pt, 0)
            lsx_tag = sanitize_for_filename(row.get('lsx_val', ''))
            if (labels_this_row > 0 and pf <= pt):
                export_chunk_colorlabel(code, qty_val, base_img, dpi,
                                        start_n=int(pf), end_n=int(pt), total_display=int(denom),
                                        out_suffix=lsx_tag, out_dir=out_dir)
            if export_mode in ('hangtag','both'):
                wtxt = extract_week_text(doc, page_idx)
                export_hangtag_generated(code, wtxt, out_dir=out_dir)
            cum_start = int(max(pt, cum_start-1)) + 1
    else:
        for _, row in df_group.iterrows():
            so_luong = to_int_safe(row.get('sl_col_val', 0))
            qty_val  = to_int_safe(row.get('qty_col_val', 0))
            num_labels = (so_luong // qty_val) if qty_val > 0 else 0
            if num_labels <= 0: continue
            sl_tong_val = to_int_safe(row.get('sltong_val', 0))
            denom = (sl_tong_val // max(qty_val,1)) if (sl_tong_val > 0 and qty_val > 0) else num_labels
            cur = 1
            while cur <= num_labels:
                end = min(cur + DEFAULT_SPLIT - 1, num_labels)
                export_chunk_colorlabel(code, qty_val, base_img, dpi,
                                        start_n=cur, end_n=end, total_display=int(denom),
                                        out_suffix='LSX', out_dir=out_dir)
                cur = end + 1
        if export_mode in ('hangtag','both'):
            wtxt = extract_week_text(doc, page_idx)
            export_hangtag_generated(code, wtxt, out_dir=out_dir)

def main():
    EXCEL_FILE, PDF_FILE, dpi, selected, argv_full = resolve_args(sys.argv)
    export_mode = 'both'
    for a in argv_full:
        if a.startswith('--export='):
            v = a.split('=',1)[1].strip().lower()
            if v in ('color','hangtag','both'): export_mode = v

    manual_from, manual_to = parse_manual_range(argv_full)
    manual_range = (manual_from, manual_to) if (manual_from and manual_to) else None
    from_excel = ('--range-from-excel' in argv_full)

    print('='*54); print('  XUAT TEM NHAN (v2.12.5)'); print('='*54)
    print(f'Excel: {EXCEL_FILE}'); print(f'PDF  : {PDF_FILE}'); print(f'DPI  : {dpi}')
    print(f'Export: {export_mode}'); print(f'Selected: {selected}')
    if manual_range: print(f'Manual range: {manual_range[0]}-{manual_range[1]}')
    print('='*54)
    ensure_dir(OUTPUT_DIR)

    df = pd.read_excel(EXCEL_FILE, dtype={'QTY': str})

    def find_col(df, names):
        for n in names:
            if n in df.columns: return n
            lowmap = {c.lower(): c for c in df.columns}
            if n.lower() in lowmap: return lowmap[n.lower()]
        return None

    code_col = find_col(df, ['Mã SP đối tác','Ma SP doi tac','Code','Mã SP','MÃ SP ĐỐI TÁC','MaSp','MASP'])
    qty_col  = find_col(df, ['QTY','Qty','qty'])
    sl_col   = find_col(df, ['Số lượng','So luong','SO LUONG','SoLuong','SOLUONG'])
    sltong_col = find_col(df, ['SL Tổng','SL tổng','SL tong','Tổng SL','Tong SL','Total N','Total Labels','TongN','Tong N','SLTONG','SL_TONG'])
    lsx_col  = find_col(df, ['LSX','Lệnh SX','Lenh SX','LenhSX','LSX.'])
    from_col = find_col(df, ['From','Từ','Tu','Start','Bắt đầu','Bat dau','Khoảng từ','Khoang tu'])
    to_col   = find_col(df, ['To','Đến','Den','End','Kết thúc','Ket thuc','Khoảng đến','Khoang den'])

    if not code_col or not qty_col or not sl_col:
        print('Khong tim thay cac cot bat buoc (Code/QTY/Số lượng).'); return

    df['code_norm'] = df[code_col].astype(str).str.replace(r'^CC', 'C', regex=True)
    df['qty_col_val'] = df[qty_col]
    df['sl_col_val']  = df[sl_col]
    df['sltong_val']  = df[sltong_col] if sltong_col else 0
    df['lsx_val']     = df[lsx_col] if lsx_col else ''

    def _to_int_or_none(v):
        try: return int(float(str(v)))
        except: return None
    df['from_val'] = df[from_col].apply(_to_int_or_none) if from_col else None
    df['to_val']   = df[to_col].apply(_to_int_or_none) if to_col else None

    if selected != 'all':
        want = [re.sub(r'[^A-Z0-9]', '', s.upper()) for s in str(selected).split(',') if str(s).strip()]
        df = df[df['code_norm'].isin(want)]
        if df.empty:
            print(f'Khong tim thay ma: {selected}'); return

    doc = fitz.open(PDF_FILE)

    mode_tag = 'from_excel' if from_excel else 'default'
    if manual_range:
        for code, df_group in df.groupby('code_norm', sort=False):
            print(f'-- Manual range cho mã: {code}')
            process_group(doc, df_group, export_mode, DEFAULT_DPI, manual_range=manual_range, mode_tag=mode_tag)
        print('[OK] Xuất Tự điền khoảng số (mỗi mã 1 file).'); return

    for code, df_group in df.groupby('code_norm', sort=False):
        process_group(doc, df_group, export_mode, DEFAULT_DPI, manual_range=None, mode_tag=mode_tag)

    if from_excel:
        print('[OK] Xuất theo khoảng từ Excel (mỗi mã 1 thư mục, mỗi dòng 1 file, chạy nối tiếp).')
    else:
        print('[OK] Mặc định: tự tách file 500 tem (mỗi mã 1 thư mục).')

if __name__ == '__main__':
    main()
