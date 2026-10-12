"""
gui_replace_links.py
====================
Giao diện đồ họa (GUI) cho công cụ gắn hyperlink vào Word.

Chạy:
    py gui_replace_links.py
"""

import re
import copy
import threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ── Lazy-loaded heavy libs (import trong background để UI hiện ngay) ─────────
openpyxl   = None
Document   = None
qn         = None
OxmlElement = None
_libs_ready = threading.Event()   # set() khi import xong


def _preload_libs():
    """Import các thư viện nặng trong background thread."""
    global openpyxl, Document, qn, OxmlElement
    import openpyxl as _openpyxl
    from docx import Document as _Document
    from docx.oxml.ns import qn as _qn
    from docx.oxml import OxmlElement as _OxmlElement
    openpyxl    = _openpyxl
    Document    = _Document
    qn          = _qn
    OxmlElement = _OxmlElement
    _libs_ready.set()

# ══════════════════════════════════════════════════════════
# ── CORE LOGIC ────────────────────────────────────────────
# ══════════════════════════════════════════════════════════
CODE_PATTERN = re.compile(r'\[H\d{2}\.\d{2}\.\d{2}\.\d{2}\]')
# Phat hien ma loi dinh dang: co khoang trang ben trong, vd [H01.01.01. 07]
CODE_FUZZY   = re.compile(r'\[H[\d\.\s]{10,18}\]')


def load_url_map(excel_path):
    # Col 0: Ma MC (H01.01.01.01...H05.05.01.08)
    # Col 1: Ten MC
    # Col 2: Link URL
    # Dieu kien: du 3 cot, col0 la ma H, col2 la URL http
    wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    url_map = {}
    CODE_RE = re.compile(r'H\d{2}\.\d{2}\.\d{2}\.\d{2}')
    for sheet in wb.worksheets:
        for row in sheet.iter_rows(values_only=True):
            if len(row) < 3:
                continue
            ma  = str(row[0]).strip() if row[0] else ''
            url = str(row[2]).strip() if row[2] else ''
            if CODE_RE.match(ma) and url.startswith('http'):
                url_map[ma] = url
    wb.close()
    return url_map


def _collect_segments(paragraph):
    # Thu thap (text, rPr_copy, old_url) tu <w:r> va <w:hyperlink> cu
    # old_url = URL cu neu segment nam trong hyperlink, None neu khong
    p_elem = paragraph._p
    segs = []
    for child in p_elem:
        if child.tag == qn('w:r'):
            t    = child.find(qn('w:t'))
            text = (t.text or '') if t is not None else ''
            rPr  = child.find(qn('w:rPr'))
            segs.append((text, copy.deepcopy(rPr) if rPr is not None else None, None))
        elif child.tag == qn('w:hyperlink'):
            # Lay URL cu tu relationship
            r_id    = child.get(qn('r:id'))
            old_url = None
            if r_id:
                try:
                    old_url = paragraph.part.rels[r_id].target_ref
                except (KeyError, AttributeError):
                    old_url = None
            for r in child.findall(qn('w:r')):
                t    = r.find(qn('w:t'))
                text = (t.text or '') if t is not None else ''
                rPr  = r.find(qn('w:rPr'))
                segs.append((text, copy.deepcopy(rPr) if rPr is not None else None, old_url))
    return segs


def _make_run(text, rPr):
    # Tao <w:r> giu nguyen dinh dang goc
    r = OxmlElement('w:r')
    if rPr is not None:
        r.append(copy.deepcopy(rPr))
    t = OxmlElement('w:t')
    t.text = text
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    r.append(t)
    return r


def _emit_range(p_elem, full_text, start, end, pos_map, segs):
    # Xuat full_text[start:end] thanh <w:r> giu dinh dang tung segment
    if start >= end:
        return
    i = start
    while i < end:
        seg_idx = pos_map[i] if i < len(pos_map) else 0
        j = i + 1
        while j < end and j < len(pos_map) and pos_map[j] == seg_idx:
            j += 1
        chunk  = full_text[i:j]
        _, rPr, _u = segs[seg_idx] if seg_idx < len(segs) else ('', None, None)
        p_elem.append(_make_run(chunk, rPr))
        i = j


def _build_link_rPr(context_rPr, link_style: str = 'auto', link_color: str = '0000FF'):
    """
    Xay dung rPr cho hyperlink.
    link_style : 'auto' | 'bold' | 'normal'
    link_color : ma hex 6 ky tu (khong co #), vi du '0000FF'
    """
    new_rPr = copy.deepcopy(context_rPr) if context_rPr is not None \
              else OxmlElement('w:rPr')
    for tag in ('w:rStyle', 'w:b', 'w:bCs', 'w:color', 'w:u'):
        el = new_rPr.find(qn(tag))
        if el is not None:
            new_rPr.remove(el)
    rStyle = OxmlElement('w:rStyle')
    rStyle.set(qn('w:val'), 'Hyperlink')
    new_rPr.insert(0, rStyle)
    orig_bold = _is_bold(context_rPr) if context_rPr is not None else False
    apply_bold = (
        orig_bold if link_style == 'auto' else
        True      if link_style == 'bold' else
        False
    )
    if apply_bold:
        new_rPr.append(OxmlElement('w:b'))
        new_rPr.append(OxmlElement('w:bCs'))
    color = OxmlElement('w:color')
    color.set(qn('w:val'), link_color.upper().lstrip('#') or '0000FF')
    new_rPr.append(color)
    return new_rPr


def _is_bold(rPr) -> bool:
    """Kiem tra xem rPr co in dam (w:b) khong."""
    if rPr is None:
        return False
    b = rPr.find(qn('w:b'))
    if b is None:
        return False
    val = b.get(qn('w:val'), '')
    return val not in ('0', 'false')  # w:b khong co val = dam


def _page_breaks_in(p_elem) -> int:
    """Dem so page break trong mot paragraph (ca lastRenderedPageBreak va br type=page)."""
    count = 0
    for el in p_elem.iter():
        if el.tag == qn('w:lastRenderedPageBreak'):
            count += 1
        elif el.tag == qn('w:br') and el.get(qn('w:type')) == 'page':
            count += 1
    return count


def process_paragraph(paragraph, url_map, page_num=1,
                      link_style='auto', link_color='0000FF'):
    # Van ban thuong: giu 100% dinh dang goc
    # Ma [Hxx] co URL -> hyperlink mau link_color, bold theo link_style
    # Ma khong co URL -> giu nguyen text + dinh dang goc
    p_elem = paragraph._p
    segs   = _collect_segments(paragraph)   # (text, rPr, old_url)
    full_text = ''.join(s[0] for s in segs)

    # Dem page break TRUOC khi xoa children
    breaks = _page_breaks_in(p_elem)

    matches = list(CODE_PATTERN.finditer(full_text))
    if not matches:
        return 0, [], breaks
    pos_map = []
    for seg_idx, (text, _, _u) in enumerate(segs):
        pos_map.extend([seg_idx] * len(text))
    replaced = 0
    changes  = []  # [(code, old_url, new_url, page, is_bold)]
    for child in list(p_elem):
        if child.tag != qn('w:pPr'):
            p_elem.remove(child)
    cursor = 0
    for m in matches:
        start, end        = m.start(), m.end()
        code_with_bracket = m.group()
        code_clean        = code_with_bracket[1:-1]
        _emit_range(p_elem, full_text, cursor, start, pos_map, segs)
        if code_clean in url_map:
            ctx_seg  = pos_map[start] if start < len(pos_map) else 0
            ctx_rPr  = segs[ctx_seg][1] if ctx_seg < len(segs) else None
            old_url  = segs[ctx_seg][2] if ctx_seg < len(segs) else None
            bold_before = _is_bold(ctx_rPr)
            part = paragraph.part
            new_url = url_map[code_clean]
            r_id = part.relate_to(
                new_url,
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
                is_external=True
            )
            hl = OxmlElement('w:hyperlink')
            hl.set(qn('r:id'), r_id)
            hl_run = OxmlElement('w:r')
            hl_run.append(_build_link_rPr(ctx_rPr, link_style, link_color))
            hl_t = OxmlElement('w:t')
            hl_t.text = code_with_bracket
            hl_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            hl_run.append(hl_t)
            hl.append(hl_run)
            p_elem.append(hl)
            replaced += 1
            changes.append((code_clean, old_url, new_url, page_num, bold_before))
        else:
            # Ma khong co trong Excel: giu nguyen text, ghi chu vao log
            ctx_seg  = pos_map[start] if start < len(pos_map) else 0
            ctx_rPr  = segs[ctx_seg][1] if ctx_seg < len(segs) else None
            old_url  = segs[ctx_seg][2] if ctx_seg < len(segs) else None
            bold_before = _is_bold(ctx_rPr)
            changes.append((code_clean, old_url, None, page_num, bold_before))  # new_url=None
            _emit_range(p_elem, full_text, start, end, pos_map, segs)
        cursor = end
    _emit_range(p_elem, full_text, cursor, len(full_text), pos_map, segs)

    # ── Quet ma loi dinh dang (fuzzy): co trong full_text nhung khong khop chinh xac
    exact_spans = {m.span() for m in CODE_PATTERN.finditer(full_text)}
    for fm in CODE_FUZZY.finditer(full_text):
        if fm.span() not in exact_spans:
            fcode = fm.group()
            ctx_seg  = pos_map[fm.start()] if fm.start() < len(pos_map) else 0
            ctx_rPr  = segs[ctx_seg][1] if ctx_seg < len(segs) else None
            old_url  = segs[ctx_seg][2] if ctx_seg < len(segs) else None
            bold_before = _is_bold(ctx_rPr)
            # new_url = '__MALFORMED__' de phan biet trong log
            changes.append((fcode, old_url, '__MALFORMED__', page_num, bold_before))

    return replaced, changes, breaks


def run_core(word_in: Path, excel_in: Path, word_out: Path,
             log_cb, done_cb, link_style: str = 'auto', link_color: str = '0000FF'):
    """Chạy trong thread riêng để không treo GUI."""
    try:
        # Đảm bảo thư viện đã load xong (thường chỉ cần <1s)
        if not _libs_ready.is_set():
            log_cb("⏳ Đang tải thư viện, vui lòng chờ...")
            _libs_ready.wait(timeout=30)

        log_cb(f"📊 Đang đọc Excel: {excel_in.name} ...")
        url_map = load_url_map(excel_in)
        log_cb(f"✅ Load được {len(url_map)} mã từ Excel.")
        if not url_map:
            log_cb("⚠️  Không đọc được mã nào. Kiểm tra lại file Excel."); done_cb(False); return

        log_cb(f"📖 Đang mở Word: {word_in.name} ...")
        doc = Document(str(word_in))

        style_label = {'auto': 'Bình thường (kế thừa)', 'bold': 'Luôn in đậm', 'normal': 'Luôn in thường'}
        log_cb(f"🔄 Đang gắn link — kiểu chữ: {style_label.get(link_style, link_style)}")
        total       = 0
        all_changes = []  # [(code, url, page, is_bold)]
        cur_page    = 1   # Trang hien tai (dem tu page break)

        for para in doc.paragraphs:
            n, ch, brk = process_paragraph(para, url_map, cur_page, link_style, link_color)
            total    += n
            all_changes.extend(ch)
            cur_page += brk

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        n, ch, brk = process_paragraph(para, url_map, cur_page, link_style, link_color)
                        total    += n
                        all_changes.extend(ch)
                        cur_page += brk

        doc.save(str(word_out))
        log_cb(f"✅ Hoàn tất! Đã thay thế {total} mã có link.")
        log_cb(f"💾 File lưu tại: {word_out}")

        # ── Ghi file log thay đổi .txt ────────────────────────
        log_path = word_out.parent / (word_out.stem + "_changes.txt")
        found     = [(ma, ou, nu, pg, bd) for ma, ou, nu, pg, bd in all_changes
                     if nu and nu != '__MALFORMED__']
        not_found = [(ma, ou, nu, pg, bd) for ma, ou, nu, pg, bd in all_changes
                     if nu is None]
        malformed = [(ma, ou, nu, pg, bd) for ma, ou, nu, pg, bd in all_changes
                     if nu == '__MALFORMED__']
        missing_count  = len(set(ma for ma, *_ in not_found))
        malform_count  = len(set(ma for ma, *_ in malformed))

        with open(log_path, "w", encoding="utf-8") as f:
            f.write("DANH SACH MA DUOC GAN LINK VA MA KHONG CO TRONG EXCEL\n")
            f.write(f"File Word : {word_out.name}\n")
            f.write(f"File Excel: {excel_in.name}\n")
            f.write(f"Tong so ma da gan link    : {total}\n")
            f.write(f"Tong so ma KHONG co Excel : {missing_count}\n")
            f.write(f"Tong so ma LOI DINH DANG  : {malform_count}\n")
            HDR = (f"{'STT':<5}  {'TRANG':^6}  {'IN DAM':^6}  "
                   f"{'MA':<22}  {'LINK CU':<55}  LINK MOI\n")
            SEP = "=" * 145
            DIV = "-" * 145

            # ── Phan 1: Ma co link ──
            f.write("\n" + SEP + "\n")
            f.write(f"[PHAN 1]  MA DA DUOC GAN LINK ({total} ma)\n")
            f.write(SEP + "\n"); f.write(HDR); f.write(DIV + "\n")
            for i, (ma, old_url, new_url, pg, bold) in enumerate(found, 1):
                dam     = "Co" if bold else "Khong"
                old_str = old_url if old_url else "Khong"
                f.write(f"{i:<5}  {pg:^6}  {dam:^6}  [{ma}]  "
                        f"{old_str:<55}  {new_url}\n")
            f.write(SEP + "\n")

            # ── Phan 2: Ma KHONG co trong Excel ──
            f.write("\n" + SEP + "\n")
            f.write(f"[PHAN 2]  MA TRONG WORD NHUNG KHONG CO TRONG EXCEL ({missing_count} ma)\n")
            f.write(SEP + "\n"); f.write(HDR); f.write(DIV + "\n")
            for i, (ma, old_url, _, pg, bold) in enumerate(not_found, 1):
                dam     = "Co" if bold else "Khong"
                old_str = old_url if old_url else "Khong"
                f.write(f"{i:<5}  {pg:^6}  {dam:^6}  [{ma}]  "
                        f"{old_str:<55}  *** Khong co trong Excel ***\n")
            f.write(SEP + "\n")

            # ── Phan 3: Ma loi dinh dang (co khoang trang) ──
            f.write("\n" + SEP + "\n")
            f.write(f"[PHAN 3]  MA LOI DINH DANG - CO KHOANG TRANG BEN TRONG ({malform_count} ma)\n")
            f.write("          Vi du: [H01.01.01. 07] thay vi [H01.01.01.07]\n")
            f.write(SEP + "\n"); f.write(HDR); f.write(DIV + "\n")
            for i, (ma, old_url, _, pg, bold) in enumerate(malformed, 1):
                dam     = "Co" if bold else "Khong"
                old_url_str = old_url if old_url else "Khong"
                # Goi y ten ma dung
                clean = re.sub(r'\s+', '', ma)  # Xoa khoang trang
                suggest = f"  (goi y: {clean})" if clean != ma else ""
                f.write(f"{i:<5}  {pg:^6}  {dam:^6}  {ma:<22}  "
                        f"{old_url_str:<55}  *** LOI DINH DANG{suggest} ***\n")
            f.write(SEP + "\n")

        log_cb(f"📄 File log: {log_path.name}")
        if not_found:
            log_cb(f"⚠️  {missing_count} mã trong Word KHONG có trong Excel (xem Phần 2)")
        if malformed:
            log_cb(f"🛑  {malform_count} mã LOI ĐỊNH DẠNG (có khoảng trắng) (xem Phần 3)")

        done_cb(True)
    except Exception as e:
        import traceback
        log_cb(f"❌ Lỗi: {e}")
        log_cb(traceback.format_exc())
        done_cb(False)


# ══════════════════════════════════════════════════════════
# ── THEME ─────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════
BG       = "#1E2230"
SIDEBAR  = "#161A27"
ACCENT   = "#0070C0"
ACCENT2  = "#00A8FF"
FG       = "#E8EAF0"
FG2      = "#8892A4"
ENTRY_BG = "#252B3B"
BTN_BG   = "#0070C0"
BTN_FG   = "#FFFFFF"
SUCCESS  = "#00C896"
ERROR    = "#FF4F6D"
FONT     = ("Segoe UI", 10)
FONT_B   = ("Segoe UI", 10, "bold")
FONT_H   = ("Segoe UI", 13, "bold")
FONT_SM  = ("Segoe UI", 9)


# ══════════════════════════════════════════════════════════
# ── APP CLASS ─────────────────────────────────────────────
# ══════════════════════════════════════════════════════════
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Gắn Hyperlink vào Word  ·  @HT-Labs")
        # ── Icon cửa sổ & taskbar ────────────────────────────────────────
        _ico = Path(__file__).parent / "icon.ico"
        if _ico.exists():
            try:
                self.iconbitmap(str(_ico))   # Windows: set icon taskbar + title bar
            except Exception:
                pass
        # ─────────────────────────────────────────────────────────────────
        self.geometry("960x720")
        self.minsize(760, 600)
        self.configure(bg=BG)
        self.resizable(True, True)
        self._build_ui()
        self._auto_detect()
        # ── Preload heavy libs ngay sau khi UI hiện ─────────────────────
        threading.Thread(target=_preload_libs, daemon=True).start()
        self.after(100, self._check_libs_ready)

    def _check_libs_ready(self):
        """Cập nhật status bar khi thư viện đã load xong."""
        if _libs_ready.is_set():
            self._status("✅ Sẵn sàng", SUCCESS)
        else:
            self._status("⏳ Đang khởi tạo thư viện...", ACCENT2)
            self.after(150, self._check_libs_ready)

    # ── tự động điền đường dẫn ────────────────────────────
    def _auto_detect(self):
        # Thu muc chua app.py la mac dinh
        here = Path(__file__).parent

        # ── Word: uu tien demo.docx, roi bat ky .docx, fallback hien duong dan mac dinh
        word_path = here / "demo.docx"
        existing = [f for f in here.glob("*.docx") if not f.name.startswith("~$")]
        if existing:
            # Uu tien demo.docx neu co
            demo = here / "demo.docx"
            word_path = demo if demo in existing else existing[0]
        # Luon hien duong dan (du file chua ton tai)
        self.var_word_in.set(str(word_path))

        # ── Excel: uu tien demo_mc.xlsx, roi bat ky .xlsx, fallback mac dinh
        excel_path = here / "demo_mc.xlsx"
        for name in ["demo_mc.xlsx", "link_minh_chung_export.xlsx"]:
            p = here / name
            if p.exists():
                excel_path = p; break
        else:
            xlsx_list = [f for f in here.glob("*.xlsx") if not f.name.startswith("~$")]
            if xlsx_list:
                excel_path = xlsx_list[0]
        # Luon hien duong dan (du file chua ton tai)
        self.var_excel.set(str(excel_path))

        # Output: cùng thư mục Word gốc, mặc định thêm _done
        self._refresh_output()

    def _refresh_output(self, *_):
        """Cập nhật đường dẫn output mỗi khi Word input hoặc tên file thay đổi."""
        word_in = self.var_word_in.get().strip()
        name    = self.var_out_name.get().strip()
        if word_in:
            p = Path(word_in)
            stem = name if name else (p.stem + "_done")
            self.var_word_out.set(str(p.parent / (stem + ".docx")))
        elif name:
            cur = self.var_word_out.get()
            if cur:
                p = Path(cur)
                self.var_word_out.set(str(p.parent / (name + ".docx")))

    # ── UI ────────────────────────────────────────────────
    def _build_ui(self):
        # ── Sidebar ───────────────────────────────────────
        sidebar = tk.Frame(self, bg=SIDEBAR, width=200)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        tk.Label(sidebar, text="🔗", font=("Segoe UI Emoji", 30),
                 bg=SIDEBAR, fg=ACCENT2).pack(pady=(24, 2))
        tk.Label(sidebar, text="Gắn Hyperlink\nvào Word",
                 font=("Segoe UI", 12, "bold"), bg=SIDEBAR,
                 fg=FG, justify="center").pack()
        tk.Label(sidebar,
                 text="Thay [H..] → link\ncó thể click",
                 font=FONT_SM, bg=SIDEBAR, fg=FG2, justify="center").pack(pady=(4, 0))

        ttk.Separator(sidebar, orient="horizontal").pack(fill="x", padx=16, pady=16)

        info = [
            ("📄", "Word gốc",  "File .docx cần gắn link"),
            ("📊", "Excel URL", "Mapping mã → đường dẫn"),
            ("💾", "File xuất", "Word kết quả sẽ lưu vào"),
        ]
        for icon, title, desc in info:
            row = tk.Frame(sidebar, bg=SIDEBAR)
            row.pack(fill="x", padx=12, pady=5)
            tk.Label(row, text=icon, font=("Segoe UI Emoji", 12),
                     bg=SIDEBAR, fg=ACCENT2, width=3).pack(side="left")
            col = tk.Frame(row, bg=SIDEBAR)
            col.pack(side="left", padx=(4, 0))
            tk.Label(col, text=title, font=FONT_B, bg=SIDEBAR, fg=FG, anchor="w").pack(anchor="w")
            tk.Label(col, text=desc, font=FONT_SM, bg=SIDEBAR, fg=FG2, anchor="w").pack(anchor="w")

        tk.Label(sidebar, text="v2.0  ·  @HT-Labs",
                 font=FONT_SM, bg=SIDEBAR, fg=FG2).pack(side="bottom", pady=12)

        # ── Main (vertical: form_panel + log_panel) ────────
        main = tk.Frame(self, bg=BG)
        main.pack(side="left", fill="both", expand=True)

        # ── FORM PANEL (fixed height, compact) ─────────────
        form_panel = tk.Frame(main, bg=BG)
        form_panel.pack(fill="x", padx=20, pady=(16, 0))

        # Header (1 dòng nhỏ)
        hdr = tk.Frame(form_panel, bg=BG)
        hdr.pack(fill="x", pady=(0, 6))
        tk.Label(hdr, text="Cấu hình", font=FONT_H, bg=BG, fg=FG).pack(side="left")
        tk.Label(hdr, text="  — chọn file rồi bấm Chạy",
                 font=FONT_SM, bg=BG, fg=FG2).pack(side="left", pady=(4, 0))

        tk.Frame(form_panel, bg="#2E3650", height=1).pack(fill="x", pady=(0, 8))

        # Khai báo vars
        self.var_word_in  = tk.StringVar()
        self.var_excel    = tk.StringVar()
        self.var_word_out = tk.StringVar()
        self.var_out_name = tk.StringVar()
        self.var_word_in.trace_add("write", self._refresh_output)
        self.var_out_name.trace_add("write", self._refresh_output)

        # Grid 2 cột: trái=Word, phải=Excel
        grid = tk.Frame(form_panel, bg=BG)
        grid.pack(fill="x")
        grid.columnconfigure(0, weight=1)
        grid.columnconfigure(1, weight=1)

        def _lbl(parent, text):
            tk.Label(parent, text=text, font=FONT_SM, bg=BG, fg=FG2
                     ).pack(anchor="w", pady=(0, 2))

        def _entry_btn(parent, var, cmd):
            f = tk.Frame(parent, bg=BG)
            f.pack(fill="x", pady=(0, 8))
            f.columnconfigure(0, weight=1)
            tk.Entry(f, textvariable=var, font=FONT,
                     bg=ENTRY_BG, fg=FG, relief="flat",
                     insertbackground=FG, bd=0
                     ).grid(row=0, column=0, sticky="ew", ipady=6, padx=(0, 6))
            tk.Button(f, text="…", font=FONT_B, bg=ACCENT, fg=BTN_FG,
                      relief="flat", cursor="hand2", padx=10, pady=3,
                      command=cmd).grid(row=0, column=1, sticky="e")

        # Cột trái: Word
        left = tk.Frame(grid, bg=BG)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        _lbl(left, "📄 File Word gốc (.docx)")
        _entry_btn(left, self.var_word_in, lambda: self._browse_word())

        # Cột phải: Excel
        right = tk.Frame(grid, bg=BG)
        right.grid(row=0, column=1, sticky="nsew")
        _lbl(right, "📊 File Excel mapping (.xlsx)")
        _entry_btn(right, self.var_excel,
                   lambda: self._browse(self.var_excel, [("Excel","*.xlsx"),("All","*.*")]))

        # Hàng 2: Output path + style/color (ngang)
        row2 = tk.Frame(form_panel, bg=BG)
        row2.pack(fill="x")
        row2.columnconfigure(0, weight=1)
        row2.columnconfigure(1, weight=1)

        # Output: ten file + full path (col trai)
        out_left = tk.Frame(row2, bg=BG)
        out_left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        _lbl(out_left, "✏️ Tên file xuất  (bỏ trống = _done)")
        nf = tk.Frame(out_left, bg=BG)
        nf.pack(fill="x", pady=(0, 4))
        nf.columnconfigure(0, weight=1)
        tk.Entry(nf, textvariable=self.var_out_name, font=FONT,
                 bg=ENTRY_BG, fg=FG, relief="flat", insertbackground=FG, bd=0
                 ).grid(row=0, column=0, sticky="ew", ipady=6, padx=(0, 6))
        tk.Label(nf, text="_done", font=FONT_SM, bg=BG, fg=FG2
                 ).grid(row=0, column=1, sticky="w")

        _lbl(out_left, "💾 Đường dẫn xuất (tự động)")
        of = tk.Frame(out_left, bg=BG)
        of.pack(fill="x", pady=(0, 4))
        of.columnconfigure(0, weight=1)
        tk.Entry(of, textvariable=self.var_word_out, font=FONT,
                 bg=ENTRY_BG, fg=FG2, relief="flat",
                 insertbackground=FG, bd=0, state="readonly"
                 ).grid(row=0, column=0, sticky="ew", ipady=6, padx=(0, 6))
        tk.Button(of, text="…", font=FONT_B, bg=ACCENT, fg=BTN_FG,
                  relief="flat", cursor="hand2", padx=10, pady=3,
                  command=lambda: self._save_browse(self.var_word_out)
                  ).grid(row=0, column=1, sticky="e")

        # Style + Color (col phai)
        out_right = tk.Frame(row2, bg=BG)
        out_right.grid(row=0, column=1, sticky="nsew")

        _lbl(out_right, "🅱 Kiểu chữ khi gắn link")
        self.var_link_style = tk.StringVar(value='auto')
        _style_labels = ['Binh thuong (ke thua goc)', 'Luon in dam', 'Luon in thuong']
        _style_vals   = ['auto', 'bold', 'normal']
        style_cb = ttk.Combobox(out_right, values=_style_labels,
                                state='readonly', font=FONT_SM, width=26)
        style_cb.current(0)
        style_cb.pack(anchor='w', ipady=4, pady=(0, 6), fill='x')
        def _on_style(e, cb=style_cb, vals=_style_vals):
            self.var_link_style.set(vals[cb.current()])
        style_cb.bind('<<ComboboxSelected>>', _on_style)

        _lbl(out_right, "🎨 Màu chữ  (hex 6 ký tự)")
        cf = tk.Frame(out_right, bg=BG)
        cf.pack(anchor='w', pady=(0, 2))
        tk.Label(cf, text="#", font=("Consolas", 10), bg=BG, fg=FG2).pack(side='left')
        self.var_link_color = tk.StringVar(value='0000FF')
        self._color_entry = tk.Entry(cf, textvariable=self.var_link_color,
                 font=("Consolas", 10), bg=ENTRY_BG, fg=FG,
                 width=8, relief='flat', insertbackground=FG)
        self._color_entry.pack(side='left', ipady=4, padx=(2, 8))
        self._color_swatch = tk.Label(cf, text='      ', bg='#0000FF',
                                      width=4, height=1, relief='flat')
        self._color_swatch.pack(side='left')
        tk.Button(cf, text="↺ Mặc định", font=FONT_SM,
                  bg=ENTRY_BG, fg=FG2, relief='flat', cursor='hand2',
                  padx=6, pady=1,
                  command=lambda: self.var_link_color.set('0000FF')
                  ).pack(side='left', padx=(8, 0))
        # Canh bao nam duoi
        self._color_warn = tk.Label(out_right, text='', font=FONT_SM, bg=BG, fg=ERROR, anchor='w')
        self._color_warn.pack(anchor='w', pady=(0, 4))
        def _update_swatch(*_):
            raw = self.var_link_color.get().strip().lstrip('#')
            valid = len(raw) == 6 and all(c in '0123456789ABCDEFabcdef' for c in raw)
            if valid:
                self._color_entry.configure(bg=ENTRY_BG)
                self._color_swatch.configure(bg='#' + raw.upper())
                self._color_warn.configure(text='')
            else:
                if raw:
                    self._color_entry.configure(bg='#4A1A1A')
                    self._color_warn.configure(text='⚠ Mã màu sai (VD: FF0000)')
                else:
                    self._color_entry.configure(bg=ENTRY_BG)
                    self._color_warn.configure(text='')
        self.var_link_color.trace_add('write', _update_swatch)


        tk.Frame(form_panel, bg="#2E3650", height=1).pack(fill="x", pady=(10, 0))

        # ── Buttons + progress ─────────────────────────────
        btn_panel = tk.Frame(main, bg=BG)
        btn_panel.pack(fill="x", padx=20, pady=(8, 4))

        self.btn_run = tk.Button(
            btn_panel, text="▶  Chạy — Gắn Link",
            font=("Segoe UI", 11, "bold"),
            bg=BTN_BG, fg=BTN_FG, activebackground=ACCENT2,
            activeforeground=BTN_FG, relief="flat", cursor="hand2",
            padx=20, pady=9, command=self._on_run
        )
        self.btn_run.pack(side="left")

        self.btn_open = tk.Button(
            btn_panel, text="📂  Mở file kết quả",
            font=FONT, bg=ENTRY_BG, fg=FG2,
            activebackground="#333", activeforeground=FG,
            relief="flat", cursor="hand2",
            padx=14, pady=9, state="disabled",
            command=self._open_output
        )
        self.btn_open.pack(side="left", padx=(10, 0))

        self.lbl_status = tk.Label(btn_panel, text=" ", font=FONT_SM, bg=BG, fg=FG2)
        self.lbl_status.pack(side="left", padx=(16, 0))

        style_ttk = ttk.Style()
        style_ttk.theme_use("default")
        style_ttk.configure("Blue.Horizontal.TProgressbar",
                             troughcolor=ENTRY_BG, background=ACCENT, thickness=5)
        self.progress = ttk.Progressbar(
            main, style="Blue.Horizontal.TProgressbar", mode="indeterminate"
        )
        self.progress.pack(fill="x", padx=20, pady=(2, 0))

        # ── LOG PANEL (expand chiếm toàn bộ phần còn lại) ─
        log_bar = tk.Frame(main, bg=BG)
        log_bar.pack(fill="x", padx=20, pady=(10, 4))
        tk.Label(log_bar, text="📋  Nhật ký", font=FONT_B, bg=BG, fg=FG).pack(side="left")
        tk.Button(log_bar, text="Xóa", font=FONT_SM, bg=ENTRY_BG, fg=FG2,
                  relief="flat", cursor="hand2", padx=8, pady=2,
                  command=self._clear_log).pack(side="right")

        log_wrap = tk.Frame(main, bg=ENTRY_BG, bd=0)
        log_wrap.pack(fill="both", expand=True, padx=20, pady=(0, 16))

        self.log_box = tk.Text(
            log_wrap, bg=ENTRY_BG, fg=FG,
            font=("Consolas", 10), relief="flat",
            wrap="word", state="disabled",
            insertbackground=FG, padx=12, pady=10
        )
        self.log_box.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(log_wrap, command=self.log_box.yview)
        sb.pack(side="right", fill="y")
        self.log_box.configure(yscrollcommand=sb.set)

        self.log_box.tag_configure("ok",   foreground=SUCCESS)
        self.log_box.tag_configure("err",  foreground=ERROR)
        self.log_box.tag_configure("info", foreground=ACCENT2)
        self.log_box.tag_configure("warn", foreground="#FFB347")


    # ── helpers ───────────────────────────────────────────
    def _file_row(self, parent, row, label, var, cmd):
        """Hàng nhập file cho Word và Excel (2 cột đầu tiên)."""
        tk.Label(parent, text=label, font=FONT_B, bg=BG, fg=FG
                 ).grid(row=row*2, column=0, sticky="w", pady=(10, 2), columnspan=2)
        ent = tk.Entry(parent, textvariable=var, font=FONT,
                       bg=ENTRY_BG, fg=FG, relief="flat",
                       insertbackground=FG, bd=0)
        ent.grid(row=row*2+1, column=0, sticky="ew", ipady=8, padx=(0, 8))
        tk.Button(parent, text="…", font=FONT_B, bg=ACCENT, fg=BTN_FG,
                  relief="flat", cursor="hand2", padx=12, pady=5,
                  command=cmd).grid(row=row*2+1, column=1, sticky="e")

    def _browse_word(self):
        """Browse Word gốc → cập nhật tự động output path và đường dẫn Excel."""
        init = str(Path(self.var_word_in.get()).parent) \
               if self.var_word_in.get() else str(Path.home() / "Downloads")
        p = filedialog.askopenfilename(
            filetypes=[("Word", "*.docx"), ("All", "*.*")],
            initialdir=init
        )
        if p:
            self.var_word_in.set(p)
            word_dir = Path(p).parent

            # ── Tự tìm Excel trong cùng thư mục với Word ──────────────
            # Ưu tiên: link_minh_chung_export.xlsx → bất kỳ .xlsx nào
            candidates = list(word_dir.glob("link_minh_chung_export.xlsx"))
            if not candidates:
                # Thử thư mục con "New folder"
                candidates = list((word_dir / "New folder").glob("link_minh_chung_export.xlsx"))
            if not candidates:
                # Lấy bất kỳ .xlsx nào trong cùng thư mục
                candidates = [f for f in word_dir.glob("*.xlsx")
                              if not f.name.startswith("~$")]
            if candidates:
                self.var_excel.set(str(candidates[0]))

            # Cập nhật đường dẫn output
            if not self.var_out_name.get():
                self._refresh_output()


    def _browse(self, var, ftypes):
        init = str(Path(var.get()).parent) if var.get() else str(Path.home() / "Downloads")
        p = filedialog.askopenfilename(filetypes=ftypes, initialdir=init)
        if p: var.set(p)

    def _save_browse(self, var):
        # Bắt đầu ở thư mục của Word gốc nếu có
        word_in = self.var_word_in.get().strip()
        if word_in and Path(word_in).exists():
            init_dir = str(Path(word_in).parent)
        elif var.get():
            init_dir = str(Path(var.get()).parent)
        else:
            init_dir = str(Path.home() / "Downloads")
        init_file = Path(var.get()).name if var.get() else "output_done.docx"
        p = filedialog.asksaveasfilename(
            defaultextension=".docx", filetypes=[("Word", "*.docx")],
            initialdir=init_dir, initialfile=init_file
        )
        if p:
            var.set(p)
            # Đồng bộ lại ô tên file
            self.var_out_name.set(Path(p).stem)

    def _log(self, msg: str):
        def _w():
            self.log_box.configure(state="normal")
            tag = "ok"   if msg.startswith("✅") else \
                  "err"  if msg.startswith("❌") or msg.startswith("⚠") else \
                  "info" if any(msg.startswith(c) for c in "📊📖🔄📄💾") else None
            self.log_box.insert("end", msg + "\n", tag)
            self.log_box.see("end")
            self.log_box.configure(state="disabled")
        self.after(0, _w)

    def _clear_log(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

    def _status(self, msg, color=FG2):
        self.after(0, lambda: self.lbl_status.configure(text=msg, fg=color))

    # ── run ───────────────────────────────────────────────
    def _on_run(self):
        word_in  = Path(self.var_word_in.get().strip())
        excel_in = Path(self.var_excel.get().strip())
        word_out = Path(self.var_word_out.get().strip())

        if not word_in.exists():
            messagebox.showerror("Lỗi", f"Không tìm thấy file Word:\n{word_in}"); return
        if not excel_in.exists():
            messagebox.showerror("Lỗi", f"Không tìm thấy file Excel:\n{excel_in}"); return
        if not word_out.parent.exists():
            messagebox.showerror("Lỗi", f"Thư mục lưu không tồn tại:\n{word_out.parent}"); return

        # Kiểm tra mã màu
        raw_color = self.var_link_color.get().strip().lstrip('#')
        if not (len(raw_color) == 6 and all(c in '0123456789ABCDEFabcdef' for c in raw_color)):
            messagebox.showerror("Mã màu sai",
                f"Mã màu '{raw_color}' không hợp lệ!\n"
                "Phải là 6 ký tự hex (0-9, A-F).\nVí dụ: FF0000, 0000FF, 0070C0")
            self._color_entry.focus_set()
            return

        self.btn_run.configure(state="disabled", text="⏳  Đang xử lý...")
        self.btn_open.configure(state="disabled")
        self.progress.start(12)
        self._status("Đang chạy...", ACCENT2)
        self._log("─" * 50)
        self._log(f"📄 Word  : {word_in.name}")
        self._log(f"📊 Excel : {excel_in.name}")
        self._log(f"💾 Output: {word_out}")

        def done(ok):
            self.progress.stop()
            self.btn_run.configure(state="normal", text="▶  Chạy — Gắn Link")
            if ok:
                self.btn_open.configure(state="normal")
                self._status("✅ Hoàn tất!", SUCCESS)
            else:
                self._status("❌ Có lỗi — xem log bên dưới.", ERROR)

        threading.Thread(
            target=run_core,
            args=(word_in, excel_in, word_out, self._log,
                  lambda ok: self.after(0, done, ok),
                  self.var_link_style.get(),
                  self.var_link_color.get().strip().lstrip('#')),
            daemon=True
        ).start()

    def _open_output(self):
        import subprocess
        p = self.var_word_out.get().strip()
        if p and Path(p).exists():
            subprocess.Popen(["start", "", p], shell=True)
        else:
            messagebox.showwarning("Không tìm thấy", f"File chưa tồn tại:\n{p}")


if __name__ == "__main__":
    App().mainloop()
