# -*- coding: utf-8 -*-
"""
docx_utils.py - Thư viện tiện ích dùng chung cho toàn bộ động cơ xử lý & xuất Word (.docx).
Single Source of Truth giúp loại bỏ lặp code (DRY) giữa:
  export_dossier.py, export_to_trinh.py, export_khang_nghi.py,
  export_6step_report.py, universal_visual_filler.py, fast_reader.py, verify_system.py.
"""
import os
import re
import sys
import unicodedata
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

FONT_NAME = "Times New Roman"

PLACEHOLDER_RE = re.compile(r"\{\{[^}]+\}\}")
UNFILLED_MARKER_RE = re.compile(r"\[CẦN ĐIỀN[^\]]*\]")


def ensure_utf8_stdout():
    """Đảm bảo stdout/stderr in tiếng Việt UTF-8 ổn định trên Windows."""
    for stream in (sys.stdout, sys.stderr):
        if stream is not None and getattr(stream, "encoding", None) != "utf-8":
            try:
                stream.reconfigure(encoding="utf-8")
            except Exception:
                pass


def set_run_font(run, size_pt=14, bold=False, italic=False, color=None, color_hex=None):
    """Chuẩn hóa font Times New Roman đầy đủ 4 thuộc tính rFonts (ascii, hAnsi, cs, eastAsia)."""
    run.font.name = FONT_NAME
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    if color is not None:
        run.font.color.rgb = color
    elif color_hex:
        run.font.color.rgb = RGBColor.from_string(color_hex)

    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
        rFonts.set(qn(f"w:{attr}"), FONT_NAME)


def run_has_media(run_el):
    """Kiểm tra run XML có chứa VML line/shape, pict hoặc drawing không."""
    return (
        run_el.find(qn("w:pict")) is not None
        or run_el.find(qn("w:drawing")) is not None
        or run_el.find(".//{urn:schemas-microsoft-com:vml}line") is not None
        or run_el.find(".//{urn:schemas-microsoft-com:vml}shape") is not None
    )


def replace_para_runs_preserve_vml(p, segments):
    """
    Thay các text runs trong đoạn văn bằng danh sách segments:
      [(text, size_pt, bold, italic, is_br), ...]
    nhưng BẢO TOÀN 100% mọi run chứa <w:pict>/<w:drawing>/<v:line>/<v:shape>.
    """
    p_el = p._p
    preserved_runs = []
    for child in list(p_el):
        if child.tag == qn("w:r"):
            if run_has_media(child):
                for sub in list(child):
                    if sub.tag in (qn("w:t"), qn("w:br")):
                        child.remove(sub)
                preserved_runs.append(child)
            p_el.remove(child)

    for mr in preserved_runs:
        p_el.append(mr)

    for seg in segments:
        txt = seg[0]
        sz = seg[1] if len(seg) > 1 else 14
        b = seg[2] if len(seg) > 2 else False
        it = seg[3] if len(seg) > 3 else False
        is_br = seg[4] if len(seg) > 4 else False
        r = p.add_run()
        if is_br:
            r.add_break()
            if txt:
                r.add_text(txt)
        else:
            r.text = txt
        set_run_font(r, size_pt=sz, bold=b, italic=it)


def replace_para_text_preserve_vml(p, text, size_pt=13, bold=False, italic=False):
    """Thay toàn bộ text của đoạn văn nhưng bảo toàn 100% VML/pict/drawing."""
    replace_para_runs_preserve_vml(p, [(text, size_pt, bold, italic, False)])


def replace_tags_preserve_format(p, replacements, default_size_pt=None):
    """
    Thay thế các tag {{...}} theo từng run (run-level replacement):
    - BẢO TOÀN 100% định dạng nội dòng (in đậm / in nghiêng / cỡ chữ của các run khác trên cùng dòng).
    - Xử lý chính xác cả trường hợp Word tự tách 1 tag {{...}} ra nhiều run XML liên tiếp.
    - BẢO TOÀN 100% run chứa VML/pict/drawing.
    Trả về True nếu có ít nhất 1 tag được thay thế.
    """
    if not replacements:
        return False

    text_runs = [r for r in p.runs if not run_has_media(r._r)]
    if not text_runs:
        full_text = p.text
        new_text = full_text
        for tag, val in sorted(replacements.items(), key=lambda kv: len(kv[0]), reverse=True):
            if tag in new_text:
                new_text = new_text.replace(tag, str(val))
        if new_text != full_text:
            r = p.add_run(new_text)
            set_run_font(r, size_pt=default_size_pt or 14)
            return True
        return False

    # Lọc các tag thực sự xuất hiện trong đoạn văn (ưu tiên tag dài trước)
    full_text = "".join(r.text for r in text_runs)
    active_items = [
        (tag, str(val))
        for tag, val in sorted(replacements.items(), key=lambda kv: len(kv[0]), reverse=True)
        if tag and tag in full_text
    ]
    if not active_items:
        return False

    changed = False
    for tag, val in active_items:
        while True:
            full_text = "".join(r.text for r in text_runs)
            start = full_text.find(tag)
            if start == -1:
                break
            end = start + len(tag)

            # Xây bảng ánh xạ vị trí ký tự -> (run_idx, offset_trong_run)
            pos = 0
            start_r = end_r = None
            start_off = end_off = 0
            for idx, r in enumerate(text_runs):
                rlen = len(r.text)
                if start_r is None and pos <= start < pos + rlen:
                    start_r = idx
                    start_off = start - pos
                if pos < end <= pos + rlen:
                    end_r = idx
                    end_off = end - pos
                    break
                pos += rlen

            if start_r is None or end_r is None:
                break

            if start_r == end_r:
                cur = text_runs[start_r].text
                text_runs[start_r].text = cur[:start_off] + val + cur[end_off:]
            else:
                prefix = text_runs[start_r].text[:start_off]
                suffix = text_runs[end_r].text[end_off:]
                text_runs[start_r].text = prefix + val
                for mid in range(start_r + 1, end_r):
                    text_runs[mid].text = ""
                text_runs[end_r].text = suffix

            touched_run = text_runs[start_r]
            if default_size_pt is not None and touched_run.font.size is None:
                touched_run.font.size = Pt(default_size_pt)
            if not touched_run.font.name:
                touched_run.font.name = FONT_NAME
            changed = True

    return changed


def format_para(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_cm=1.25, left_cm=0,
                before_pt=0, after_pt=6, line_spacing=None):
    pf = p.paragraph_format
    pf.alignment = align
    pf.first_line_indent = Cm(first_line_cm)
    pf.left_indent = Cm(left_cm)
    pf.right_indent = Cm(0)
    pf.space_before = Pt(before_pt)
    pf.space_after = Pt(after_pt)
    if line_spacing is not None:
        pf.line_spacing = line_spacing
    return p


def set_no_borders(tbl):
    tblPr = tbl._tbl.tblPr
    old = tblPr.find(qn("w:tblBorders"))
    if old is not None:
        tblPr.remove(old)
    borders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"), "none")
        b.set(qn("w:sz"), "0")
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), "auto")
        borders.append(b)
    tblPr.append(borders)


def set_table_widths_and_indent(tbl, total_dxa, col_dxas, indent_dxa=None):
    tblPr = tbl._tbl.tblPr
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:w"), str(total_dxa))
    tblW.set(qn("w:type"), "dxa")

    if indent_dxa is not None:
        tblInd = tblPr.find(qn("w:tblInd"))
        if tblInd is None:
            tblInd = OxmlElement("w:tblInd")
            tblPr.append(tblInd)
        tblInd.set(qn("w:w"), str(indent_dxa))
        tblInd.set(qn("w:type"), "dxa")

    tblGrid = tbl._tbl.find(qn("w:tblGrid"))
    if tblGrid is not None:
        for gc, w in zip(tblGrid.findall(qn("w:gridCol")), col_dxas):
            gc.set(qn("w:w"), str(w))

    for tr in tbl._tbl.findall(qn("w:tr")):
        for tc, w in zip(tr.findall(qn("w:tc")), col_dxas):
            tcPr = tc.find(qn("w:tcPr"))
            if tcPr is None:
                tcPr = OxmlElement("w:tcPr")
                tc.insert(0, tcPr)
            tcW = tcPr.find(qn("w:tcW"))
            if tcW is None:
                tcW = OxmlElement("w:tcW")
                tcPr.append(tcW)
            tcW.set(qn("w:w"), str(w))
            tcW.set(qn("w:type"), "dxa")


def clear_cell_paragraphs(cell):
    tc = cell._tc
    for p in list(tc.findall(qn("w:p"))):
        tc.remove(p)


def check_file_lock(path):
    """Kiểm tra file đích có đang bị mở/khóa quyền ghi hay không."""
    if not os.path.exists(path):
        return True, None
    try:
        with open(path, "r+b"):
            return True, None
    except PermissionError as e:
        return False, f"File đang bị mở hoặc khóa quyền ghi: {path} ({e})"
    except OSError as e:
        return False, f"Không thể truy cập file đích: {path} ({e})"


def iter_all_paragraphs(container):
    """Duyệt đệ quy mọi đoạn văn trong document/header/footer kể cả bảng lồng nhiều cấp."""
    for p in getattr(container, "paragraphs", []):
        yield p
    seen_cells = set()
    for tbl in getattr(container, "tables", []):
        for row in tbl.rows:
            for cell in row.cells:
                tc = cell._tc
                if tc in seen_cells:
                    continue
                seen_cells.add(tc)
                yield from iter_all_paragraphs(cell)


def iter_document_paragraphs(doc):
    """Duyệt toàn bộ đoạn văn của cả body, bảng lồng nhiều cấp và mọi header/footer."""
    yield from iter_all_paragraphs(doc)
    seen_parts = set()
    for sec in doc.sections:
        for hf in (
            sec.header,
            sec.first_page_header,
            sec.even_page_header,
            sec.footer,
            sec.first_page_footer,
            sec.even_page_footer,
        ):
            if hf is None:
                continue
            try:
                part_el = hf._element
            except Exception:
                continue
            if part_el in seen_parts:
                continue
            seen_parts.add(part_el)
            yield from iter_all_paragraphs(hf)


def collect_all_text(doc):
    """Thu thập toàn bộ văn bản trong document (body, bảng lồng đa cấp, header/footer)."""
    return "\n".join(p.text for p in iter_document_paragraphs(doc))


def scan_unfilled_tokens(text):
    """
    Quét sót placeholder {{...}} và thẻ thiếu dữ liệu [CẦN ĐIỀN...].
    Trả về (placeholders_list, unfilled_markers_list).
    """
    placeholders = sorted(set(PLACEHOLDER_RE.findall(text)))
    unfilled_markers = sorted(set(UNFILLED_MARKER_RE.findall(text)))
    return placeholders, unfilled_markers


def slugify_vn(text):
    """Chuyển chuỗi tiếng Việt thành slug ASCII an toàn cho tên thư mục & file."""
    if not text:
        return ""
    s = str(text).replace("đ", "d").replace("Đ", "D")
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"[^A-Za-z0-9]+", "_", s)
    s = re.sub(r"_+", "_", s).strip("_")
    return s
