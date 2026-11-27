# -*- coding: utf-8 -*-
"""
Script xuất Tờ trình Chánh án TANDTC (.docx) chuẩn 100% định dạng gốc từ JSON Metadata.
Tuân thủ AGENTS.md (Bước 6):
- Mã hóa UTF-8 Unicode hoàn toàn bằng python-docx (không dùng PowerShell COM).
- Dựa trên phôi XML gốc base_to_trinh_tandtc.docx (giữ 100% VML lines, styles, margins, footer).
- Tích hợp File-lock check, Dry-run mode, Placeholder scan và Cross-consistency check.
"""
import argparse
import json
import os
import re
import sys
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import paths_config as cfg

BASE_TEMPLATE = os.path.join(cfg.TEMPLATES_DIR, "base_to_trinh_tandtc.docx")
FONT_NAME = "Times New Roman"


def set_run_font(run, size_pt=14, bold=False, italic=False, color_hex=None):
    run.font.name = FONT_NAME
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    if color_hex:
        run.font.color.rgb = RGBColor.from_string(color_hex)
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
        rFonts.set(qn(f"w:{attr}"), FONT_NAME)


def replace_para_text_preserve_vml(p, text, size_pt=13, bold=False, italic=False):
    """Thay nội dung text của đoạn văn nhưng BẢO TOÀN 100% run chứa <w:pict>/<w:drawing> (VML lines)."""
    p_el = p._p
    preserved_runs = []
    for child in list(p_el):
        if child.tag == qn("w:r"):
            has_media = (
                child.find(qn("w:pict")) is not None
                or child.find(qn("w:drawing")) is not None
                or child.find(".//{urn:schemas-microsoft-com:vml}line") is not None
                or child.find(".//{urn:schemas-microsoft-com:vml}shape") is not None
            )
            if has_media:
                # Chỉ giữ các node pict/drawing bên trong run này, xóa text cũ nếu lẫn vào
                for sub in list(child):
                    if sub.tag == qn("w:t"):
                        child.remove(sub)
                preserved_runs.append(child)
            p_el.remove(child)

    # Thêm lại media runs trước để anchor giữ đúng vị trí
    for mr in preserved_runs:
        p_el.append(mr)

    run = p.add_run(text)
    set_run_font(run, size_pt=size_pt, bold=bold, italic=italic)


def format_para(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_cm=1.25, before_pt=0, after_pt=6, line_spacing=None):
    pf = p.paragraph_format
    pf.alignment = align
    pf.first_line_indent = Cm(first_line_cm)
    pf.left_indent = Cm(0)
    pf.right_indent = Cm(0)
    pf.space_before = Pt(before_pt)
    pf.space_after = Pt(after_pt)
    if line_spacing is not None:
        pf.line_spacing = line_spacing
    return p


def add_body_para(doc, text="", size_pt=14, bold=False, italic=False,
                  align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_cm=1.25,
                  before_pt=0, after_pt=6, line_spacing=None):
    p = doc.add_paragraph()
    format_para(p, align=align, first_line_cm=first_line_cm,
                before_pt=before_pt, after_pt=after_pt, line_spacing=line_spacing)
    if text:
        r = p.add_run(text)
        set_run_font(r, size_pt=size_pt, bold=bold, italic=italic)
    return p


def add_rich_para(container, segments, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                  first_line_cm=1.25, before_pt=0, after_pt=6, line_spacing=None):
    """segments: list of tuples (text, size_pt, bold, italic)"""
    p = container.add_paragraph()
    format_para(p, align=align, first_line_cm=first_line_cm,
                before_pt=before_pt, after_pt=after_pt, line_spacing=line_spacing)
    for seg in segments:
        txt = seg[0]
        sz = seg[1] if len(seg) > 1 else 14
        b = seg[2] if len(seg) > 2 else False
        it = seg[3] if len(seg) > 3 else False
        r = p.add_run(txt)
        set_run_font(r, size_pt=sz, bold=b, italic=it)
    return p


def set_table_widths_and_indent(tbl, total_dxa, col_dxas, indent_dxa=0):
    tblPr = tbl._tbl.tblPr
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:w"), str(total_dxa))
    tblW.set(qn("w:type"), "dxa")

    tblInd = tblPr.find(qn("w:tblInd"))
    if tblInd is None:
        tblInd = OxmlElement("w:tblInd")
        tblPr.append(tblInd)
    tblInd.set(qn("w:w"), str(indent_dxa))
    tblInd.set(qn("w:type"), "dxa")

    tblGrid = tbl._tbl.find(qn("w:tblGrid"))
    if tblGrid is not None:
        gridCols = tblGrid.findall(qn("w:gridCol"))
        for gc, w in zip(gridCols, col_dxas):
            gc.set(qn("w:w"), str(w))

    for tr in tbl._tbl.findall(qn("w:tr")):
        tcs = tr.findall(qn("w:tc"))
        for tc, w in zip(tcs, col_dxas):
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


def clear_cell_paragraphs(cell):
    tc = cell._tc
    for p in list(tc.findall(qn("w:p"))):
        tc.remove(p)


def populate_header_table(doc, meta):
    """Bảng 1: Quốc hiệu, Tiêu ngữ, Đơn vị trình, Địa danh - Ngày tháng (bảo toàn VML lines)."""
    tbl0 = doc.tables[0]
    # Ô (0,0): TOÀ ÁN NHÂN DÂN TỐI CAO / VỤ GIÁM ĐỐC KIỂM TRA II
    c00 = tbl0.cell(0, 0)
    replace_para_text_preserve_vml(c00.paragraphs[0], meta.get("co_quan_cap_tren", "TOÀ ÁN NHÂN DÂN TỐI CAO"), size_pt=13, bold=False)
    replace_para_text_preserve_vml(c00.paragraphs[1], meta.get("don_vi_trinh", "VỤ GIÁM ĐỐC KIỂM TRA II"), size_pt=13, bold=True)

    # Ô (0,1): Quốc hiệu & Tiêu ngữ
    c01 = tbl0.cell(0, 1)
    replace_para_text_preserve_vml(c01.paragraphs[0], "CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM", size_pt=13, bold=True)
    replace_para_text_preserve_vml(c01.paragraphs[1], "Độc lập - Tự do - Hạnh phúc", size_pt=14, bold=True)

    # Ô (1,1): Địa danh, ngày tháng năm
    c11 = tbl0.cell(1, 1)
    ngay = meta.get("ngay_trinh") or meta.get("ngay") or "..."
    thang = meta.get("thang_trinh") or meta.get("thang") or "..."
    nam = meta.get("nam_trinh") or meta.get("nam") or "2026"
    ngay_str = f"{meta.get('dia_danh', 'Hà Nội')}, ngày {ngay} tháng {thang} năm {nam}"
    replace_para_text_preserve_vml(c11.paragraphs[0], ngay_str, size_pt=14, bold=False, italic=True)


def populate_title_block(doc, meta):
    """Khối tiêu đề TỜ TRÌNH & Trích yếu vụ án (các đoạn body [1..6] trong phôi)."""
    # Trong doc.paragraphs của phôi:
    # [0] dòng trống đệm sau Table 0
    # [1] TỜ TRÌNH
    # [2] CHÁNH ÁN TÒA ÁN NHÂN DÂN TỐI CAO
    # [3] Về vụ án dân sự “...”
    # [4] (Bản án dân sự phúc thẩm số ...
    # [5] ngày ... của Tòa án nhân dân ...)
    # [6] --------------------------------
    paras = doc.paragraphs
    replace_para_text_preserve_vml(paras[1], meta.get("tieu_de_van_ban", "TỜ TRÌNH"), size_pt=14, bold=True)
    replace_para_text_preserve_vml(paras[2], meta.get("nguoi_nhan_trinh", "CHÁNH ÁN TÒA ÁN NHÂN DÂN TỐI CAO"), size_pt=14, bold=True)

    # Đoạn 3: Về vụ án [loại] (đứng) + “Quan hệ tranh chấp” (nghiêng)
    p3 = paras[3]
    for child in list(p3._p):
        if child.tag == qn("w:r"):
            p3._p.remove(child)
    loai_va = meta.get("loai_vu_an") or "dân sự"
    qh_tc = meta.get("quan_he_tranh_chap") or meta.get("tieu_de_vu_an") or ""
    r3a = p3.add_run(f"Về vụ án {loai_va} ")
    set_run_font(r3a, size_pt=14, bold=False, italic=False)
    r3b = p3.add_run(f"“{qh_tc}”")
    set_run_font(r3b, size_pt=14, bold=False, italic=True)

    # Đoạn 4 & 5: (Bản án ... phúc thẩm số ... / ngày ... của ...)
    loai_ba = meta.get("loai_ban_an_pt") or "dân sự phúc thẩm"
    so_ba = meta.get("so_ba_pt") or "..."
    ngay_ba = meta.get("ngay_ba_pt") or "..."
    toa_an = meta.get("toa_an_pt") or "TAND..."
    line1 = meta.get("trich_yeu_ban_an_dong_1") or f"(Bản án {loai_ba} số {so_ba}"
    line2 = meta.get("trich_yeu_ban_an_dong_2") or f"ngày {ngay_ba} của {toa_an})"
    replace_para_text_preserve_vml(paras[4], line1, size_pt=14, bold=False, italic=False)
    replace_para_text_preserve_vml(paras[5], line2, size_pt=14, bold=False, italic=False)

    # Đoạn 6: 32 dấu gạch ngang
    replace_para_text_preserve_vml(paras[6], "--------------------------------", size_pt=14, bold=False, italic=False)


def populate_split_table(doc, meta, duong_su):
    """Bảng 2: Bảng chia đôi đặc thù TANDTC (Khung TKV bên trái | Đương sự bên phải)."""
    tbl1 = doc.tables[1]
    wide_mode = bool(meta.get("wide_party_table", False))
    if wide_mode:
        # Chế độ rộng (giống File Hiền - Thành): 17.50cm, tblInd = -1.50cm
        set_table_widths_and_indent(tbl1, total_dxa=9923, col_dxas=[5104, 4819], indent_dxa=-851)
        inner_w = 4815
        inner_ind = 31
    else:
        # Chế độ chuẩn lề trang 3cm/2cm (giống File Nhung - Đệ): 16.25cm, tblInd = 0
        set_table_widths_and_indent(tbl1, total_dxa=9214, col_dxas=[4253, 4961], indent_dxa=0)
        inner_w = 3995
        inner_ind = 0

    # --- Cột trái: Bảng lồng Khung TKV (cỡ 12pt) ---
    left_cell = tbl1.cell(0, 0)
    nested_tbl = left_cell.tables[0]
    set_table_widths_and_indent(nested_tbl, total_dxa=inner_w, col_dxas=[inner_w], indent_dxa=inner_ind)
    tkv_cell = nested_tbl.cell(0, 0)
    clear_cell_paragraphs(tkv_cell)

    ho_ten_tkv = meta.get("ho_ten_tkv") or "[CẦN ĐIỀN TÊN TKV]"
    cd_tp = meta.get("chuc_danh_tham_phan") or "Thẩm phán TANDTC"
    ht_tp = meta.get("ho_ten_tham_phan") or "[CẦN ĐIỀN TÊN THẨM PHÁN]"
    tkv_lines = [
        f"- TKV: {ho_ten_tkv}",
        f"- {cd_tp}: {ht_tp}",
        f"- Ngày báo cáo VT: {meta.get('ngay_bao_cao_vt', '')}".rstrip(),
        f"- Ngày báo cáo PCA: {meta.get('ngay_bao_cao_pca', '')}".rstrip(),
        f"- Ngày báo cáo Tổ TP: {meta.get('ngay_bao_cao_to_tp', '')}".rstrip(),
        f"- Ngày trình CA: {meta.get('ngay_trinh_ca', '')}".rstrip(),
        f"- Lý do trình: {meta.get('ly_do_trinh', '')}".rstrip(),
    ]
    for ln in tkv_lines:
        p = tkv_cell.add_paragraph()
        format_para(p, align=WD_ALIGN_PARAGRAPH.LEFT, first_line_cm=0, before_pt=0, after_pt=6)
        r = p.add_run(ln)
        set_run_font(r, size_pt=12, bold=False, italic=False)
    # 2 dòng trống cuối khung để ghi tay
    for _ in range(2):
        p = tkv_cell.add_paragraph()
        format_para(p, align=WD_ALIGN_PARAGRAPH.LEFT, first_line_cm=0, before_pt=0, after_pt=6)

    # --- Cột phải: Danh sách Đương sự (cỡ 14pt) ---
    right_cell = tbl1.cell(0, 1)
    clear_cell_paragraphs(right_cell)

    # 1 đoạn trống đầu ô (giống file gốc)
    p0 = right_cell.add_paragraph()
    format_para(p0, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_cm=0, before_pt=0, after_pt=6)

    add_rich_para(right_cell, [("Có các đương sự là:", 14, False, False)], first_line_cm=0, after_pt=6)

    def render_party_group(label, items, before_pt=0):
        if not items:
            return
        if len(items) == 1 and label in ("Nguyên đơn", "Bị đơn"):
            it = items[0]
            line = f" {it['ho_ten']}"
            if it.get("nam_sinh"):
                line += f", sinh năm {it['nam_sinh']};"
            else:
                line += ";"
            add_rich_para(
                right_cell,
                [("- ", 14, False, False), (f"{label}:", 14, True, False), (line, 14, False, False)],
                first_line_cm=0, before_pt=before_pt, after_pt=6
            )
            if it.get("dia_chi"):
                add_rich_para(right_cell, [(f"Địa chỉ: {it['dia_chi']}", 14, False, False)], first_line_cm=0, after_pt=6)
            if it.get("dai_dien"):
                dd = it["dai_dien"]
                add_rich_para(
                    right_cell,
                    [("- ", 14, False, False), (dd["vai_tro"], 14, False, True), (f": {dd['thong_tin']}", 14, False, False)],
                    first_line_cm=0, after_pt=6
                )
                if dd.get("dia_chi"):
                    add_rich_para(right_cell, [(f"Địa chỉ: {dd['dia_chi']}", 14, False, False)], first_line_cm=0, after_pt=6)
        else:
            add_rich_para(
                right_cell,
                [("- ", 14, False, False), (f"{label}:", 14, True, False)],
                first_line_cm=0, before_pt=before_pt, after_pt=6
            )
            for idx, it in enumerate(items, 1):
                line = f"{idx}. {it['ho_ten']}"
                if it.get("nam_sinh"):
                    line += f", sinh năm {it['nam_sinh']};"
                else:
                    line += ";"
                add_rich_para(right_cell, [(line, 14, False, False)], first_line_cm=0, after_pt=6)
                if it.get("dia_chi"):
                    add_rich_para(right_cell, [(f"Địa chỉ: {it['dia_chi']}", 14, False, False)], first_line_cm=0, after_pt=6)
                if it.get("dai_dien"):
                    dd = it["dai_dien"]
                    add_rich_para(
                        right_cell,
                        [("- ", 14, False, False), (dd["vai_tro"], 14, False, True), (f": {dd['thong_tin']}", 14, False, False)],
                        first_line_cm=0, after_pt=6
                    )
                    if dd.get("dia_chi"):
                        add_rich_para(right_cell, [(f"Địa chỉ: {dd['dia_chi']}", 14, False, False)], first_line_cm=0, after_pt=6)

    render_party_group("Nguyên đơn", duong_su.get("nguyen_don", []), before_pt=0)
    render_party_group("Bị đơn", duong_su.get("bi_don", []), before_pt=12)
    render_party_group("Người có quyền lợi, nghĩa vụ liên quan", duong_su.get("nguoi_lien_quan", []), before_pt=12)

    nhan_de_nghi = duong_su.get("nhan_nguoi_de_nghi", "Người đề nghị giám đốc thẩm")
    nguoi_de_nghi_val = duong_su.get("nguoi_de_nghi") or duong_su.get("nguoi_gui_don") or ""
    if not nguoi_de_nghi_val and duong_su.get("bi_don"):
        nguoi_de_nghi_val = duong_su["bi_don"][0].get("ho_ten", "")
    if nguoi_de_nghi_val:
        add_rich_para(
            right_cell,
            [(f"- {nhan_de_nghi}", 14, False, True), (f": {nguoi_de_nghi_val}", 14, False, False)],
            first_line_cm=0, before_pt=12, after_pt=6
        )


def add_signature_table(doc, ho_ten_tkv, chuc_danh_ky="Thư ký viên"):
    """Bảng 3: Bảng chữ ký Thư ký viên (2 cột không viền, mỗi cột 7.99cm)."""
    tbl = doc.add_table(rows=1, cols=2)
    set_no_borders(tbl)
    set_table_widths_and_indent(tbl, total_dxa=9060, col_dxas=[4530, 4530], indent_dxa=0)

    c_left = tbl.cell(0, 0)
    c_right = tbl.cell(0, 1)
    c_left.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    c_right.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    p_l = c_left.paragraphs[0]
    format_para(p_l, align=WD_ALIGN_PARAGRAPH.CENTER, first_line_cm=0, after_pt=6)

    p_r0 = c_right.paragraphs[0]
    format_para(p_r0, align=WD_ALIGN_PARAGRAPH.CENTER, first_line_cm=0, after_pt=6)
    r0 = p_r0.add_run(chuc_danh_ky)
    set_run_font(r0, size_pt=14, bold=True)

    for _ in range(3):
        pe = c_right.add_paragraph()
        format_para(pe, align=WD_ALIGN_PARAGRAPH.CENTER, first_line_cm=0, after_pt=6)

    p_name = c_right.add_paragraph()
    format_para(p_name, align=WD_ALIGN_PARAGRAPH.CENTER, first_line_cm=0, after_pt=6)
    rn = p_name.add_run(ho_ten_tkv)
    set_run_font(rn, size_pt=14, bold=True)


def check_file_lock(path):
    """Kiểm tra file đích có đang bị mở/khóa không trước khi ghi đè."""
    if not os.path.exists(path):
        return True, None
    try:
        with open(path, "r+b"):
            return True, None
    except PermissionError as e:
        return False, f"File đang bị mở hoặc khóa quyền ghi: {path} ({e})"
    except OSError as e:
        return False, f"Không thể truy cập file đích: {path} ({e})"


def collect_all_text(doc):
    texts = [p.text for p in doc.paragraphs]
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    texts.append(p.text)
                for ntbl in cell.tables:
                    for nrow in ntbl.rows:
                        for ncell in nrow.cells:
                            for np in ncell.paragraphs:
                                texts.append(np.text)
    return "\n".join(texts)


def validate_document(doc, data):
    errors = []
    warnings = []
    full_text = collect_all_text(doc)

    # 1. Placeholder scan
    placeholders = re.findall(r"\{\{[^}]+\}\}", full_text)
    if placeholders:
        errors.append(f"Còn sót placeholder chưa điền: {sorted(set(placeholders))}")

    # 2. Cross-consistency check
    meta = data.get("metadata", {})
    ds = data.get("duong_su", {})
    ng_de_nghi = ds.get("nguoi_de_nghi") or ds.get("nguoi_gui_don") or (ds.get("bi_don", [{}])[0].get("ho_ten", "") if ds.get("bi_don") else "")
    checks = [
        ("Số bản án phúc thẩm", meta.get("so_ba_pt", "")),
        ("Ngày bản án phúc thẩm", meta.get("ngay_ba_pt", "")),
        ("Tên Thư ký viên", meta.get("ho_ten_tkv", "")),
        ("Tên Thẩm phán", meta.get("ho_ten_tham_phan", "")),
    ]
    if ng_de_nghi:
        checks.append(("Người đề nghị GĐT/TT", ng_de_nghi))
    for label, val in checks:
        if not val:
            continue
        count = full_text.count(val)
        if count < 2:
            warnings.append(f"{label} ('{val}') chỉ xuất hiện {count} lần — cần rà soát lại tính nhất quán.")

    return errors, warnings


def build_to_trinh(data, out_path, dry_run=False):
    if not os.path.exists(BASE_TEMPLATE):
        raise FileNotFoundError(f"Không tìm thấy phôi chuẩn: {BASE_TEMPLATE}")

    meta = data["metadata"]
    duong_su = data["duong_su"]
    heading_size = int(meta.get("roman_heading_size_pt", 14))

    doc = Document(BASE_TEMPLATE)

    # 1. Bảng 1: Quốc hiệu - Tiêu ngữ - Ngày tháng
    populate_header_table(doc, meta)

    # 2. Khối Tiêu đề & Trích yếu
    populate_title_block(doc, meta)

    # 3. Bảng 2: Bảng chia đôi (Khung TKV | Đương sự)
    populate_split_table(doc, meta, duong_su)

    # 4. Mục I. TÓM TẮT NỘI DUNG VỤ ÁN
    add_body_para(doc, "I. TÓM TẮT NỘI DUNG VỤ ÁN", size_pt=heading_size, bold=True, before_pt=12, after_pt=6)
    for block in data.get("muc_1_tom_tat", []):
        if isinstance(block, str):
            add_body_para(doc, block, size_pt=14)
        else:
            if block.get("tieu_de"):
                add_body_para(doc, block["tieu_de"], size_pt=14, italic=True)
            for para in block.get("doan_van", []):
                add_body_para(doc, para, size_pt=14)

    # 5. Mục II. QUÁ TRÌNH GIẢI QUYẾT
    add_body_para(doc, "II. QUÁ TRÌNH GIẢI QUYẾT", size_pt=heading_size, bold=True, before_pt=6, after_pt=6)
    m2 = data.get("muc_2_qua_trinh") or data.get("muc_2_noi_dung_an") or {}
    # Sơ thẩm
    st_tieu_de = m2.get("so_tham_tieu_de", "1. Tòa án cấp sơ thẩm tuyên xử:")
    add_body_para(doc, st_tieu_de, size_pt=14, bold=True)
    for qd in m2.get("so_tham_quyet_dinh", []):
        add_body_para(doc, qd, size_pt=14, italic=True)
    for note in m2.get("so_tham_ghi_chu", []):
        add_body_para(doc, note, size_pt=14)
    # Phúc thẩm
    pt_tieu_de = m2.get("phuc_tham_tieu_de", "2. Tòa án cấp phúc thẩm tuyên xử:")
    add_body_para(doc, pt_tieu_de, size_pt=14, bold=True)
    for qd_pt in m2.get("phuc_tham_quyet_dinh", []):
        add_body_para(doc, qd_pt, size_pt=14, italic=True)
    for ts in m2.get("tien_su_sau_phuc_tham", []):
        add_body_para(doc, ts, size_pt=14)

    # 6. Mục III. ĐỀ NGHỊ CỦA ĐƯƠNG SỰ
    add_body_para(doc, "III. ĐỀ NGHỊ CỦA ĐƯƠNG SỰ", size_pt=heading_size, bold=True, before_pt=6, after_pt=6)
    m3 = data.get("muc_3_de_nghi") or data.get("muc_3_nhan_dinh") or {}
    m3_mo_dau = m3.get("mo_dau") or "Đương sự có đơn đề nghị xem xét bản án theo thủ tục giám đốc thẩm."
    add_body_para(doc, m3_mo_dau, size_pt=14)
    for ly_do in m3.get("cac_ly_do", []):
        add_body_para(doc, ly_do, size_pt=14)
    for vp in m3.get("vi_pham_to_tung", []):
        add_body_para(doc, f"- Về tố tụng: {vp}", size_pt=14)
    for vp in m3.get("vi_pham_noi_dung", []):
        add_body_para(doc, f"- Về nội dung: {vp}", size_pt=14)

    # 7. Mục IV. ĐỀ XUẤT GIẢI QUYẾT ĐƠN ĐỀ NGHỊ
    add_body_para(doc, "IV. ĐỀ XUẤT GIẢI QUYẾT ĐƠN ĐỀ NGHỊ", size_pt=heading_size, bold=True, before_pt=6, after_pt=6)
    m4 = data.get("muc_4_de_xuat") or data.get("muc_4_y_kien_de_xuat") or {}

    # IV.1 Ý kiến đề xuất của cán bộ nghiên cứu
    add_body_para(doc, "1. Ý kiến đề xuất của cán bộ nghiên cứu:", size_pt=14, bold=True)
    noi_dung_pt = m4.get("noi_dung_phan_tich") or m4.get("noi_dung_khang_nghi") or []
    for item in noi_dung_pt:
        if isinstance(item, str):
            # Tự nhận diện đánh số 1), 2)... để in đậm số thứ tự như file gốc
            m_num = re.match(r"^(\d+\))\s+(.*)$", item)
            m_sub = re.match(r"^(-\s+)([^:]+)(:\s*.*)$", item)
            if m_num:
                add_rich_para(doc, [(m_num.group(1), 14, True, False), (" " + m_num.group(2), 14, False, False)])
            elif m_sub:
                add_rich_para(doc, [(m_sub.group(1), 14, False, False), (m_sub.group(2), 14, False, True), (m_sub.group(3), 14, False, False)])
            else:
                add_body_para(doc, item, size_pt=14)
        else:
            add_body_para(doc, item.get("text", ""), size_pt=14, bold=item.get("bold", False), italic=item.get("italic", False))

    # Đoạn chốt đề xuất của TKV
    de_xuat_chot = m4.get("de_xuat_chot") or m4.get("de_xuat_chinh")
    if de_xuat_chot:
        add_rich_para(
            doc,
            [("Thư ký viên đề xuất", 14, True, False), (f": {de_xuat_chot}", 14, False, False)],
            after_pt=0 if not m4.get("kinh_trinh") else 6
        )
    if m4.get("kinh_trinh"):
        add_body_para(doc, m4["kinh_trinh"], size_pt=14, after_pt=0)

    # Bảng chữ ký Thư ký viên (áp dụng cho cả 2 trường hợp Kháng nghị & Trả lời đơn)
    ho_ten_tkv = meta.get("ho_ten_tkv") or "[CẦN ĐIỀN TÊN TKV]"
    cd_tp = meta.get("chuc_danh_tham_phan") or "Thẩm phán TANDTC"
    ht_tp = meta.get("ho_ten_tham_phan") or "[CẦN ĐIỀN TÊN THẨM PHÁN]"
    add_signature_table(doc, ho_ten_tkv, chuc_danh_ky=meta.get("chuc_danh_ky_tkv", "Thư ký viên"))

    # IV.2 Ý kiến của Thẩm phán (có thể điền dự thảo do AI đề xuất hoặc để trống)
    nam_y_kien = meta.get("nam_trinh") or meta.get("nam") or "2026"
    tieu_de_tp = f"2. Ý kiến của {cd_tp} {ht_tp} ngày      /     /{nam_y_kien}:"
    add_body_para(doc, tieu_de_tp, size_pt=14, bold=True, before_pt=12, after_pt=6)

    du_thao_tp = m4.get("du_thao_y_kien_tham_phan", [])
    if du_thao_tp:
        for para in du_thao_tp:
            add_body_para(doc, para, size_pt=14)
    else:
        for _ in range(3):
            add_body_para(doc, "", size_pt=14)

    # IV.3, IV.4, IV.5: Ý kiến Lãnh đạo (tự động bỏ qua khi che_do_rut_gon=True theo TB 383/679)
    che_do_rut_gon = bool(meta.get("che_do_rut_gon", False))
    lanh_dao_mac_dinh = [] if che_do_rut_gon else [
        f"3. Ý kiến của Lãnh đạo Vụ (Vụ trưởng / Phó Vụ trưởng) ngày      /     /{nam_y_kien}:",
        f"4. Ý kiến của Phó Chánh án Tòa án nhân dân tối cao ngày      /     /{nam_y_kien}:",
        f"5. Ý kiến của Chánh án Tòa án nhân dân tối cao ngày      /     /{nam_y_kien}:",
    ]
    cac_muc_ld = m4.get("y_kien_lanh_dao", lanh_dao_mac_dinh)
    for muc in cac_muc_ld:
        add_body_para(doc, muc, size_pt=14, bold=True, before_pt=12, after_pt=6)
        for _ in range(3):
            add_body_para(doc, "", size_pt=14)

    # --- VALIDATION GATE ---
    errors, warnings = validate_document(doc, data)
    print("================ BÁO CÁO KIỂM TRA XUẤT TỜ TRÌNH (STEP 6 LOG) ================")
    print(f"- Chế độ chạy     : {'DRY-RUN (chỉ kiểm tra, không ghi file)' if dry_run else 'XUẤT FILE CHÍNH THỨC'}")
    print(f"- File đích       : {out_path}")
    print(f"- Lỗi nghiêm trọng: {len(errors)}")
    for e in errors:
        print(f"  [ERROR] {e}")
    print(f"- Cảnh báo        : {len(warnings)}")
    for w in warnings:
        print(f"  [WARN]  {w}")

    if errors:
        raise ValueError("Dừng xuất file do phát hiện lỗi kiểm tra chéo/placeholder!")

    if dry_run:
        print("[DRY-RUN OK] Dữ liệu & định dạng hợp lệ 100%.")
        return

    # File-lock check trước khi ghi
    ok, lock_err = check_file_lock(out_path)
    if not ok:
        raise PermissionError(lock_err)

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    doc.save(out_path)
    print(f"[SUCCESS] Đã xuất Tờ trình chuẩn 100% định dạng gốc tại: {out_path}")


def main():
    parser = argparse.ArgumentParser(description="Xuất Tờ trình TANDTC (.docx) chuẩn UTF-8")
    parser.add_argument("--input", required=True, help="Đường dẫn file JSON dữ liệu đã khóa ở Bước 5.5")
    parser.add_argument("--output", required=True, help="Đường dẫn file .docx đầu ra trong 03_Ket_qua_Xu_ly/<Số>_Vu_<ten_vu>/")
    parser.add_argument("--dry-run", action="store_true", help="Chạy kiểm tra không ghi file")
    args = parser.parse_args()

    with open(args.input, "r", encoding="utf-8") as f:
        data = json.load(f)

    build_to_trinh(data, args.output, dry_run=args.dry_run)


if __name__ == "__main__":
    main()
