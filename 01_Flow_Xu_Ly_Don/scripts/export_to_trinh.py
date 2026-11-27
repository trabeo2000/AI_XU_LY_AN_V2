# -*- coding: utf-8 -*-
"""
Script xuất Tờ trình Chánh án TANDTC (.docx) chuẩn 100% định dạng gốc từ JSON Metadata.
Tuân thủ AGENTS.md (Bước 6):
- Mã hóa UTF-8 Unicode hoàn toàn bằng python-docx (không dùng PowerShell COM).
- Dựa trên phôi XML gốc base_to_trinh_tandtc.docx (giữ 100% VML lines, styles, margins, footer).
- Tích hợp chuẩn hóa Schema thống nhất, kiểm tra DOM phôi, File-lock check, Dry-run mode,
  Placeholder/[CẦN ĐIỀN] scan và Cross-consistency check.
"""
import argparse
import json
import os
import re
import sys
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import paths_config as cfg
import docx_utils as du
import payload_schema as ps

du.ensure_utf8_stdout()

BASE_TEMPLATE = cfg.BASE_TO_TRINH_TEMPLATE
FONT_NAME = du.FONT_NAME

# Tái xuất các hàm tiện ích dùng chung để bảo đảm tương thích ngược 100%
set_run_font = du.set_run_font
replace_para_text_preserve_vml = du.replace_para_text_preserve_vml
format_para = du.format_para
set_table_widths_and_indent = du.set_table_widths_and_indent
set_no_borders = du.set_no_borders
clear_cell_paragraphs = du.clear_cell_paragraphs
check_file_lock = du.check_file_lock
collect_all_text = du.collect_all_text


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


def verify_template_dom(doc):
    """Kiểm định cấu trúc phôi base_to_trinh_tandtc.docx trước khi ghi dữ liệu theo chỉ số."""
    if len(doc.tables) < 2:
        raise ValueError(f"Phôi Tờ trình không hợp lệ: cần >= 2 bảng, hiện có {len(doc.tables)}.")
    if len(doc.paragraphs) < 7:
        raise ValueError(f"Phôi Tờ trình không hợp lệ: cần >= 7 đoạn tiêu đề, hiện có {len(doc.paragraphs)}.")
    if len(doc.tables[1].cell(0, 0).tables) < 1:
        raise ValueError("Phôi Tờ trình không hợp lệ: thiếu bảng lồng Khung TKV trong ô (0,0) của Bảng 2.")


def populate_header_table(doc, meta):
    """Bảng 1: Quốc hiệu, Tiêu ngữ, Đơn vị trình, Địa danh - Ngày tháng (bảo toàn VML lines)."""
    tbl0 = doc.tables[0]
    c00 = tbl0.cell(0, 0)
    replace_para_text_preserve_vml(
        c00.paragraphs[0], meta.get("co_quan_cap_tren", ps.DEFAULT_CO_QUAN), size_pt=13, bold=False
    )
    replace_para_text_preserve_vml(
        c00.paragraphs[1], meta.get("don_vi_trinh", ps.DEFAULT_DON_VI), size_pt=13, bold=True
    )

    c01 = tbl0.cell(0, 1)
    replace_para_text_preserve_vml(c01.paragraphs[0], "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", size_pt=13, bold=True)
    replace_para_text_preserve_vml(c01.paragraphs[1], "Độc lập - Tự do - Hạnh phúc", size_pt=14, bold=True)

    c11 = tbl0.cell(1, 1)
    ngay = meta.get("ngay_trinh") or "..."
    thang = meta.get("thang_trinh") or "..."
    nam = meta.get("nam_trinh") or "2026"
    ngay_str = f"{meta.get('dia_danh', 'Hà Nội')}, ngày {ngay} tháng {thang} năm {nam}"
    replace_para_text_preserve_vml(c11.paragraphs[0], ngay_str, size_pt=14, bold=False, italic=True)


def populate_title_block(doc, meta):
    """Khối tiêu đề TỜ TRÌNH & Trích yếu vụ án (các đoạn body [1..6] trong phôi)."""
    paras = doc.paragraphs
    replace_para_text_preserve_vml(paras[1], meta.get("tieu_de_van_ban", "TỜ TRÌNH"), size_pt=14, bold=True)
    replace_para_text_preserve_vml(paras[2], meta.get("nguoi_nhan_trinh", "CHÁNH ÁN TÒA ÁN NHÂN DÂN TỐI CAO"), size_pt=14, bold=True)

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

    loai_ba = meta.get("loai_ban_an_pt") or "dân sự phúc thẩm"
    so_ba = meta.get("so_ba_pt") or "..."
    ngay_ba = meta.get("ngay_ba_pt") or "..."
    toa_an = meta.get("toa_an_pt") or "TAND..."
    line1 = meta.get("trich_yeu_ban_an_dong_1") or f"(Bản án {loai_ba} số {so_ba}"
    line2 = meta.get("trich_yeu_ban_an_dong_2") or f"ngày {ngay_ba} của {toa_an})"
    replace_para_text_preserve_vml(paras[4], line1, size_pt=14, bold=False, italic=False)
    replace_para_text_preserve_vml(paras[5], line2, size_pt=14, bold=False, italic=False)
    replace_para_text_preserve_vml(paras[6], "--------------------------------", size_pt=14, bold=False, italic=False)


def _render_dai_dien_items(right_cell, dd_list):
    """Hiển thị danh sách người đại diện đã chuẩn hóa (hỗ trợ cả 1 hoặc nhiều đại diện)."""
    for dd in dd_list or []:
        vai_tro = dd.get("vai_tro") or "Người đại diện"
        thong_tin = dd.get("thong_tin") or dd.get("mo_ta") or ""
        add_rich_para(
            right_cell,
            [("- ", 14, False, False), (vai_tro, 14, False, True), (f": {thong_tin}", 14, False, False)],
            first_line_cm=0, after_pt=6
        )
        dia_chi = dd.get("dia_chi") or ""
        if dia_chi and dia_chi.lower() not in thong_tin.lower():
            add_rich_para(right_cell, [(f"Địa chỉ: {dia_chi}", 14, False, False)], first_line_cm=0, after_pt=6)


def _format_party_heading_and_address(it):
    """Trả về (main_line, optional_address_line) từ party item đã chuẩn hóa."""
    if it.get("ho_ten"):
        line = it["ho_ten"]
        if it.get("nam_sinh"):
            line += f", sinh năm {it['nam_sinh']};"
        else:
            line += ";"
        return line, it.get("dia_chi", "")
    return it.get("mo_ta", ""), ""


def populate_split_table(doc, meta, duong_su):
    """Bảng 2: Bảng chia đôi đặc thù TANDTC (Khung TKV bên trái | Đương sự bên phải)."""
    tbl1 = doc.tables[1]
    wide_mode = bool(meta.get("wide_party_table", False))
    if wide_mode:
        set_table_widths_and_indent(tbl1, total_dxa=9923, col_dxas=[5104, 4819], indent_dxa=-851)
        inner_w = 4815
        inner_ind = 31
    else:
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
    for _ in range(2):
        p = tkv_cell.add_paragraph()
        format_para(p, align=WD_ALIGN_PARAGRAPH.LEFT, first_line_cm=0, before_pt=0, after_pt=6)

    # --- Cột phải: Danh sách Đương sự (cỡ 14pt) ---
    right_cell = tbl1.cell(0, 1)
    clear_cell_paragraphs(right_cell)

    p0 = right_cell.add_paragraph()
    format_para(p0, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_cm=0, before_pt=0, after_pt=6)
    add_rich_para(right_cell, [("Có các đương sự là:", 14, False, False)], first_line_cm=0, after_pt=6)

    def render_party_group(label, items, group_reps=None, before_pt=0):
        if not items:
            return
        if len(items) == 1 and label in ("Nguyên đơn", "Bị đơn"):
            it = items[0]
            main_line, addr_line = _format_party_heading_and_address(it)
            add_rich_para(
                right_cell,
                [("- ", 14, False, False), (f"{label}:", 14, True, False), (f" {main_line}", 14, False, False)],
                first_line_cm=0, before_pt=before_pt, after_pt=6
            )
            if addr_line:
                add_rich_para(right_cell, [(f"Địa chỉ: {addr_line}", 14, False, False)], first_line_cm=0, after_pt=6)
            _render_dai_dien_items(right_cell, it.get("dai_dien"))
        else:
            add_rich_para(
                right_cell,
                [("- ", 14, False, False), (f"{label}:", 14, True, False)],
                first_line_cm=0, before_pt=before_pt, after_pt=6
            )
            for idx, it in enumerate(items, 1):
                main_line, addr_line = _format_party_heading_and_address(it)
                add_rich_para(right_cell, [(f"{idx}. {main_line}", 14, False, False)], first_line_cm=0, after_pt=6)
                if addr_line:
                    add_rich_para(right_cell, [(f"Địa chỉ: {addr_line}", 14, False, False)], first_line_cm=0, after_pt=6)
                _render_dai_dien_items(right_cell, it.get("dai_dien"))
        _render_dai_dien_items(right_cell, group_reps)

    render_party_group("Nguyên đơn", duong_su.get("nguyen_don", []), duong_su.get("dai_dien_nguyen_don"), before_pt=0)
    render_party_group("Bị đơn", duong_su.get("bi_don", []), duong_su.get("dai_dien_bi_don"), before_pt=12)
    render_party_group(
        "Người có quyền lợi, nghĩa vụ liên quan",
        duong_su.get("nguoi_lien_quan", []),
        duong_su.get("dai_dien_nguoi_lien_quan"),
        before_pt=12,
    )

    nhan_de_nghi = duong_su.get("nhan_nguoi_de_nghi", "Người đề nghị giám đốc thẩm")
    nguoi_de_nghi_val = duong_su.get("nguoi_de_nghi", "")
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


def validate_document(doc, norm_data, allow_draft=False):
    errors, warnings = ps.validate_payload(norm_data, "to_trinh", allow_draft=allow_draft)
    full_text = collect_all_text(doc)

    # 1. Quét sót placeholder {{...}} và thẻ thiếu dữ liệu [CẦN ĐIỀN...]
    placeholders, unfilled = du.scan_unfilled_tokens(full_text)
    if placeholders:
        errors.append(f"Còn sót placeholder chưa điền: {placeholders}")
    if unfilled:
        (warnings if allow_draft else errors).append(f"Còn thẻ thiếu dữ liệu bắt buộc: {unfilled}")

    # 2. Cross-consistency check (bỏ qua giá trị rỗng hoặc '...')
    meta = norm_data.get("metadata", {})
    ds = norm_data.get("duong_su", {})
    checks = [
        ("Số bản án phúc thẩm", meta.get("so_ba_pt", "")),
        ("Ngày bản án phúc thẩm", meta.get("ngay_ba_pt", "")),
        ("Tên Thư ký viên", meta.get("ho_ten_tkv", "")),
        ("Tên Thẩm phán", meta.get("ho_ten_tham_phan", "")),
    ]
    if ds.get("nguoi_de_nghi"):
        checks.append(("Người đề nghị GĐT/TT", ds["nguoi_de_nghi"]))
    for label, val in checks:
        val_s = str(val or "").strip()
        if not val_s or val_s in ("...", "TAND...") or "[CẦN ĐIỀN" in val_s:
            continue
        count = full_text.count(val_s)
        if count < 2:
            warnings.append(f"{label} ('{val_s}') chỉ xuất hiện {count} lần — cần rà soát lại tính nhất quán.")

    return errors, warnings


def build_to_trinh(data, out_path, dry_run=False, allow_draft=False):
    if not os.path.exists(BASE_TEMPLATE):
        raise FileNotFoundError(f"Không tìm thấy phôi chuẩn: {BASE_TEMPLATE}")

    norm = ps.normalize_payload(data)
    meta = norm["metadata"]
    duong_su = norm["duong_su"]
    heading_size = int(meta.get("roman_heading_size_pt", 14))

    doc = Document(BASE_TEMPLATE)
    verify_template_dom(doc)

    # 1. Bảng 1: Quốc hiệu - Tiêu ngữ - Ngày tháng
    populate_header_table(doc, meta)

    # 2. Khối Tiêu đề & Trích yếu
    populate_title_block(doc, meta)

    # 3. Bảng 2: Bảng chia đôi (Khung TKV | Đương sự)
    populate_split_table(doc, meta, duong_su)

    # 4. Mục I. TÓM TẮT NỘI DUNG VỤ ÁN
    add_body_para(doc, "I. TÓM TẮT NỘI DUNG VỤ ÁN", size_pt=heading_size, bold=True, before_pt=12, after_pt=6)
    for block in norm.get("muc_1_tom_tat", []):
        if isinstance(block, str):
            add_body_para(doc, block, size_pt=14)
        elif isinstance(block, dict):
            if block.get("tieu_de"):
                add_body_para(doc, block["tieu_de"], size_pt=14, italic=True)
            for para in block.get("doan_van", []):
                add_body_para(doc, str(para), size_pt=14)

    # 5. Mục II. QUÁ TRÌNH GIẢI QUYẾT
    add_body_para(doc, "II. QUÁ TRÌNH GIẢI QUYẾT", size_pt=heading_size, bold=True, before_pt=6, after_pt=6)
    m2 = norm.get("muc_2_qua_trinh") or {}
    add_body_para(doc, m2.get("so_tham_tieu_de", "1. Tòa án cấp sơ thẩm tuyên xử:"), size_pt=14, bold=True)
    for qd in m2.get("so_tham_quyet_dinh", []):
        add_body_para(doc, str(qd), size_pt=14, italic=True)
    for note in m2.get("so_tham_ghi_chu", []):
        add_body_para(doc, str(note), size_pt=14)

    add_body_para(doc, m2.get("phuc_tham_tieu_de", "2. Tòa án cấp phúc thẩm tuyên xử:"), size_pt=14, bold=True)
    for qd_pt in m2.get("phuc_tham_quyet_dinh", []):
        add_body_para(doc, str(qd_pt), size_pt=14, italic=True)
    for ts in m2.get("tien_su_sau_phuc_tham", []):
        add_body_para(doc, str(ts), size_pt=14)

    # 6. Mục III. ĐỀ NGHỊ CỦA ĐƯƠNG SỰ
    add_body_para(doc, "III. ĐỀ NGHỊ CỦA ĐƯƠNG SỰ", size_pt=heading_size, bold=True, before_pt=6, after_pt=6)
    m3 = norm.get("muc_3_de_nghi") or {}
    thu_tuc_lower = meta.get("thu_tuc", "GIÁM ĐỐC THẨM").lower()
    m3_mo_dau = m3.get("mo_dau") or f"Đương sự có đơn đề nghị xem xét bản án theo thủ tục {thu_tuc_lower}."
    add_body_para(doc, m3_mo_dau, size_pt=14)
    for ly_do in m3.get("cac_ly_do", []):
        add_body_para(doc, str(ly_do), size_pt=14)
    for vp in m3.get("vi_pham_to_tung", []):
        add_body_para(doc, f"- Về tố tụng: {vp}", size_pt=14)
    for vp in m3.get("vi_pham_noi_dung", []):
        add_body_para(doc, f"- Về nội dung: {vp}", size_pt=14)

    # 7. Mục IV. ĐỀ XUẤT GIẢI QUYẾT ĐƠN ĐỀ NGHỊ
    add_body_para(doc, "IV. ĐỀ XUẤT GIẢI QUYẾT ĐƠN ĐỀ NGHỊ", size_pt=heading_size, bold=True, before_pt=6, after_pt=6)
    m4 = norm.get("muc_4_de_xuat") or {}

    add_body_para(doc, "1. Ý kiến đề xuất của cán bộ nghiên cứu:", size_pt=14, bold=True)
    for item in m4.get("noi_dung_phan_tich", []):
        if isinstance(item, str):
            m_num = re.match(r"^(\d+\))\s+(.*)$", item)
            m_sub = re.match(r"^(-\s+)([^:]+)(:\s*.*)$", item)
            if m_num:
                add_rich_para(doc, [(m_num.group(1), 14, True, False), (" " + m_num.group(2), 14, False, False)])
            elif m_sub:
                add_rich_para(doc, [(m_sub.group(1), 14, False, False), (m_sub.group(2), 14, False, True), (m_sub.group(3), 14, False, False)])
            else:
                add_body_para(doc, item, size_pt=14)
        elif isinstance(item, dict):
            add_body_para(doc, item.get("text", ""), size_pt=14, bold=item.get("bold", False), italic=item.get("italic", False))

    de_xuat_chot = m4.get("de_xuat_chot", "")
    if de_xuat_chot:
        add_rich_para(
            doc,
            [("Thư ký viên đề xuất", 14, True, False), (f": {de_xuat_chot}", 14, False, False)],
            after_pt=0 if not m4.get("kinh_trinh") else 6
        )
    if m4.get("kinh_trinh"):
        add_body_para(doc, m4["kinh_trinh"], size_pt=14, after_pt=0)

    ho_ten_tkv = meta.get("ho_ten_tkv") or "[CẦN ĐIỀN TÊN TKV]"
    cd_tp = meta.get("chuc_danh_tham_phan") or "Thẩm phán TANDTC"
    ht_tp = meta.get("ho_ten_tham_phan") or "[CẦN ĐIỀN TÊN THẨM PHÁN]"
    add_signature_table(doc, ho_ten_tkv, chuc_danh_ky=meta.get("chuc_danh_ky_tkv", "Thư ký viên"))

    nam_y_kien = meta.get("nam_trinh") or "2026"
    tieu_de_tp = f"2. Ý kiến của {cd_tp} {ht_tp} ngày      /     /{nam_y_kien}:"
    add_body_para(doc, tieu_de_tp, size_pt=14, bold=True, before_pt=12, after_pt=6)

    du_thao_tp = m4.get("du_thao_y_kien_tham_phan", [])
    if du_thao_tp:
        for para in du_thao_tp:
            add_body_para(doc, str(para), size_pt=14)
    else:
        for _ in range(3):
            add_body_para(doc, "", size_pt=14)

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
    errors, warnings = validate_document(doc, norm, allow_draft=allow_draft)
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
        raise ValueError("Dừng xuất Tờ trình do phát hiện lỗi kiểm tra dữ liệu/placeholder!")

    if dry_run:
        print("[DRY-RUN OK] Dữ liệu & định dạng Tờ trình hợp lệ 100%.")
        return

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
    parser.add_argument("--allow-draft", action="store_true", help="Cho phép xuất bản nháp khi còn thẻ [CẦN ĐIỀN...]")
    args = parser.parse_args()

    with open(args.input, "r", encoding="utf-8") as f:
        data = json.load(f)

    build_to_trinh(data, args.output, dry_run=args.dry_run, allow_draft=args.allow_draft)


if __name__ == "__main__":
    main()
