# -*- coding: utf-8 -*-
"""
Script xuất Quyết định Kháng nghị GĐT / Tái thẩm của Chánh án TANDTC (.docx) chuẩn 100% Bản phát hành gốc.
Tuân thủ AGENTS.md (Bước 6):
- Mã hóa UTF-8 Unicode hoàn toàn bằng python-docx (không dùng PowerShell COM).
- Dựa trên phôi XML gốc base_khang_nghi_tandtc.docx (giữ 100% 2 đường kẻ VML, footer1/footer2, styles, margins).
- Tích hợp File-lock check, Dry-run mode, Placeholder scan và Cross-consistency check.
"""
import argparse
import json
import os
import re
import sys
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import paths_config as cfg

BASE_TEMPLATE = os.path.join(cfg.TEMPLATES_DIR, "base_khang_nghi_tandtc.docx")
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
            has_media = (
                child.find(qn("w:pict")) is not None
                or child.find(qn("w:drawing")) is not None
                or child.find(".//{urn:schemas-microsoft-com:vml}line") is not None
                or child.find(".//{urn:schemas-microsoft-com:vml}shape") is not None
            )
            if has_media:
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


def format_para(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_cm=1.0, left_cm=0,
                before_pt=5, after_pt=0, line_spacing=1.15):
    pf = p.paragraph_format
    pf.alignment = align
    pf.first_line_indent = Cm(first_line_cm)
    pf.left_indent = Cm(left_cm)
    pf.right_indent = Cm(0)
    pf.space_before = Pt(before_pt)
    pf.space_after = Pt(after_pt)
    pf.line_spacing = line_spacing
    return p


def add_kn_para(doc, text="", size_pt=14, bold=False, italic=False,
                align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_cm=1.0,
                before_pt=5, after_pt=0, line_spacing=1.15):
    p = doc.add_paragraph()
    format_para(p, align=align, first_line_cm=first_line_cm,
                before_pt=before_pt, after_pt=after_pt, line_spacing=line_spacing)
    if text:
        r = p.add_run(text)
        set_run_font(r, size_pt=size_pt, bold=bold, italic=italic)
    return p


def add_rich_kn_para(container, segments, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                     first_line_cm=1.0, left_cm=0, before_pt=5, after_pt=0, line_spacing=1.15):
    p = container.add_paragraph()
    format_para(p, align=align, first_line_cm=first_line_cm, left_cm=left_cm,
                before_pt=before_pt, after_pt=after_pt, line_spacing=line_spacing)
    for seg in segments:
        txt = seg[0]
        sz = seg[1] if len(seg) > 1 else 14
        b = seg[2] if len(seg) > 2 else False
        it = seg[3] if len(seg) > 3 else False
        r = p.add_run(txt)
        set_run_font(r, size_pt=sz, bold=b, italic=it)
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


def set_table_cols_dxa(tbl, col_dxas):
    tblPr = tbl._tbl.tblPr
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:w"), str(sum(col_dxas)))
    tblW.set(qn("w:type"), "dxa")

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


def populate_header_table(doc, meta):
    """Bảng 1: TÒA ÁN NHÂN DÂN TỐI CAO / Số: ... | Quốc hiệu - Tiêu ngữ / Ngày tháng (giữ 100% VML lines)."""
    tbl0 = doc.tables[0]
    c_left = tbl0.cell(0, 0)
    c_right = tbl0.cell(0, 1)

    # Cột trái đoạn 0: TÒA ÁN NHÂN DÂN TỐI CAO (12pt Bold) + VML shape giữ nguyên
    replace_para_runs_preserve_vml(
        c_left.paragraphs[0],
        [(meta.get("co_quan_ban_hanh", "TÒA ÁN NHÂN DÂN TỐI CAO"), 12, True, False, False)]
    )
    # Cột trái đoạn 1: \nSố: .../2026/KN-DS (14pt thường)
    so_kn = f"Số: {meta.get('so_quyet_dinh', '.../2026/KN-DS')}"
    replace_para_runs_preserve_vml(
        c_left.paragraphs[1],
        [(" ", 14, False, False, False), (so_kn, 14, False, False, True)]
    )

    # Cột phải đoạn 0: Quốc hiệu (12pt Bold) \n Tiêu ngữ (14pt Bold)
    replace_para_runs_preserve_vml(
        c_right.paragraphs[0],
        [
            ("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", 12, True, False, False),
            ("Độc lập - Tự do - Hạnh phúc", 14, True, False, True),
        ]
    )
    # Cột phải đoạn 1: Địa danh, ngày tháng năm (14pt Italic) + VML Line 5 giữ nguyên
    ngay_ky = meta.get('ngay_ky', meta.get('ngay_trinh', '...'))
    thang_ky = meta.get('thang_ky', meta.get('thang_trinh', '...'))
    nam_ky = meta.get('nam_ky', meta.get('nam_trinh', '2026'))
    ngay_str = f"{meta.get('dia_danh', 'Hà Nội')}, ngày {ngay_ky} tháng {thang_ky} năm {nam_ky}"
    replace_para_runs_preserve_vml(
        c_right.paragraphs[1],
        [(ngay_str, 14, False, True, False)]
    )


def populate_title_block(doc, meta):
    """Khối tiêu đề QUYẾT ĐỊNH / KHÁNG NGHỊ... / Đối với... / CHÁNH ÁN / TANDTC."""
    paras = doc.paragraphs
    replace_para_runs_preserve_vml(paras[2], [(meta.get("tieu_de_1", "QUYẾT ĐỊNH"), 14, True, False, False)])
    replace_para_runs_preserve_vml(paras[3], [(meta.get("tieu_de_2", "KHÁNG NGHỊ GIÁM ĐỐC THẨM"), 14, True, False, False)])

    loai_ba = meta.get('loai_ban_an_pt', 'dân sự phúc thẩm')
    so_ba = meta.get('so_ba_pt', '...')
    ngay_ba = meta.get('ngay_ba_pt', '...')
    toa_an = meta.get('toa_an_pt', 'Tòa án nhân dân...')
    d1 = meta.get("doi_voi_dong_1") or f"Đối với Bản án {loai_ba} số {so_ba} ngày {ngay_ba}"
    d2 = meta.get("doi_voi_dong_2") or f"của {toa_an}"
    replace_para_runs_preserve_vml(paras[4], [(d1, 14, True, False, False)])
    replace_para_runs_preserve_vml(paras[5], [(d2, 14, True, False, False)])
    replace_para_runs_preserve_vml(paras[6], [(meta.get("tham_quyen_dong_1", "CHÁNH ÁN"), 14, True, False, False)])
    replace_para_runs_preserve_vml(paras[7], [(meta.get("tham_quyen_dong_2", "TÒA ÁN NHÂN DÂN TỐI CAO"), 14, True, False, False)])


def add_parties_section(doc, meta, duong_su):
    """Khối Căn cứ pháp lý & Danh sách đương sự (nhãn tư cách IN NGHIÊNG, firstLine=1.0cm)."""
    can_cu = meta.get("can_cu_phap_ly") or [
        "Căn cứ vào Điều 327 và Điều 328 của Bộ luật Tố tụng dân sự;"
    ]
    for cc in can_cu:
        add_kn_para(doc, cc, size_pt=14, bold=False, italic=False,
                    first_line_cm=1.0, before_pt=0, after_pt=0, line_spacing=1.15)

    loai_va = meta.get('loai_vu_an', 'dân sự')
    qh_tc = meta.get('quan_he_tranh_chap', 'Tranh chấp')
    add_rich_kn_para(
        doc,
        [
            (f"Sau khi nghiên cứu hồ sơ vụ án {loai_va} “", 14, False, False),
            (qh_tc, 14, False, True),
            ("” giữa các đương sự:", 14, False, False),
        ],
        first_line_cm=1.0, before_pt=5, after_pt=0, line_spacing=1.15
    )

    def _get_desc(it):
        if isinstance(it, str):
            return it
        if not isinstance(it, dict):
            return str(it)
        if it.get("mo_ta"):
            return it["mo_ta"]
        parts = []
        if it.get("ho_ten"):
            parts.append(it["ho_ten"])
        if it.get("nam_sinh"):
            parts.append(f"sinh năm {it['nam_sinh']}")
        if it.get("dia_chi"):
            parts.append(f"địa chỉ: {it['dia_chi']}")
        return ", ".join(parts) if parts else str(it)

    def render_group(label, items, group_reps=None):
        if not items:
            return
        if len(items) == 1 and label in ("Nguyên đơn", "Bị đơn"):
            it = items[0]
            add_rich_kn_para(
                doc,
                [(f"- {label}:", 14, False, True), (f" {_get_desc(it)}", 14, False, False)],
                first_line_cm=1.0, before_pt=5, after_pt=0, line_spacing=1.15
            )
            for dd in it.get("dai_dien", []):
                dd_role = dd.get("vai_tro", "Người đại diện") if isinstance(dd, dict) else "Người đại diện"
                add_rich_kn_para(
                    doc,
                    [(dd_role, 14, False, True), (f": {_get_desc(dd)}", 14, False, False)],
                    first_line_cm=1.0, before_pt=5, after_pt=0, line_spacing=1.15
                )
        else:
            add_rich_kn_para(
                doc,
                [(f"- {label}:", 14, False, True)],
                first_line_cm=1.0, before_pt=5, after_pt=0, line_spacing=1.15
            )
            for idx, it in enumerate(items, 1):
                add_kn_para(doc, f"{idx}. {_get_desc(it)}", size_pt=14,
                            first_line_cm=1.0, before_pt=5, after_pt=0, line_spacing=1.15)
                for dd in it.get("dai_dien", []):
                    dd_role = dd.get("vai_tro", "Người đại diện") if isinstance(dd, dict) else "Người đại diện"
                    add_rich_kn_para(
                        doc,
                        [(dd_role, 14, False, True), (f": {_get_desc(dd)}", 14, False, False)],
                        first_line_cm=1.0, before_pt=5, after_pt=0, line_spacing=1.15
                    )
        for gdd in (group_reps or []):
            gdd_role = gdd.get("vai_tro", "Người đại diện") if isinstance(gdd, dict) else "Người đại diện"
            add_rich_kn_para(
                doc,
                [(gdd_role, 14, False, True), (f": {_get_desc(gdd)}", 14, False, False)],
                first_line_cm=1.0, before_pt=5, after_pt=0, line_spacing=1.15
            )

    render_group("Nguyên đơn", duong_su.get("nguyen_don", []), duong_su.get("dai_dien_nguyen_don", []))
    render_group("Bị đơn", duong_su.get("bi_don", []), duong_su.get("dai_dien_bi_don", []))
    render_group("Người có quyền lợi, nghĩa vụ liên quan", duong_su.get("nguoi_lien_quan", []), duong_su.get("dai_dien_nguoi_lien_quan", []))


def add_noi_nhan_and_signature_table(doc, noi_nhan, ky_duyet):
    """Bảng 2: Cột trái Nơi nhận (9.74cm, 11pt) | Cột phải Chữ ký (6.24cm, 12pt Bold / 14pt Bold)."""
    tbl = doc.add_table(rows=1, cols=2)
    set_no_borders(tbl)
    set_table_cols_dxa(tbl, [5524, 3538])

    c_left = tbl.cell(0, 0)
    c_right = tbl.cell(0, 1)

    # Cột trái: Nơi nhận (11pt, line=1.0)
    p0 = c_left.paragraphs[0]
    format_para(p0, align=WD_ALIGN_PARAGRAPH.LEFT, first_line_cm=0, before_pt=0, after_pt=0, line_spacing=1.0)
    r0 = p0.add_run("Nơi nhận:")
    set_run_font(r0, size_pt=11, bold=True, italic=True)

    for line in noi_nhan:
        p = c_left.add_paragraph()
        format_para(p, align=WD_ALIGN_PARAGRAPH.LEFT, first_line_cm=0, before_pt=0, after_pt=0, line_spacing=1.0)
        r = p.add_run(line)
        set_run_font(r, size_pt=11, bold=False, italic=False)

    # Cột phải: Chữ ký (căn giữa, left=-0.18cm, line=1.0)
    chuc_danh_lines = ky_duyet.get("chuc_danh_lines", ["KT. CHÁNH ÁN", "PHÓ CHÁNH ÁN"])
    p_first = c_right.paragraphs[0]
    format_para(p_first, align=WD_ALIGN_PARAGRAPH.CENTER, first_line_cm=0, left_cm=-0.18,
                before_pt=0, after_pt=0, line_spacing=1.0)
    rf = p_first.add_run(chuc_danh_lines[0])
    set_run_font(rf, size_pt=12, bold=True)

    for extra_cd in chuc_danh_lines[1:]:
        p = c_right.add_paragraph()
        format_para(p, align=WD_ALIGN_PARAGRAPH.CENTER, first_line_cm=0, left_cm=-0.18,
                    before_pt=0, after_pt=0, line_spacing=1.0)
        r = p.add_run(extra_cd)
        set_run_font(r, size_pt=12, bold=True)

    # 6 dòng trống ký tên & đóng dấu (đúng chuẩn Bản phát hành gốc)
    for _ in range(6):
        pe = c_right.add_paragraph()
        format_para(pe, align=WD_ALIGN_PARAGRAPH.CENTER, first_line_cm=0, left_cm=-0.18,
                    before_pt=0, after_pt=0, line_spacing=1.0)

    p_name = c_right.add_paragraph()
    format_para(p_name, align=WD_ALIGN_PARAGRAPH.CENTER, first_line_cm=0, left_cm=-0.18,
                before_pt=5, after_pt=0, line_spacing=1.0)
    rn = p_name.add_run(ky_duyet["ho_ten_nguoi_ky"])
    set_run_font(rn, size_pt=14, bold=True)


def check_file_lock(path):
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
    return "\n".join(texts)


def validate_document(doc, data):
    errors = []
    warnings = []
    full_text = collect_all_text(doc)

    placeholders = re.findall(r"\{\{[^}]+\}\}", full_text)
    if placeholders:
        errors.append(f"Còn sót placeholder chưa điền: {sorted(set(placeholders))}")

    meta = data.get("metadata", {})
    nguoi_ky = data.get("ky_duyet", {}).get("ho_ten_nguoi_ky", "")
    checks = [
        ("Số bản án phúc thẩm", meta.get("so_ba_pt", ""), 3),
        ("Ngày bản án phúc thẩm", meta.get("ngay_ba_pt", ""), 3),
        ("Tòa án phúc thẩm", meta.get("toa_an_pt", ""), 3),
    ]
    if nguoi_ky:
        checks.append(("Tên người ký", nguoi_ky, 1))
    for label, val, min_cnt in checks:
        if not val:
            continue
        count = full_text.count(val)
        if count < min_cnt:
            warnings.append(f"{label} ('{val}') xuất hiện {count} lần (kỳ vọng >= {min_cnt}) — kiểm tra lại tính nhất quán.")

    return errors, warnings


def build_khang_nghi(data, out_path, dry_run=False):
    if not os.path.exists(BASE_TEMPLATE):
        raise FileNotFoundError(f"Không tìm thấy phôi chuẩn: {BASE_TEMPLATE}")

    meta = data["metadata"]
    doc = Document(BASE_TEMPLATE)

    # 1. Bảng 1: Cơ quan / Số quyết định | Quốc hiệu / Tiêu ngữ / Ngày tháng
    populate_header_table(doc, meta)

    # 2. Tiêu đề QUYẾT ĐỊNH KHÁNG NGHỊ... / CHÁNH ÁN TANDTC
    populate_title_block(doc, meta)

    # 3. Căn cứ pháp lý & Danh sách đương sự
    add_parties_section(doc, meta, data["duong_su"])

    # 4. NHẬN THẤY:
    add_kn_para(doc, "NHẬN THẤY:", size_pt=14, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
                first_line_cm=0, before_pt=12, after_pt=0, line_spacing=1.15)
    nt = data.get("nhan_thay", [])
    if isinstance(nt, dict):
        if nt.get("mo_dau_ban_an"):
            add_kn_para(doc, nt["mo_dau_ban_an"], size_pt=14, first_line_cm=1.0, before_pt=5, after_pt=0, line_spacing=1.15)
        for qd in nt.get("quyet_dinh_ban_an", []):
            add_kn_para(doc, qd, size_pt=14, italic=True, first_line_cm=1.0, before_pt=5, after_pt=0, line_spacing=1.15)
        for dt in nt.get("dien_bien_sau_phuc_tham", []):
            add_kn_para(doc, dt, size_pt=14, first_line_cm=1.0, before_pt=5, after_pt=0, line_spacing=1.15)
    elif isinstance(nt, list):
        for p in nt:
            add_kn_para(doc, str(p), size_pt=14, first_line_cm=1.0, before_pt=5, after_pt=0, line_spacing=1.15)

    # 5. XÉT THẤY:
    add_kn_para(doc, "XÉT THẤY:", size_pt=14, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
                first_line_cm=0, before_pt=12, after_pt=12, line_spacing=1.0)
    xt = data.get("xet_thay", [])
    if isinstance(xt, list):
        for doan in xt:
            add_kn_para(doc, str(doan), size_pt=14, first_line_cm=1.0, before_pt=5, after_pt=0, line_spacing=1.15)
    elif isinstance(xt, dict):
        for k, v in xt.items():
            add_kn_para(doc, str(v), size_pt=14, first_line_cm=1.0, before_pt=5, after_pt=0, line_spacing=1.15)

    # Câu chuyển "Vì các lẽ trên;"
    p_ket = doc.add_paragraph()
    format_para(p_ket, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_cm=1.0, before_pt=5, after_pt=0, line_spacing=1.15)
    r_tab = p_ket.add_run("\tVì các lẽ trên;")
    set_run_font(r_tab, size_pt=14)

    # 6. QUYẾT ĐỊNH:
    add_kn_para(doc, "QUYẾT ĐỊNH:", size_pt=14, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
                first_line_cm=0, before_pt=12, after_pt=12, line_spacing=1.0)
    qds = data.get("quyet_dinh", [])
    if isinstance(qds, list):
        for idx, qd in enumerate(qds):
            is_last = (idx == len(qds) - 1)
            add_kn_para(doc, str(qd), size_pt=14, first_line_cm=1.0,
                        before_pt=0, after_pt=(12 if is_last else 5), line_spacing=1.15)
    elif isinstance(qds, dict):
        for k, qd in qds.items():
            add_kn_para(doc, str(qd), size_pt=14, first_line_cm=1.0, before_pt=5, after_pt=0, line_spacing=1.15)

    # 7. Bảng 2: Nơi nhận & Chữ ký
    noi_nhan = data.get("noi_nhan") or [
        "Như Điều 3;",
        "VKSNDTC (để biết);",
        "TAND cấp tỉnh;",
        "Lưu hồ sơ vụ án."
    ]
    ky_duyet = data.get("ky_duyet") or {
        "chuc_danh_lines": ["KT. CHÁNH ÁN", "PHÓ CHÁNH ÁN"],
        "ho_ten_nguoi_ky": meta.get("nguoi_ky_khang_nghi") or "[CẦN ĐIỀN TÊN NGƯỜI KÝ]"
    }
    add_noi_nhan_and_signature_table(doc, noi_nhan, ky_duyet)

    # --- VALIDATION GATE ---
    errors, warnings = validate_document(doc, data)
    print("============= BÁO CÁO KIỂM TRA XUẤT QUYẾT ĐỊNH KHÁNG NGHỊ (STEP 6 LOG) =============")
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
        print("[DRY-RUN OK] Dữ liệu & định dạng Quyết định Kháng nghị hợp lệ 100%.")
        return

    ok, lock_err = check_file_lock(out_path)
    if not ok:
        raise PermissionError(lock_err)

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    doc.save(out_path)
    print(f"[SUCCESS] Đã xuất Quyết định Kháng nghị chuẩn 100% Bản phát hành gốc tại: {out_path}")


def main():
    parser = argparse.ArgumentParser(description="Xuất Quyết định Kháng nghị TANDTC (.docx) chuẩn UTF-8")
    parser.add_argument("--input", required=True, help="Đường dẫn file JSON dữ liệu đã khóa ở Bước 5.5")
    parser.add_argument("--output", required=True, help="Đường dẫn file .docx đầu ra trong 03_Ket_qua_Xu_ly/<Số>_Vu_<ten_vu>/")
    parser.add_argument("--dry-run", action="store_true", help="Chạy kiểm tra không ghi file")
    args = parser.parse_args()

    with open(args.input, "r", encoding="utf-8") as f:
        data = json.load(f)

    build_khang_nghi(data, args.output, dry_run=args.dry_run)


if __name__ == "__main__":
    main()
