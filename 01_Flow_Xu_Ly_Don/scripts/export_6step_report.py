# -*- coding: utf-8 -*-
"""
export_6step_report.py - Bộ xuất Báo cáo Toàn diện Nghiên cứu Hồ sơ 6 bước (.docx)
chuẩn thể thức hành chính & tư pháp TANDTC trong < 1 giây.

Nhận đầu vào:
- File Markdown (.md) chứa báo cáo 6 bước, HOẶC
- File JSON (.json) chứa dữ liệu payload 6 bước.

Đặc tính định dạng:
- Lề trang chuẩn hành chính: Trái 3.0cm, Phải 2.0cm, Trên 2.0cm, Dưới 2.0cm (A4).
- Font: Times New Roman 13-14pt, giãn dòng 1.2pt.
- Header chuẩn: TÒA ÁN NHÂN DÂN TỐI CAO / VỤ GIÁM ĐỐC KIỂM TRA II & Quốc hiệu.
- Bảng biểu ma trận: Header màu Deep Navy (#003366) chữ trắng, xen kẽ dòng zebra, viền xám nhạt.
- Callout box [TRẠNG THÁI HỒ SƠ ĐÃ KHÓA]: Nền xám nhạt (#F4F6F9), viền trái xanh Navy đậm.
"""
import argparse
import json
import os
import re
import sys
import docx
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

if sys.stdout.encoding != "utf-8":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


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

FONT_NAME = "Times New Roman"
COLOR_PRIMARY = RGBColor(0, 51, 102)     # Deep Navy (#003366)
COLOR_SECONDARY = RGBColor(180, 0, 0)   # Dark Red / Highlight
COLOR_TEXT = RGBColor(30, 30, 30)       # Charcoal Body Text
COLOR_MUTED = RGBColor(90, 90, 90)      # Gray caption


def set_run_font(run, size_pt=13, bold=False, italic=False, color=COLOR_TEXT):
    run.font.name = FONT_NAME
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
        rFonts.set(qn(f"w:{attr}"), FONT_NAME)


def add_formatted_runs(p, text, default_size=13, default_color=COLOR_TEXT):
    """Phân giải chuỗi có **in đậm** và *in nghiêng* thành các run chuẩn."""
    pattern = re.compile(r"(\*\*.*?\*\*|\*.*?\*)")
    parts = pattern.split(text)
    for part in parts:
        if not part:
            continue
        if part.startswith("**") and part.endswith("**") and len(part) >= 4:
            r = p.add_run(part[2:-2])
            set_run_font(r, size_pt=default_size, bold=True, italic=False, color=default_color)
        elif part.startswith("*") and part.endswith("*") and len(part) >= 2:
            r = p.add_run(part[1:-1])
            set_run_font(r, size_pt=default_size, bold=False, italic=True, color=default_color)
        else:
            r = p.add_run(part)
            set_run_font(r, size_pt=default_size, bold=False, italic=False, color=default_color)


def add_callout(doc, text, bold_title=""):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.columns[0].width = Cm(16.0)
    
    cell = table.cell(0, 0)
    shading = parse_xml(r'<w:shd {} w:fill="F4F6F9"/>'.format(nsdecls("w")))
    cell._tc.get_or_add_tcPr().append(shading)
    
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(
        r'<w:tcBorders {}><w:top w:val="none"/><w:left w:val="single" w:sz="24" w:space="0" w:color="003366"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>'.format(nsdecls("w"))
    )
    tcPr.append(borders)
    
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.2
    
    if bold_title:
        r_title = p.add_run(bold_title + "\n")
        set_run_font(r_title, size_pt=12.5, bold=True, italic=False, color=COLOR_PRIMARY)
        
    for line in text.strip().split("\n"):
        line = line.strip()
        if not line:
            continue
        p_line = cell.add_paragraph() if p.text else p
        p_line.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_line.paragraph_format.space_after = Pt(2)
        p_line.paragraph_format.line_spacing = 1.15
        add_formatted_runs(p_line, line, default_size=12, default_color=COLOR_TEXT)
        
    p_spacer = doc.add_paragraph()
    p_spacer.paragraph_format.space_after = Pt(4)


def add_styled_table(doc, headers, data, col_widths=None):
    if not headers or not data:
        return
    num_cols = len(headers)
    table = doc.add_table(rows=len(data) + 1, cols=num_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    # Tính chiều rộng cột tự động nếu không truyền
    if not col_widths or len(col_widths) != num_cols:
        col_w = 16.0 / num_cols
        col_widths = [col_w] * num_cols

    # Header Row
    hdr_cells = table.rows[0].cells
    for i, h_text in enumerate(headers):
        hdr_cells[i].width = Cm(col_widths[i])
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        r = p.add_run(h_text.strip())
        set_run_font(r, size_pt=11.5, bold=True, italic=False, color=RGBColor(255, 255, 255))
        
        shading = parse_xml(r'<w:shd {} w:fill="003366"/>'.format(nsdecls("w")))
        hdr_cells[i]._tc.get_or_add_tcPr().append(shading)
        hdr_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # Data Rows
    for r_idx, row_values in enumerate(data):
        row_cells = table.rows[r_idx + 1].cells
        bg_color = "F9FAFC" if r_idx % 2 == 1 else "FFFFFF"
        for c_idx in range(num_cols):
            row_cells[c_idx].width = Cm(col_widths[c_idx])
            p = row_cells[c_idx].paragraphs[0]
            val = str(row_values[c_idx]).strip() if c_idx < len(row_values) else ""
            align = WD_ALIGN_PARAGRAPH.CENTER if (c_idx == 0 and len(val) < 8) else WD_ALIGN_PARAGRAPH.JUSTIFY
            p.alignment = align
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.line_spacing = 1.15
            add_formatted_runs(p, val, default_size=11, default_color=COLOR_TEXT)
            
            shading = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls("w"), bg_color))
            row_cells[c_idx]._tc.get_or_add_tcPr().append(shading)
            row_cells[c_idx].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # Viền bảng xám nhẹ (#D0D5DD)
    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            borders = parse_xml(
                r'<w:tcBorders {}><w:top w:val="single" w:sz="4" w:space="0" w:color="D0D5DD"/><w:left w:val="single" w:sz="4" w:space="0" w:color="D0D5DD"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="D0D5DD"/><w:right w:val="single" w:sz="4" w:space="0" w:color="D0D5DD"/></w:tcBorders>'.format(nsdecls("w"))
            )
            tcPr.append(borders)
            
    doc_p = table._tbl.getparent()
    p_after = docx.text.paragraph.Paragraph(OxmlElement("w:p"), doc_p)
    p_after.paragraph_format.space_after = Pt(6)


def build_doc_from_markdown(md_text, out_path, title_case="", subtitle="", dry_run=False):
    doc = docx.Document()
    
    # 1. Cấu hình lề trang A4 tiêu chuẩn (3 - 2 - 2 - 2 cm)
    for s in doc.sections:
        s.top_margin = Cm(2.0)
        s.bottom_margin = Cm(2.0)
        s.left_margin = Cm(3.0)
        s.right_margin = Cm(2.0)
        s.page_width = Cm(21.0)
        s.page_height = Cm(29.7)
        
    # 2. Header Quốc hiệu & Đơn vị (Chuẩn TANDTC)
    t_header = doc.add_table(rows=1, cols=2)
    t_header.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_header.autofit = False
    t_header.columns[0].width = Cm(7.5)
    t_header.columns[1].width = Cm(8.5)
    
    p0 = t_header.cell(0, 0).paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p0.add_run("TÒA ÁN NHÂN DÂN TỐI CAO\nVỤ GIÁM ĐỐC KIỂM TRA II\n──────────────")
    set_run_font(r1, size_pt=11, bold=True, color=COLOR_PRIMARY)
    
    p1 = t_header.cell(0, 1).paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p1.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc\n──────────────────────")
    set_run_font(r2, size_pt=11, bold=True, color=COLOR_PRIMARY)
    
    # 3. Tiêu đề Báo cáo
    p_t = doc.add_paragraph()
    p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t.paragraph_format.space_before = Pt(12)
    p_t.paragraph_format.space_after = Pt(4)
    r_t = p_t.add_run("BÁO CÁO TOÀN DIỆN KẾT QUẢ NGHIÊN CỨU HỒ SƠ VỤ ÁN\nQUY TRÌNH 6 BƯỚC XỬ LÝ ĐƠN ĐỀ NGHỊ GIÁM ĐỐC THẨM")
    set_run_font(r_t, size_pt=14.5, bold=True, color=COLOR_SECONDARY)
    
    if title_case or subtitle:
        p_sub = doc.add_paragraph()
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_sub.paragraph_format.space_after = Pt(12)
        r_sub = p_sub.add_run(f"{title_case}\n{subtitle}".strip())
        set_run_font(r_sub, size_pt=12, bold=True, italic=True, color=COLOR_PRIMARY)

    # 4. Phân tích nội dung Markdown
    lines = md_text.splitlines()
    in_table = False
    table_headers = []
    table_rows = []
    in_callout = False
    callout_lines = []
    callout_title = ""

    def flush_table():
        nonlocal in_table, table_headers, table_rows
        if in_table and table_headers and table_rows:
            add_styled_table(doc, table_headers, table_rows)
        in_table = False
        table_headers = []
        table_rows = []

    def flush_callout():
        nonlocal in_callout, callout_lines, callout_title
        if in_callout and callout_lines:
            add_callout(doc, "\n".join(callout_lines), bold_title=callout_title)
        in_callout = False
        callout_lines = []
        callout_title = ""

    for line in lines:
        stripped = line.strip()
        
        # Kiểm tra bảng markdown: | col1 | col2 |
        if stripped.startswith("|") and stripped.endswith("|"):
            flush_callout()
            cells = [c.strip() for c in stripped.split("|")[1:-1]]
            if all(c.startswith("---") or c.startswith(":-") for c in cells if c):
                # Dòng phân cách header
                continue
            if not in_table:
                in_table = True
                table_headers = cells
                table_rows = []
            else:
                table_rows.append(cells)
            continue
        else:
            flush_table()

        # Kiểm tra Callout box: > ...
        if stripped.startswith(">"):
            content = stripped[1:].strip()
            if not in_callout:
                in_callout = True
                callout_lines = []
                callout_title = ""
                if "**[" in content:
                    m = re.search(r"\*\*\[(.*?)\]\*\*", content)
                    if m:
                        callout_title = f"[{m.group(1)}]"
                        content = content.replace(m.group(0), "").strip()
            if content:
                callout_lines.append(content)
            continue
        else:
            flush_callout()

        # Tiêu đề H1 (# ...)
        if stripped.startswith("# "):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(14)
            p.paragraph_format.space_after = Pt(6)
            p.paragraph_format.keep_with_next = True
            r = p.add_run(stripped[2:].strip())
            set_run_font(r, size_pt=14, bold=True, color=COLOR_PRIMARY)
            continue

        # Tiêu đề H2 (## ...)
        if stripped.startswith("## "):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.keep_with_next = True
            r = p.add_run(stripped[3:].strip())
            set_run_font(r, size_pt=13, bold=True, color=COLOR_PRIMARY)
            continue

        # Tiêu đề H3 (### ...)
        if stripped.startswith("### "):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.keep_with_next = True
            r = p.add_run(stripped[4:].strip())
            set_run_font(r, size_pt=12.5, bold=True, color=COLOR_PRIMARY)
            continue

        # Đoạn văn thông thường
        if stripped:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent = Cm(1.0) if not stripped.startswith(("-", "•", "1.", "2.", "3.", "4.", "5.")) else Cm(0)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.line_spacing = 1.2
            add_formatted_runs(p, stripped, default_size=13, default_color=COLOR_TEXT)

    flush_table()
    flush_callout()

    if dry_run:
        print(f"[DRY-RUN OK] Báo cáo 6 bước hợp lệ 100% (chưa ghi file): {out_path}")
        return

    ok, lock_err = check_file_lock(out_path)
    if not ok:
        raise PermissionError(lock_err)

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    doc.save(out_path)
    print(f"[SUCCESS] Đã xuất Báo cáo 6 bước chuẩn đẹp tại: {out_path}")


def main():
    parser = argparse.ArgumentParser(description="Xuất file Báo cáo Toàn diện Nghiên cứu 6 Bước sang Word (.docx)")
    parser.add_argument("--input", "-i", required=True, help="Đường dẫn file .md hoặc .json")
    parser.add_argument("--output", "-o", required=True, help="Đường dẫn file .docx đầu ra")
    parser.add_argument("--title", "-t", default="", help="Tiêu đề vụ án")
    parser.add_argument("--subtitle", "-s", default="", help="Phụ đề / Thông tin bản án")
    parser.add_argument("--dry-run", action="store_true", help="Chỉ kiểm tra cấu trúc, không ghi file")
    args = parser.parse_args()

    input_path = os.path.abspath(args.input)
    if not os.path.exists(input_path):
        print(f"[ERROR] Không tìm thấy file: {input_path}")
        sys.exit(1)

    ext = os.path.splitext(input_path)[1].lower()
    if ext == ".md":
        with open(input_path, "r", encoding="utf-8") as f:
            md_content = f.read()
        build_doc_from_markdown(md_content, args.output, title_case=args.title, subtitle=args.subtitle, dry_run=args.dry_run)
    elif ext == ".json":
        with open(input_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        meta = data.get("metadata", {})
        md_parts = [
            "# THÔNG TIN ĐƯƠNG SỰ & BẢN ÁN BỊ KHIẾU NẠI",
            f"- **Vụ việc:** {meta.get('tieu_de_vu_an') or meta.get('quan_he_tranh_chap', '')}",
            f"- **Bản án phúc thẩm:** {meta.get('so_ba_pt', '')} ngày {meta.get('ngay_ba_pt') or meta.get('ngay_tuyen_pt', '')} của {meta.get('toa_an_pt', '')}",
            f"- **Bản án sơ thẩm:** {meta.get('so_ba_st', '')} ngày {meta.get('ngay_ba_st') or meta.get('ngay_tuyen_st', '')} của {meta.get('toa_an_st', '')}\n",
            "# I. TÓM TẮT NỘI DUNG VỤ ÁN",
        ]
        for block in data.get("muc_1_tom_tat", []):
            if isinstance(block, str):
                md_parts.append(block)
            elif isinstance(block, dict):
                if block.get("tieu_de"):
                    md_parts.append(f"## {block['tieu_de']}")
                for p in block.get("doan_van", []):
                    md_parts.append(str(p))

        md_parts.append("\n# II. QUÁ TRÌNH GIẢI QUYẾT TẠI CÁC CẤP TÒA ÁN")
        m2 = data.get("muc_2_qua_trinh") or data.get("muc_2_noi_dung_an") or {}
        md_parts.append(f"## {m2.get('so_tham_tieu_de', 'Cấp sơ thẩm')}")
        for p in (m2.get("so_tham_quyet_dinh") or m2.get("so_tham_noi_dung") or []):
            md_parts.append(str(p))
        md_parts.append(f"## {m2.get('phuc_tham_tieu_de', 'Cấp phúc thẩm')}")
        for p in (m2.get("phuc_tham_quyet_dinh") or m2.get("phuc_tham_noi_dung") or []):
            md_parts.append(str(p))

        md_parts.append("\n# III. NỘI DUNG ĐƠN KHIẾU NẠI & YÊU CẦU")
        m3 = data.get("muc_3_de_nghi", {})
        if m3.get("mo_dau"):
            md_parts.append(str(m3["mo_dau"]))
        for yc in (m3.get("cac_ly_do") or m3.get("cac_noi_dung_khieu_nai") or m3.get("cac_noi_dung_khiếu_nai") or []):
            md_parts.append(str(yc) if str(yc).startswith("-") else f"- {yc}")

        md_parts.append("\n# IV. KẾT QUẢ ĐÁNH GIÁ CĂN CỨ & ĐỀ XUẤT XỬ LÝ")
        m4 = data.get("muc_4_de_xuat", {})
        if m4.get("nhan_dinh_chung"):
            md_parts.append(str(m4["nhan_dinh_chung"]))
        for item in m4.get("noi_dung_phan_tich", []):
            md_parts.append(item.get("text", "") if isinstance(item, dict) else str(item))
        for sub in m4.get("cac_muc_phan_tich", []):
            md_parts.append(f"### {sub.get('tieu_de', '')}")
            for p in sub.get("doan_van", []):
                md_parts.append(str(p))
        de_xuat = m4.get("de_xuat_chot") or m4.get("ket_luan_de_xuat") or ""
        if de_xuat:
            md_parts.append(f"## Quan điểm đề xuất: {de_xuat}")

        full_md = "\n".join(md_parts)
        build_doc_from_markdown(
            full_md,
            args.output,
            title_case=meta.get("tieu_de_vu_an") or meta.get("quan_he_tranh_chap", ""),
            subtitle=f"Bản án phúc thẩm số: {meta.get('so_ba_pt', '')}",
            dry_run=args.dry_run,
        )
    else:
        print("[ERROR] Định dạng không hỗ trợ. Hãy dùng .md hoặc .json")


if __name__ == "__main__":
    main()
