# -*- coding: utf-8 -*-
"""
fast_reader.py - Bộ công cụ đọc & trích xuất siêu tốc cho mọi định dạng hồ sơ:
- Word (.docx): Trích xuất Header/Footer, văn bản, Textbox và bảng biểu lồng đa cấp (chống lặp ô gộp).
- Excel (.xlsx): Trích xuất các sheet (có kiểm tra tên sheet & giới hạn --max-rows).
- PDF có text layer & Hybrid PDF: Trích xuất text siêu tốc, tự động cảnh báo từng trang scan ảnh
  ẩn bên trong PDF gộp (Hybrid PDF) để không bỏ sót chứng cứ viết tay/scan.
- PDF Scan (ảnh): Ghép trang giữ đúng tỷ lệ (hỗ trợ trộn khổ dọc A4 + sơ đồ ngang A3),
  lưu cache chống đè tên bằng hash đường dẫn, có cờ --clean-cache.
"""
import argparse
import hashlib
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import paths_config as cfg
import docx_utils as du

du.ensure_utf8_stdout()

SCRIPTS_DIR = cfg.SCRIPTS_DIR
DEFAULT_CACHE_ROOT = cfg.RENDERED_CACHE_DIR


def _escape_md_cell(val):
    return str(val or "").replace("\r", "").replace("\n", " ").replace("|", "\\|").strip()


def _extract_table_markdown(tbl, depth=0):
    """Trích xuất bảng Word sang Markdown, chống lặp ô gộp ngang và hỗ trợ bảng lồng."""
    lines = ["\n[BẢNG BIỂU]" if depth == 0 else "\n[BẢNG LỒNG]"]
    rows_data = []
    nested_blocks = []

    for row in tbl.rows:
        row_vals = []
        seen_tcs = set()
        for cell in row.cells:
            tc = cell._tc
            if tc in seen_tcs:
                continue
            seen_tcs.add(tc)
            cell_paras = [p.text.strip() for p in cell.paragraphs if p.text.strip()]
            row_vals.append(_escape_md_cell(" ".join(cell_paras)))
            for ntbl in cell.tables:
                nested_blocks.append(_extract_table_markdown(ntbl, depth=depth + 1))
        if any(row_vals):
            rows_data.append(row_vals)

    if rows_data:
        max_cols = max(len(r) for r in rows_data)
        norm_rows = [r + [""] * (max_cols - len(r)) for r in rows_data]
        lines.append("| " + " | ".join(norm_rows[0]) + " |")
        lines.append("| " + " | ".join(["---"] * max_cols) + " |")
        for r in norm_rows[1:]:
            lines.append("| " + " | ".join(r) + " |")

    for nb in nested_blocks:
        lines.append(nb)
    lines.append("[HẾT BẢNG]\n" if depth == 0 else "[HẾT BẢNG LỒNG]\n")
    return "\n".join(lines)


def _extract_textbox_lines(element):
    """Trích xuất văn bản bên trong Textbox (w:txbxContent) nếu có."""
    from docx.oxml.ns import qn
    texts = []
    for txbx in element.findall(".//" + qn("w:txbxContent")):
        for p_el in txbx.findall(".//" + qn("w:p")):
            t_parts = [t.text for t in p_el.findall(".//" + qn("w:t")) if t.text]
            joined = "".join(t_parts).strip()
            if joined:
                texts.append(joined)
    return texts


def read_docx(file_path):
    import docx
    doc = docx.Document(file_path)
    output = [f"# NỘI DUNG TÀI LIỆU WORD: {os.path.basename(file_path)}\n"]

    # 1. Header / Footer (nếu có nội dung)
    seen_hf = set()
    hf_lines = []
    for sec in doc.sections:
        for label, hf in (("Header", sec.header), ("Footer", sec.footer)):
            if hf is None:
                continue
            hf_el = hf._element
            if hf_el in seen_hf:
                continue
            seen_hf.add(hf_el)
            txt = " | ".join(p.text.strip() for p in hf.paragraphs if p.text.strip())
            if txt:
                hf_lines.append(f"- [{label}]: {txt}")
    if hf_lines:
        output.append("[ĐẦU TRANG / CHÂN TRANG]")
        output.extend(hf_lines)
        output.append("")

    # 2. Đọc tuần tự đoạn văn, bảng (kể cả bảng lồng) và Textbox
    seen_txbx = set()
    for element in doc.element.body:
        tag = element.tag.split("}")[-1] if "}" in element.tag else element.tag
        if tag == "p":
            p = docx.text.paragraph.Paragraph(element, doc)
            text = p.text.strip()
            if text:
                output.append(text)
            for tb_line in _extract_textbox_lines(element):
                if tb_line not in seen_txbx:
                    seen_txbx.add(tb_line)
                    output.append(f"[TEXTBOX] {tb_line}")
        elif tag == "tbl":
            tbl = docx.table.Table(element, doc)
            output.append(_extract_table_markdown(tbl, depth=0))

    return "\n".join(output)


def read_xlsx(file_path, sheet_name=None, max_rows=500):
    import openpyxl
    wb = openpyxl.load_workbook(file_path, data_only=True)
    try:
        output = [f"# NỘI DUNG FILE EXCEL: {os.path.basename(file_path)}\n"]

        if sheet_name:
            if sheet_name not in wb.sheetnames:
                raise ValueError(
                    f"Không tìm thấy sheet '{sheet_name}'. Các sheet hiện có: {wb.sheetnames}"
                )
            target_sheets = [wb[sheet_name]]
        else:
            target_sheets = wb.worksheets

        for sheet in target_sheets:
            output.append(f"## Sheet: {sheet.title} (Tổng {sheet.max_row} dòng, {sheet.max_column} cột)")
            clean_rows = []
            for r in sheet.iter_rows(values_only=True):
                if any(cell is not None and str(cell).strip() != "" for cell in r):
                    clean_rows.append([_escape_md_cell(c) if c is not None else "" for c in r])

            if not clean_rows:
                output.append("*(Sheet rỗng)*\n")
                continue

            display_rows = clean_rows[:max_rows]
            header = display_rows[0]
            output.append("| " + " | ".join(header) + " |")
            output.append("| " + " | ".join(["---"] * len(header)) + " |")
            for r in display_rows[1:]:
                output.append("| " + " | ".join(r) + " |")

            if len(clean_rows) > max_rows:
                output.append(f"\n*(Đã cắt bớt {len(clean_rows) - max_rows} dòng phía dưới — dùng `--max-rows` để xem thêm)*")
            output.append("\n")

        return "\n".join(output)
    finally:
        wb.close()


def parse_page_ranges(range_str, total_pages):
    """Phân giải chuỗi range như '1-3,160-166' thành danh sách 0-indexed page numbers."""
    pages = set()
    for part in range_str.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            s_str, e_str = part.split("-", 1)
            s = max(1, int(s_str.strip()))
            e = min(total_pages, int(e_str.strip()))
            if s > e:
                s, e = e, s
            for p in range(s, e + 1):
                pages.add(p - 1)
        else:
            p = int(part.strip())
            if 1 <= p <= total_pages:
                pages.add(p - 1)
    return sorted(pages)


def _compress_page_numbers(one_indexed_pages):
    """Gom danh sách trang [5,6,7,12,13] thành chuỗi gọn '5-7,12-13'."""
    if not one_indexed_pages:
        return ""
    pages = sorted(set(one_indexed_pages))
    ranges = []
    start = prev = pages[0]
    for p in pages[1:]:
        if p == prev + 1:
            prev = p
        else:
            ranges.append(f"{start}-{prev}" if start != prev else str(start))
            start = prev = p
    ranges.append(f"{start}-{prev}" if start != prev else str(start))
    return ",".join(ranges)


def read_pdf(file_path, pages_str=None, search_keyword=None, render_images=False,
             stitch=2, dpi=135, out_dir=None):
    import fitz  # pymupdf
    doc = fitz.open(file_path)
    try:
        total_pages = len(doc)
        if total_pages == 0:
            return f"[WARN] File PDF rỗng (0 trang): {file_path}"

        if pages_str:
            target_pages = parse_page_ranges(pages_str, total_pages)
        else:
            target_pages = list(range(total_pages))

        # 1. Kiểm tra tỷ lệ lớp chữ trên phạm vi xử lý (chính xác cả với Hybrid PDF)
        sample_indices = target_pages if len(target_pages) <= 40 else (
            target_pages[:15] + target_pages[len(target_pages) // 2 - 5 : len(target_pages) // 2 + 5] + target_pages[-15:]
        )
        text_count = sum(1 for pno in sample_indices if len(doc[pno].get_text().strip()) > 40)
        has_text_layer = (text_count / max(1, len(sample_indices))) > 0.4

        output = [
            f"================ BÁO CÁO PHÂN TÍCH PDF: {os.path.basename(file_path)} ================",
            f"- Tổng số trang : {total_pages} trang",
            f"- Định dạng     : {'PDF có lớp chữ số (Digital/Hybrid PDF)' if has_text_layer else 'PDF Scan ảnh (Scanned Document)'}",
            f"- Phạm vi xử lý : {len(target_pages)} trang ({pages_str or 'Toàn bộ'})",
        ]

        # Tìm kiếm từ khóa trong lớp chữ
        if search_keyword:
            kw = search_keyword.strip().lower()
            output.append(f"\n--- KẾT QUẢ TÌM KIẾM TỪ KHÓA: '{search_keyword}' ---")
            matches = 0
            scanned_pages = []
            for pno in range(total_pages):
                txt = doc[pno].get_text()
                if len(txt.strip()) <= 40 and doc[pno].get_images():
                    scanned_pages.append(pno + 1)
                if kw in txt.lower():
                    matches += 1
                    idx = txt.lower().find(kw)
                    snippet = txt[max(0, idx - 100):min(len(txt), idx + 200)].replace("\n", " ")
                    output.append(f"- Trang {pno + 1}: ...{snippet}...")
            if matches == 0:
                output.append("- Không tìm thấy từ khóa trong lớp chữ (hoặc file là PDF scan).")
            if scanned_pages:
                output.append(
                    f"\n[CẢNH BÁO HYBRID PDF] Có {len(scanned_pages)} trang scan ảnh không tìm được bằng từ khóa: "
                    f"trang {_compress_page_numbers(scanned_pages)}."
                )
            return "\n".join(output)

        # Trích xuất lớp chữ số (kèm phát hiện từng trang scan ẩn bên trong Hybrid PDF)
        if has_text_layer and not render_images:
            output.append("\n--- NỘI DUNG TRÍCH XUẤT TỪ LỚP CHỮ SỐ ---")
            scanned_in_target = []
            for pno in target_pages:
                page = doc[pno]
                txt = page.get_text().strip()
                output.append(f"\n=== [Trang {pno + 1} / {total_pages}] ===")
                if len(txt) <= 40 and page.get_images():
                    scanned_in_target.append(pno + 1)
                    output.append(txt if txt else "*(Trang scan ảnh không có lớp chữ — cần render ảnh để đọc!)*")
                else:
                    output.append(txt if txt else "*(Trang trống)*")

            if scanned_in_target:
                compact = _compress_page_numbers(scanned_in_target)
                output.append(
                    f"\n[CẢNH BÁO HYBRID PDF] Phát hiện {len(scanned_in_target)} trang scan ảnh ẩn trong phạm vi đọc "
                    f"(trang {compact}). Để không bỏ sót chứng cứ, hãy chạy thêm:\n"
                    f"  python 01_Flow_Xu_Ly_Don/scripts/fast_reader.py --input \"{file_path}\" --render --pages \"{compact}\" --stitch 2"
                )
            return "\n".join(output)

        # Render ảnh vào .temp/rendered_cache/<safe_name>_<hash8>/ (chống đè tên tuyệt đối)
        if not out_dir:
            base_name = os.path.splitext(os.path.basename(file_path))[0].strip()
            safe_name = du.slugify_vn(base_name)[:32] or "pdf_pages"
            path_hash = hashlib.md5(os.path.abspath(file_path).lower().encode("utf-8")).hexdigest()[:8]
            out_dir = os.path.join(DEFAULT_CACHE_ROOT, f"{safe_name}_{path_hash}")
        os.makedirs(out_dir, exist_ok=True)

        output.append(f"- Thư mục ảnh   : {out_dir}")
        rendered_files = []

        if stitch > 1:
            i = 0
            while i < len(target_pages):
                chunk = target_pages[i:i + stitch]
                widths = [doc[pno].rect.width for pno in chunk]
                heights = [doc[pno].rect.height for pno in chunk]
                total_w = sum(widths)
                max_h = max(heights)

                comb_doc = fitz.open()
                try:
                    comb_page = comb_doc.new_page(width=total_w, height=max_h)
                    page_labels = []
                    x_cursor = 0.0
                    for pno, pw, ph in zip(chunk, widths, heights):
                        target_rect = fitz.Rect(x_cursor, 0, x_cursor + pw, ph)
                        comb_page.show_pdf_page(target_rect, doc, pno)
                        x_cursor += pw
                        page_labels.append(str(pno + 1))

                    pix = comb_page.get_pixmap(dpi=dpi)
                    tag_name = "_".join(page_labels)
                    out_file = os.path.join(out_dir, f"pages_{tag_name}.png")
                    pix.save(out_file)
                    rendered_files.append((out_file, f"Trang {', '.join(page_labels)}"))
                finally:
                    comb_doc.close()
                i += stitch
        else:
            for pno in target_pages:
                pix = doc[pno].get_pixmap(dpi=dpi)
                out_file = os.path.join(out_dir, f"page_{pno + 1:03d}.png")
                pix.save(out_file)
                rendered_files.append((out_file, f"Trang {pno + 1}"))

        output.append(f"\n[THÀNH CÔNG] Đã render {len(rendered_files)} file ảnh tối ưu (DPI={dpi}, Ghép {stitch} trang/ảnh):")
        for fpath, label in rendered_files:
            output.append(f"  + {label}: {fpath}")

        output.append("\n[GỢI Ý ĐỌC CHO AI AGENT]: Hãy gọi lệnh đọc các file ảnh trên ĐỒNG THỜI trong 1 turn (batch read) để tiết kiệm 80% thời gian!")
        return "\n".join(output)
    finally:
        doc.close()


def clean_rendered_cache():
    """Dọn sạch toàn bộ ảnh render tạm trong scripts/.temp/rendered_cache/."""
    if os.path.exists(DEFAULT_CACHE_ROOT):
        shutil.rmtree(DEFAULT_CACHE_ROOT, ignore_errors=True)
        print(f"[CLEANUP] Đã dọn sạch cache ảnh render: {DEFAULT_CACHE_ROOT}")
    else:
        print("[INFO] Thư mục cache ảnh render hiện đang trống.")


def main():
    parser = argparse.ArgumentParser(description="fast_reader: Bộ đọc siêu tốc Word, Excel, PDF và Ảnh")
    parser.add_argument("--input", "-i", help="Đường dẫn file cần đọc (.docx, .xlsx, .pdf, ảnh)")
    parser.add_argument("--pages", "-p", help="Dải trang PDF cần đọc (VD: '1-3,160-166')")
    parser.add_argument("--search", "-s", help="Tìm kiếm từ khóa trong PDF")
    parser.add_argument("--render", "-r", action="store_true", help="Ép render ảnh cho PDF")
    parser.add_argument("--stitch", type=int, default=2, help="Số trang ghép vào 1 ảnh (mặc định 2)")
    parser.add_argument("--dpi", type=int, default=135, help="Độ phân giải DPI tối ưu (mặc định 135)")
    parser.add_argument("--out-dir", "-o", help="Thư mục lưu ảnh render")
    parser.add_argument("--sheet", help="Tên sheet Excel cần đọc")
    parser.add_argument("--max-rows", type=int, default=500, help="Số dòng tối đa mỗi sheet Excel (mặc định 500)")
    parser.add_argument("--clean-cache", action="store_true", help="Dọn sạch thư mục .temp/rendered_cache/")
    args = parser.parse_args()

    if args.clean_cache:
        clean_rendered_cache()
        if not args.input:
            return

    if not args.input:
        parser.error("Cần truyền --input <file> (hoặc dùng --clean-cache).")

    input_path = os.path.abspath(args.input)
    if not os.path.exists(input_path):
        print(f"[ERROR] Không tìm thấy file: {input_path}")
        sys.exit(1)

    ext = os.path.splitext(input_path)[1].lower()

    if ext == ".docx":
        print(read_docx(input_path))
    elif ext == ".doc":
        print(
            f"[ERROR] File '{os.path.basename(input_path)}' là định dạng Word 97-2003 (.doc) mà python-docx không đọc trực tiếp được.\n"
            "👉 Vui lòng mở bằng Word và Lưu thành .docx hoặc .pdf trước khi đọc."
        )
        sys.exit(1)
    elif ext == ".xlsx":
        print(read_xlsx(input_path, sheet_name=args.sheet, max_rows=args.max_rows))
    elif ext == ".xls":
        print(
            f"[ERROR] File '{os.path.basename(input_path)}' là định dạng Excel 97-2003 (.xls) mà openpyxl không đọc trực tiếp được.\n"
            "👉 Vui lòng mở bằng Excel và Lưu thành .xlsx trước khi đọc."
        )
        sys.exit(1)
    elif ext == ".pdf":
        print(read_pdf(
            input_path,
            pages_str=args.pages,
            search_keyword=args.search,
            render_images=args.render,
            stitch=args.stitch,
            dpi=args.dpi,
            out_dir=args.out_dir,
        ))
    else:
        print(f"[INFO] Định dạng {ext} — nếu là file ảnh, hãy mở trực tiếp bằng công cụ đọc ảnh (read_file).")


if __name__ == "__main__":
    main()
