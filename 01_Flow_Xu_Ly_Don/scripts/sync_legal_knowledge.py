# -*- coding: utf-8 -*-
"""
sync_legal_knowledge.py - Trình Quản Trị & Tự Động Đồng Bộ Kho Tri Thức Pháp Lý TANDTC.
Quản lý vòng đời văn bản (Thêm mới, sửa đổi, hết hiệu lực), trích xuất và sinh file
Markdown (.md) chuẩn 2 tầng (Header tra cứu nhanh + Toàn văn chuẩn hóa) phục vụ đọc
siêu tốc trong Quy trình 6 bước xử lý hồ sơ đơn Giám đốc thẩm / Tái thẩm.

Cách dùng:
  - Kiểm tra trạng thái kho (Dry-run, không sửa file):
      python 01_Flow_Xu_Ly_Don/scripts/sync_legal_knowledge.py --check
  - Đồng bộ toàn bộ file mới / file sửa đổi sang .md:
      python 01_Flow_Xu_Ly_Don/scripts/sync_legal_knowledge.py --sync
  - Đồng bộ riêng 1 file cụ thể:
      python 01_Flow_Xu_Ly_Don/scripts/sync_legal_knowledge.py --file "đường_dẫn"
"""
import argparse
import hashlib
import json
import os
import re
import shutil
import sys
from urllib.parse import quote

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import paths_config as cfg
import docx_utils as du

du.ensure_utf8_stdout()

CHECKSUM_FILE = os.path.join(cfg.LEGAL_DOCS_DIR, ".legal_checksums.json")

HISTORIC_YEAR_RE = re.compile(r"(?<!\d)(1959|1986|1987|1990|1991|1993|1995|2000|2003|2005|2013)(?!\d)")
CURRENT_LAW_PATTERNS = (
    ("đất đai", ("2024",)),
    ("dân sự", ("2015",)),
    ("tố tụng dân sự", ("2015", "2025", "sđbs")),
    ("hngđ", ("2014",)),
    ("hôn nhân", ("2014",)),
    ("nhà ở", ("2023",)),
    ("kinh doanh bất động sản", ("2023",)),
    ("công chứng", ("2024",)),
    ("tổ chức tín dụng", ("2024",)),
    ("doanh nghiệp", ("2020",)),
    ("thương mại", ("2005",)),
    ("tổ chức tòa án", ("2024", "2025")),
)


def get_file_md5(filepath):
    """Tính mã băm MD5 để theo dõi thay đổi nội dung file."""
    h = hashlib.md5()
    with open(filepath, "rb") as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest()


def load_checksums():
    """Tải sổ theo dõi checksum."""
    if os.path.exists(CHECKSUM_FILE):
        try:
            with open(CHECKSUM_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_checksums(data):
    """Lưu sổ theo dõi checksum."""
    with open(CHECKSUM_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def convert_doc_to_docx(doc_path):
    """Chuyển đổi file .doc cũ (Word 97-2003) sang .docx bằng Word COM nếu có (đảm bảo luôn Quit)."""
    docx_path = os.path.splitext(doc_path)[0] + ".docx"
    if os.path.exists(docx_path):
        return docx_path
    word = None
    try:
        import win32com.client
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(os.path.abspath(doc_path))
        try:
            doc.SaveAs(os.path.abspath(docx_path), 16)  # 16 = wdFormatXMLDocument
        finally:
            doc.Close(False)
        return docx_path
    except Exception:
        return None
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass


def extract_text_from_docx(docx_path):
    """Trích xuất text và heading chuẩn từ file .docx."""
    import docx
    doc = docx.Document(docx_path)
    lines = []
    for p in doc.paragraphs:
        txt = p.text.strip()
        if not txt:
            continue
        if re.match(r"^(CHƯƠNG|Chương)\s+[IVXLCDM0-9]+", txt, re.IGNORECASE):
            lines.append(f"\n# {txt}\n")
        elif re.match(r"^(MỤC|Mục)\s+[IVXLCDM0-9]+", txt, re.IGNORECASE):
            lines.append(f"\n## {txt}\n")
        elif re.match(r"^(Điều|ĐIỀU)\s+\d+[\.:\s]", txt):
            lines.append(f"\n### {txt}\n")
        else:
            lines.append(txt)

    for table in doc.tables:
        lines.append("\n[BẢNG BIỂU]")
        for row in table.rows:
            row_txt = [cell.text.strip().replace("\n", " ").replace("|", "\\|") for cell in row.cells]
            lines.append("| " + " | ".join(row_txt) + " |")
        lines.append("[HẾT BẢNG]\n")

    return "\n".join(lines)


def extract_text_from_pdf(pdf_path):
    """Trích xuất text từ file PDF có lớp chữ; tự nhận diện PDF scan theo tỷ lệ trang có chữ."""
    try:
        import fitz
        doc = fitz.open(pdf_path)
        try:
            total = len(doc)
            if total == 0:
                return None
            pages_text = []
            text_pages = 0
            for page_num in range(total):
                txt = doc[page_num].get_text().strip()
                if len(txt) > 50:
                    text_pages += 1
                pages_text.append(f"\n<!-- Trang {page_num + 1} -->\n" + txt)

            # Yêu cầu > 40% số trang có chữ (tránh nhầm PDF scan chỉ có 1 trang bìa có text)
            if (text_pages / total) <= 0.4:
                return None
            return "\n".join(pages_text)
        finally:
            doc.close()
    except Exception:
        return None


def determine_lifecycle_status(rel_path, filename):
    """
    Xác định chính xác trạng thái vòng đời văn bản theo thứ tự ưu tiên loại văn bản
    (tránh gắn nhầm Nghị quyết HĐTP năm 2003/2005 thành luật lịch sử hết hiệu lực).
    """
    fn_lower = filename.lower()
    rel_lower = rel_path.lower()

    # 1. Nghị quyết HĐTP / Quốc hội thi hành luật
    if "nghị quyết" in fn_lower or "nghị-quyết" in fn_lower or re.search(r"\bnq\b", fn_lower):
        return "🔵 [NGHỊ QUYẾT HỘI ĐỒNG THẨM PHÁN / QUỐC HỘI - ÁP DỤNG TRỌNG TÂM]"

    # 2. Công văn / Giải đáp nghiệp vụ TANDTC
    if "công văn" in fn_lower or "giải đáp" in fn_lower or "giai dap" in fn_lower:
        return "🟣 [CÔNG VĂN / GIẢI ĐÁP NGHIỆP VỤ TANDTC - THAM KHẢO ĐƯỜNG LỐI]"

    # 3. Án lệ chính thức của TANDTC
    if "án lệ" in fn_lower or "an le" in fn_lower or "an_le" in rel_lower or "04_04_an_le" in rel_lower:
        return "⚖️ [ÁN LỆ TANDTC - ÁP DỤNG BẮT BUỘC TRONG VỤ ÁN TƯƠNG TỰ]"

    # 4. Quyết định Giám đốc thẩm HĐTP TANDTC
    if "gđt" in fn_lower or "gdt" in fn_lower or "giám đốc thẩm" in fn_lower or "04_03_tuyen_tap" in rel_lower:
        return "⚖️ [QUYẾT ĐỊNH GIÁM ĐỐC THẨM HĐTP TANDTC - ÁN TƯƠNG TỰ]"

    # 4. Bộ luật / Luật hiện hành
    for law_kw, year_tokens in CURRENT_LAW_PATTERNS:
        if law_kw in fn_lower and any(tok in fn_lower for tok in year_tokens):
            return "🟢 [CÒN HIỆU LỰC - VĂN BẢN HIỆN HÀNH]"

    # 5. Bộ luật / Luật / Pháp lệnh lịch sử qua các thời kỳ
    m_year = HISTORIC_YEAR_RE.search(fn_lower)
    if m_year and any(k in fn_lower for k in ("luật", "bộ luật", "pháp lệnh")):
        return f"🟡 [ĐÃ HẾT HIỆU LỰC HIỆN HÀNH - DÙNG CHO GIAO DỊCH/HÀNH VI GIAI ĐOẠN {m_year.group(1)}]"

    return "⚪ [VĂN BẢN PHÁP LÝ NGHIỆP VỤ]"


def generate_2tier_markdown(rel_path, filename, raw_text, status_tag, in_ban_goc=True):
    """Tạo cấu trúc Markdown chuẩn 2 tầng (Header tra cứu nhanh + Toàn văn chuẩn hóa)."""
    title = os.path.splitext(filename)[0]
    encoded_fn = quote(filename, safe="._-")
    link_target = f"_Ban_goc/{encoded_fn}" if in_ban_goc else encoded_fn

    articles = re.findall(r"(?:###\s+)?(Điều\s+\d+[\.:\s][^\n\r]{0,80})", raw_text)

    tier1 = [
        f"# 📖 {title.upper()}",
        f"> **Trạng thái:** {status_tag}  ",
        f"> **Đường dẫn file gốc:** [`{filename}`]({link_target})  ",
        "> *File Markdown này được chuẩn hóa bởi `sync_legal_knowledge.py` phục vụ đọc siêu tốc trong quy trình xử lý án.*",
        "\n---",
        "## I. MỤC LỤC TRA CỨU NHANH (QUICK SEARCH INDEX)",
        f"- **Tổng số Điều / Mục nhận diện:** {len(articles)} điều khoản.",
    ]

    if articles:
        tier1.append("\n**Danh mục các Điều tiêu biểu:**")
        for a in articles[:35]:
            tier1.append(f"- {a.strip('# \t\r')}")
        if len(articles) > 35:
            tier1.append(f"- *... và {len(articles) - 35} điều khoản tiếp theo xem tại Tầng 2 bên dưới.*")

    tier1.append("\n---")
    tier1.append("## II. TOÀN VĂN VĂN BẢN (STANDARDIZED FULL TEXT)\n")

    return "\n".join(tier1) + "\n" + raw_text


def sync_file(file_path, dry_run=False, checksums=None):
    """Xử lý đồng bộ 1 file văn bản sang .md (nhận dict checksums dùng chung để tối ưu I/O)."""
    filename = os.path.basename(file_path)
    base_name, ext = os.path.splitext(filename)
    ext = ext.lower()

    if ext not in (".pdf", ".docx", ".doc"):
        return None, "Bỏ qua (không phải file văn bản)", None

    standalone = checksums is None
    if standalone:
        checksums = load_checksums()

    dir_path = os.path.dirname(file_path)
    in_ban_goc = os.path.basename(dir_path) == "_Ban_goc"
    bg_subfolder = os.path.join(dir_path, "_Ban_goc")

    if not in_ban_goc and os.path.exists(bg_subfolder) and os.path.isdir(bg_subfolder):
        if not dry_run:
            target_bg_path = os.path.join(bg_subfolder, filename)
            shutil.move(file_path, target_bg_path)
            file_path = target_bg_path
            dir_path = bg_subfolder
        in_ban_goc = True

    parent_dir = os.path.dirname(dir_path) if in_ban_goc else dir_path
    md_path = os.path.join(parent_dir, f"{base_name}.md")
    rel_path = os.path.relpath(file_path, cfg.ROOT_DIR).replace("\\", "/")
    status_tag = determine_lifecycle_status(rel_path, filename)

    current_md5 = get_file_md5(file_path)

    if os.path.exists(md_path) and checksums.get(rel_path) == current_md5:
        return md_path, "Đã đồng bộ (không thay đổi)", rel_path

    if dry_run:
        return md_path, f"[CẦN ĐỒNG BỘ] -> Tạo {os.path.basename(md_path)}", rel_path

    raw_text = None
    if ext == ".docx":
        raw_text = extract_text_from_docx(file_path)
    elif ext == ".doc":
        docx_temp = convert_doc_to_docx(file_path)
        if docx_temp and os.path.exists(docx_temp):
            raw_text = extract_text_from_docx(docx_temp)
    elif ext == ".pdf":
        raw_text = extract_text_from_pdf(file_path)

    if not raw_text:
        raw_text = (
            f"> [!NOTE] File gốc `{filename}` là tài liệu scan hoặc định dạng đặc thù.\n"
            f"> Để đọc chính xác 100%, hãy sử dụng lệnh:\n"
            f"> `python 01_Flow_Xu_Ly_Don/scripts/fast_reader.py --input \"{rel_path}\"`"
        )

    md_content = generate_2tier_markdown(rel_path, filename, raw_text, status_tag, in_ban_goc=in_ban_goc)

    with open(md_path, "w", encoding="utf-8") as f:
        f.write(md_content)

    checksums[rel_path] = current_md5
    if standalone:
        save_checksums(checksums)
    return md_path, f"[THÀNH CÔNG] Đã tạo/cập nhật {os.path.basename(md_path)}", rel_path


def scan_and_sync(dry_run=False):
    """Quét toàn bộ thư mục 04_Van_ban_phap_ly và đồng bộ (nạp/lưu checksum 1 lần duy nhất)."""
    print("=" * 70)
    mode_str = "KIỂM TRA TRẠNG THÁI (DRY-RUN)" if dry_run else "THỰC THI ĐỒNG BỘ TOÀN BỘ (SYNC)"
    print(f"  HỆ THỐNG ĐỒNG BỘ KHO TRI THỨC PHÁP LÝ TANDTC: {mode_str}")
    print("=" * 70)

    checksums = load_checksums()
    seen_rel_paths = set()
    results = {"synced": 0, "skipped": 0, "needed": 0}
    ext_priority = {".docx": 3, ".pdf": 2, ".doc": 1}

    for root, dirs, files in os.walk(cfg.LEGAL_DOCS_DIR):
        dirs[:] = [
            d for d in sorted(dirs)
            if not d.startswith(".") and d not in ("_Ban_goc_Luu_tru_Zip_Rar", "04_03_03_Ban_goc_Luu_tru_Zip_Rar")
        ]

        # Nhóm file theo base_name, ưu tiên nhất quán: .docx > .pdf > .doc
        file_map = {}
        for fn in sorted(files):
            base, ext = os.path.splitext(fn)
            ext_l = ext.lower()
            if ext_l in ext_priority:
                prev_fn = file_map.get(base)
                if prev_fn is None or ext_priority[ext_l] > ext_priority[os.path.splitext(prev_fn)[1].lower()]:
                    file_map[base] = fn

        for _, fn in sorted(file_map.items()):
            full_path = os.path.join(root, fn)
            _, msg, rel_path = sync_file(full_path, dry_run=dry_run, checksums=checksums)
            if rel_path:
                seen_rel_paths.add(rel_path)
            rel_fn = os.path.relpath(full_path, cfg.LEGAL_DOCS_DIR)
            print(f"  * {rel_fn:<60} : {msg}")
            if "THÀNH CÔNG" in msg:
                results["synced"] += 1
            elif "CẦN ĐỒNG BỘ" in msg:
                results["needed"] += 1
            else:
                results["skipped"] += 1

    # Dọn các checksum cũ của file đã bị xóa/di chuyển và lưu 1 lần duy nhất
    if not dry_run:
        pruned_checksums = {k: v for k, v in checksums.items() if k in seen_rel_paths}
        save_checksums(pruned_checksums)

    # Rà soát file .md mồ côi
    for root, dirs, files in os.walk(cfg.LEGAL_DOCS_DIR):
        if os.path.basename(root) == "_Ban_goc":
            continue
        bg_dir = os.path.join(root, "_Ban_goc")
        bg_bases = set()
        if os.path.exists(bg_dir):
            for bg_f in os.listdir(bg_dir):
                bg_bases.add(os.path.splitext(bg_f)[0])
        for fn in files:
            if fn.endswith(".md") and fn != "README.md":
                base = os.path.splitext(fn)[0]
                if bg_bases and base not in bg_bases:
                    rel_orphan = os.path.relpath(os.path.join(root, fn), cfg.LEGAL_DOCS_DIR)
                    print(f"  ⚠️ [MỒ CÔI - FILE GỐC ĐÃ BỊ XÓA]: {rel_orphan}")

    print("-" * 70)
    if dry_run:
        print(f"Tổng kết: {results['needed']} file cần tạo/đồng bộ mới, {results['skipped']} file đã chuẩn.")
        print("👉 Để thực thi đồng bộ, hãy chạy: python 01_Flow_Xu_Ly_Don/scripts/sync_legal_knowledge.py --sync")
    else:
        print(f"Tổng kết: Đã đồng bộ thành công {results['synced']} file Markdown, {results['skipped']} file giữ nguyên.")
    print("=" * 70)


def main():
    parser = argparse.ArgumentParser(description="Đồng bộ kho văn bản pháp lý sang Markdown")
    parser.add_argument("--check", action="store_true", help="Kiểm tra trạng thái (dry-run)")
    parser.add_argument("--sync", action="store_true", help="Thực thi đồng bộ toàn bộ file mới/sửa đổi")
    parser.add_argument("--file", type=str, help="Chỉ định 1 file cụ thể để đồng bộ")
    args = parser.parse_args()

    if args.file:
        if not os.path.exists(args.file):
            print(f"[LỖI] Không tìm thấy file: {args.file}")
            sys.exit(1)
        _, msg, _ = sync_file(args.file, dry_run=False)
        print(f"Kết quả: {msg}")
    elif args.sync:
        scan_and_sync(dry_run=False)
    else:
        scan_and_sync(dry_run=True)


if __name__ == "__main__":
    main()
