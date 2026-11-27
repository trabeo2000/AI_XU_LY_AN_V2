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
from urllib.parse import unquote

if sys.stdout.encoding != "utf-8":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import paths_config as cfg

CHECKSUM_FILE = os.path.join(cfg.LEGAL_DOCS_DIR, ".legal_checksums.json")

# Quy tắc phân loại vòng đời văn bản dựa trên tên và đường dẫn
HISTORIC_FOLDERS = ["04_01_02_Dat_dai_qua_cac_thoi_ky", "01_Dan_su_va_Hop_dong", "03_Hon_nhan_Gia_dinh_va_Thua_ke"]


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
    """Chuyển đổi file .doc cũ (Word 97-2003) sang .docx bằng Word COM nếu có."""
    docx_path = os.path.splitext(doc_path)[0] + ".docx"
    if os.path.exists(docx_path):
        return docx_path
    try:
        import win32com.client
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(os.path.abspath(doc_path))
        doc.SaveAs(os.path.abspath(docx_path), 16) # 16 = wdFormatXMLDocument
        doc.Close()
        word.Quit()
        return docx_path
    except Exception:
        # Dự phòng bằng PowerShell nếu chưa cài pywin32
        import subprocess
        ps_cmd = f"""
        $w = New-Object -ComObject Word.Application
        $w.Visible = $false
        $d = $w.Documents.Open('{os.path.abspath(doc_path)}')
        $d.SaveAs([ref]'{os.path.abspath(docx_path)}', [ref]16)
        $d.Close()
        $w.Quit()
        """
        try:
            subprocess.run(["powershell", "-NoProfile", "-Command", ps_cmd], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            if os.path.exists(docx_path):
                return docx_path
        except Exception:
            pass
    return None


def extract_text_from_docx(docx_path):
    """Trích xuất text và heading chuẩn từ file .docx."""
    import docx
    doc = docx.Document(docx_path)
    lines = []
    for p in doc.paragraphs:
        txt = p.text.strip()
        if not txt:
            continue
        # Nhận diện tiêu đề Chương / Điều / Mục
        if re.match(r"^(CHƯƠNG|Chương)\s+[IVXLCDM0-9]+", txt, re.IGNORECASE):
            lines.append(f"\n# {txt}\n")
        elif re.match(r"^(MỤC|Mục)\s+[IVXLCDM0-9]+", txt, re.IGNORECASE):
            lines.append(f"\n## {txt}\n")
        elif re.match(r"^(Điều|ĐIỀU)\s+\d+[\.:\s]", txt):
            lines.append(f"\n### {txt}\n")
        else:
            lines.append(txt)

    # Đọc bảng biểu nếu có
    for table in doc.tables:
        lines.append("\n[BẢNG BIỂU]")
        for row in table.rows:
            row_txt = [cell.text.strip().replace("\n", " ") for cell in row.cells]
            lines.append("| " + " | ".join(row_txt) + " |")
        lines.append("[HẾT BẢNG]\n")

    return "\n".join(lines)


def extract_text_from_pdf(pdf_path):
    """Trích xuất text từ file PDF (có lớp chữ) hoặc báo hiệu scan."""
    try:
        import fitz
        doc = fitz.open(pdf_path)
        pages_text = []
        is_scanned = True
        for page_num in range(len(doc)):
            page = doc[page_num]
            txt = page.get_text().strip()
            if len(txt) > 50:
                is_scanned = False
            pages_text.append(f"\n<!-- Trang {page_num + 1} -->\n" + txt)
        
        if is_scanned:
            return None # PDF scan ảnh, cần xử lý trích yếu
        return "\n".join(pages_text)
    except Exception as e:
        return None


def determine_lifecycle_status(rel_path, filename):
    """Xác định trạng thái vòng đời và tính chất áp dụng của văn bản."""
    fn_lower = filename.lower()
    
    # 1. Các luật hiện hành
    if "2024" in fn_lower or "2015" in fn_lower:
        if "đất đai 2024" in fn_lower or "dân sự 2015" in fn_lower or "tố tụng dân sự" in fn_lower:
            return "🟢 [CÒN HIỆU LỰC - VĂN BẢN HIỆN HÀNH]"
    
    # 2. Các luật lịch sử qua các thời kỳ
    for year in ["1959", "1986", "1987", "1990", "1991", "1993", "1995", "2000", "2003", "2005", "2013"]:
        if year in fn_lower:
            return f"🟡 [ĐÃ HẾT HIỆU LỰC HIỆN HÀNH - DÙNG CHO GIAO DỊCH/HÀNH VI GIAI ĐOẠN {year}]"
            
    # 3. Nghị quyết & Công văn giải đáp
    if "nghị quyết" in fn_lower or "nq" in fn_lower:
        return "🔵 [NGHỊ QUYẾT HỘI ĐỒNG THẨM PHÁN TANDTC - ÁP DỤNG TRỌNG TÂM]"
    if "công văn" in fn_lower or "giải đáp" in fn_lower:
        return "🟣 [CÔNG VĂN / GIẢI ĐÁP NGHIỆP VỤ TANDTC - THAM KHẢO ĐƯỜNG LỐI]"
    if "gđt" in fn_lower or "giám đốc thẩm" in fn_lower:
        return "⚖️ [QUYẾT ĐỊNH GIÁM ĐỐC THẨM HĐTP TANDTC - ÁN TƯƠNG TỰ]"

    return "⚪ [VĂN BẢN PHÁP LÝ NGHIỆP VỤ]"


def generate_2tier_markdown(rel_path, filename, raw_text, status_tag, in_ban_goc=True):
    """Tạo cấu trúc Markdown chuẩn 2 tầng (Header tra cứu nhanh + Toàn văn chuẩn hóa)."""
    title = os.path.splitext(filename)[0]
    link_target = f"_Ban_goc/{filename}" if in_ban_goc else filename
    
    # Quét nhanh danh sách Điều luật trong văn bản
    articles = re.findall(r"(?:###\s+)?(Điều\s+\d+[\.:\s][^\n\r]{0,80})", raw_text)
    
    # Tầng 1: Header tra cứu nhanh
    tier1 = [
        f"# 📖 {title.upper()}",
        f"> **Trạng thái:** {status_tag}  ",
        f"> **Đường dẫn file gốc:** [`{filename}`]({link_target})  ",
        f"> *File Markdown này được chuẩn hóa bởi `sync_legal_knowledge.py` phục vụ đọc siêu tốc trong quy trình xử lý án.*",
        "\n---",
        "## I. MỤC LỤC TRA CỨU NHANH (QUICK SEARCH INDEX)",
        f"- **Tổng số Điều / Mục nhận diện:** {len(articles)} điều khoản.",
    ]
    
    if articles:
        tier1.append("\n**Danh mục các Điều tiêu biểu:**")
        # In tối đa 30 điều đầu hoặc các điều quan trọng
        sample_articles = articles[:35]
        for a in sample_articles:
            clean_a = a.strip("# \t\r")
            tier1.append(f"- {clean_a}")
        if len(articles) > 35:
            tier1.append(f"- *... và {len(articles) - 35} điều khoản tiếp theo xem tại Tầng 2 bên dưới.*")

    tier1.append("\n---")
    tier1.append("## II. TOÀN VĂN VĂN BẢN (STANDARDIZED FULL TEXT)\n")

    # Tầng 2: Toàn văn
    content = "\n".join(tier1) + "\n" + raw_text
    return content


def sync_file(file_path, dry_run=False):
    """Xử lý đồng bộ 1 file văn bản sang .md."""
    filename = os.path.basename(file_path)
    base_name, ext = os.path.splitext(filename)
    ext = ext.lower()
    
    if ext not in (".pdf", ".docx", ".doc"):
        return None, "Bỏ qua (không phải file văn bản)"
    
    dir_path = os.path.dirname(file_path)
    in_ban_goc = os.path.basename(dir_path) == "_Ban_goc"
    bg_subfolder = os.path.join(dir_path, "_Ban_goc")
    
    # Tự động chuyển file gốc vào _Ban_goc nếu người dùng để ở ngoài thư mục chuyên đề
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
    checksums = load_checksums()
    
    # Kiểm tra xem file có thay đổi không
    if os.path.exists(md_path) and checksums.get(rel_path) == current_md5:
        return md_path, "Đã đồng bộ (không thay đổi)"
        
    if dry_run:
        return md_path, f"[CẦN ĐỒNG BỘ] -> Tạo {os.path.basename(md_path)}"

    # Tiến hành trích xuất nội dung
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
        # Nếu là file PDF scan hoặc không trích xuất được text tự động
        raw_text = f"> [!NOTE] File gốc `{filename}` là tài liệu scan hoặc định dạng đặc thù.\n> Để đọc chính xác 100%, hãy sử dụng lệnh:\n> `python 01_Flow_Xu_Ly_Don/scripts/fast_reader.py --input \"{rel_path}\"`"

    md_content = generate_2tier_markdown(rel_path, filename, raw_text, status_tag, in_ban_goc=in_ban_goc)
    
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(md_content)
        
    checksums[rel_path] = current_md5
    save_checksums(checksums)
    return md_path, f"[THÀNH CÔNG] Đã tạo/cập nhật {os.path.basename(md_path)}"


def scan_and_sync(dry_run=False):
    """Quét toàn bộ thư mục 04_Van_ban_phap_ly và đồng bộ."""
    print("=" * 70)
    mode_str = "KIỂM TRA TRẠNG THÁI (DRY-RUN)" if dry_run else "THỰC THI ĐỒNG BỘ TOÀN BỘ (SYNC)"
    print(f"  HỆ THỐNG ĐỒNG BỘ KHO TRI THỨC PHÁP LÝ TANDTC: {mode_str}")
    print("=" * 70)
    
    results = {"synced": 0, "skipped": 0, "needed": 0, "details": []}
    
    for root, dirs, files in os.walk(cfg.LEGAL_DOCS_DIR):
        dirs[:] = [d for d in dirs if not d.startswith(".") and d != "_Ban_goc_Luu_tru_Zip_Rar" and d != "04_03_03_Ban_goc_Luu_tru_Zip_Rar"]
        
        # Nhóm file theo base_name để ưu tiên .docx hơn .pdf và .doc
        file_map = {}
        for fn in sorted(files):
            base, ext = os.path.splitext(fn)
            ext = ext.lower()
            if ext in (".pdf", ".docx", ".doc"):
                if base not in file_map:
                    file_map[base] = fn
                elif ext == ".docx":
                    file_map[base] = fn # Ưu tiên tuyệt đối .docx, loại bỏ/bỏ qua .pdf trùng
        
        for base, fn in sorted(file_map.items()):
            full_path = os.path.join(root, fn)
            md_path, msg = sync_file(full_path, dry_run=dry_run)
            rel_fn = os.path.relpath(full_path, cfg.LEGAL_DOCS_DIR)
            print(f"  * {rel_fn:<60} : {msg}")
            results["details"].append((rel_fn, msg))
            if "THÀNH CÔNG" in msg:
                results["synced"] += 1
            elif "CẦN ĐỒNG BỘ" in msg:
                results["needed"] += 1
            else:
                results["skipped"] += 1

    # Rà soát file .md mồ côi (khi người dùng xóa file gốc)
    orphan_count = 0
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
                    orphan_count += 1

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
        md_path, msg = sync_file(args.file, dry_run=False)
        print(f"Kết quả: {msg}")
    elif args.sync:
        scan_and_sync(dry_run=False)
    else:
        # Mặc định an toàn là dry-run
        scan_and_sync(dry_run=True)


if __name__ == "__main__":
    main()
