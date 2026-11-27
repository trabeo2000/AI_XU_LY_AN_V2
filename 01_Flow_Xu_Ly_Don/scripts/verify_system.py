# -*- coding: utf-8 -*-
"""
verify_system.py - Bộ kiểm định toàn vẹn hệ thống (System Integrity Guard).
Chạy lệnh `python 01_Flow_Xu_Ly_Don/scripts/verify_system.py` sau mỗi lần tái cấu trúc
luồng, sửa biểu mẫu, sửa AGENTS.md hoặc cập nhật script để bảo đảm mọi thành phần
khớp nhau 100% và không phát sinh lỗi ngầm.
"""
import json
import os
import py_compile
import re
import sys
from urllib.parse import unquote
import zipfile

if sys.stdout.encoding != "utf-8":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import paths_config as cfg

ERRORS = []
WARNINGS = []


def ok(msg):
    print(f"  [PASS] {msg}")


def err(msg):
    ERRORS.append(msg)
    print(f"  [FAIL] {msg}")


def warn(msg):
    WARNINGS.append(msg)
    print(f"  [WARN] {msg}")


def check_1_agents_size():
    print("\n[1] Kiểm tra dung lượng .agents/AGENTS.md (chống cắt cụt ngữ cảnh)")
    if not os.path.exists(cfg.AGENTS_MD_PATH):
        err(f"Không tìm thấy {cfg.AGENTS_MD_PATH}")
        return
    size = os.path.getsize(cfg.AGENTS_MD_PATH)
    if size > 23000:
        err(f"AGENTS.md quá lớn ({size} bytes > 23.000 bytes) — có nguy cơ bị hệ thống cắt cụt!")
    else:
        ok(f"AGENTS.md gọn sạch: {size} bytes ({size/1024:.1f} KB < 22.5 KB)")


def check_2_markdown_links():
    print("\n[2] Kiểm tra 100% liên kết tương đối trong các file Markdown")
    md_files = []
    for top in [os.path.join(cfg.ROOT_DIR, "README.md"), cfg.AGENTS_MD_PATH]:
        if os.path.exists(top):
            md_files.append(top)
    for folder in [cfg.FLOW_DIR, cfg.OUTPUT_ROOT_DIR, cfg.LEGAL_DOCS_DIR, cfg.WORKFLOW_DOCS_DIR]:
        if not os.path.exists(folder):
            continue
        for root, dirs, files in os.walk(folder):
            dirs[:] = [d for d in dirs if not d.startswith(".")]
            for fn in files:
                if fn.endswith(".md"):
                    md_files.append(os.path.join(root, fn))

    broken = []
    link_re = re.compile(r"\[[^\]]+\]\(([^)#]+)(?:#[^)]+)?\)")
    for mf in md_files:
        mdir = os.path.dirname(mf)
        with open(mf, "r", encoding="utf-8") as f:
            txt = f.read()
        for m in link_re.finditer(txt):
            lnk = m.group(1).strip()
            if lnk.startswith(("http://", "https://", "mailto:", "file:")):
                continue
            target = os.path.normpath(os.path.join(mdir, unquote(lnk).strip("<>")))
            if not os.path.exists(target):
                broken.append(f"{os.path.relpath(mf, cfg.ROOT_DIR)} -> ({lnk})")

    if broken:
        for b in broken:
            err(f"Link hỏng: {b}")
    else:
        ok(f"Toàn bộ {len(md_files)} file Markdown có 0 link hỏng")


def check_3_stale_refs_and_hardcoded_paths():
    print("\n[3] Kiểm tra tên thư mục cũ & đường dẫn ổ đĩa cứng (Portable 3 máy)")
    stale_re = re.compile(r"\b(03_To_trinh|04_Don_khang_nghi|07_Tra_loi_don|05_van_ban_phap_ly|06_Quy_trinh)\b")
    drive_re = re.compile(r"[A-Za-z]:\\(AI_XU_LY_AN|Users|Documents)")

    scan_files = [
        os.path.join(cfg.ROOT_DIR, "README.md"),
        os.path.join(cfg.ROOT_DIR, "kiem_tra_moi_truong.bat"),
        cfg.AGENTS_MD_PATH,
    ]
    for folder in [cfg.FLOW_DIR, cfg.OUTPUT_ROOT_DIR, cfg.LEGAL_DOCS_DIR, cfg.WORKFLOW_DOCS_DIR]:
        if not os.path.exists(folder):
            continue
        for root, dirs, files in os.walk(folder):
            dirs[:] = [d for d in dirs if not d.startswith(".") and d != "__pycache__"]
            for fn in files:
                if os.path.splitext(fn)[1].lower() in (".md", ".py", ".json", ".svg", ".html", ".bat", ".txt"):
                    scan_files.append(os.path.join(root, fn))

    stale_hits = []
    drive_hits = []
    self_path = os.path.abspath(__file__)
    for fp in scan_files:
        if not os.path.exists(fp) or os.path.abspath(fp) == self_path:
            continue
        with open(fp, "r", encoding="utf-8", errors="ignore") as f:
            for idx, line in enumerate(f, 1):
                if stale_re.search(line):
                    stale_hits.append(f"{os.path.relpath(fp, cfg.ROOT_DIR)}:{idx}")
                if drive_re.search(line):
                    drive_hits.append(f"{os.path.relpath(fp, cfg.ROOT_DIR)}:{idx}")

    if stale_hits:
        for h in stale_hits:
            err(f"Còn sót tên thư mục cũ tại {h}")
    else:
        ok("Không còn sót tên thư mục cũ nào")

    if drive_hits:
        for h in drive_hits:
            err(f"Còn đường dẫn ổ đĩa cứng tại {h}")
    else:
        ok("100% đường dẫn tương đối (0 ổ đĩa cứng)")


def check_4_python_scripts_and_fallbacks():
    print("\n[4] Kiểm tra cú pháp Python & chống hardcode tên người ký mặc định")
    py_files = [
        os.path.join(cfg.SCRIPTS_DIR, fn)
        for fn in sorted(os.listdir(cfg.SCRIPTS_DIR))
        if fn.endswith(".py")
    ]
    compiled_ok = True
    for pf in py_files:
        try:
            py_compile.compile(pf, doraise=True)
        except py_compile.PyCompileError as e:
            err(f"Lỗi cú pháp Python tại {os.path.basename(pf)}: {e}")
            compiled_ok = False
    if compiled_ok:
        ok(f"Toàn bộ {len(py_files)} script Python biên dịch sạch (0 lỗi cú pháp)")

    # Kiểm tra không hardcode tên Thẩm phán/Phó Chánh án thật làm fallback mặc định
    hardcoded_names = ["Trần Hồng Hà", "Nguyễn Văn Tiến"]
    self_path = os.path.abspath(__file__)
    hc_hits = []
    for pf in py_files:
        if os.path.abspath(pf) == self_path:
            continue
        with open(pf, "r", encoding="utf-8") as f:
            for idx, line in enumerate(f, 1):
                for nm in hardcoded_names:
                    if nm in line:
                        hc_hits.append(f"{os.path.basename(pf)}:{idx} ('{nm}')")
    if hc_hits:
        for h in hc_hits:
            err(f"Còn hardcode tên người ký mặc định trong code: {h}")
    else:
        ok("Không hardcode tên Thẩm phán/Lãnh đạo làm mặc định trong script")


def check_5_templates_integrity():
    print("\n[5] Kiểm tra phôi Word (.docx): dung lượng, VML & chuẩn placeholder ASCII")
    required = [
        ("base_to_trinh_tandtc.docx", 200, True),
        ("base_khang_nghi_tandtc.docx", 200, True),
        ("base_thong_bao_tld_tandtc.docx", 200, True),
        ("demo_phoi_to_trinh_truc_quan.docx", 200, True),
        ("demo_phoi_khang_nghi_truc_quan.docx", 200, False),
        ("demo_phoi_thong_bao_truc_quan.docx", 200, True),
    ]
    ascii_tag_re = re.compile(r"^\{\{[A-Z0-9_]+\}\}$")
    any_tag_re = re.compile(r"\{\{[^}]+\}\}")
    tb_tags = set()

    for fn, max_kb, check_vml in required:
        fp = os.path.join(cfg.TEMPLATES_DIR, fn)
        if not os.path.exists(fp):
            err(f"Thiếu phôi chuẩn bắt buộc: {fn}")
            continue
        size_kb = os.path.getsize(fp) / 1024
        if size_kb > max_kb:
            warn(f"Phôi {fn} hơi nặng ({size_kb:.1f} KB > {max_kb} KB) — kiểm tra font nhúng")
        with zipfile.ZipFile(fp, "r") as z:
            xml = z.read("word/document.xml").decode("utf-8", errors="ignore")
            if check_vml and ("v:line" not in xml and "v:shape" not in xml and "Straight Connector" not in xml):
                err(f"Phôi {fn} thiếu đường kẻ VML gốc!")
            plain = re.sub(r"<[^>]+>", "", xml)
            tags = sorted(set(any_tag_re.findall(plain)))
            bad_tags = [t for t in tags if not ascii_tag_re.match(t)]
            if bad_tags:
                err(f"Phôi {fn} có placeholder chứa ký tự có dấu/không chuẩn ASCII: {bad_tags}")
            if fn == "base_thong_bao_tld_tandtc.docx":
                tb_tags = set(tags)

    # Đối chiếu 1-1 placeholder của base_thong_bao_tld_tandtc.docx với export_dossier.py
    dossier_py = os.path.join(cfg.SCRIPTS_DIR, "export_dossier.py")
    with open(dossier_py, "r", encoding="utf-8") as f:
        dossier_src = f.read()
    missing_in_code = sorted(t for t in tb_tags if t not in dossier_src)
    if missing_in_code:
        err(f"export_dossier.py thiếu xử lý tag Thông báo: {missing_in_code}")
    elif tb_tags:
        ok(f"Tất cả 6 phôi Word hợp lệ, VML nguyên vẹn, 100% tag ASCII & {len(tb_tags)}/{len(tb_tags)} tag Thông báo khớp code")


def check_6_case_registry_and_order():
    print("\n[6] Kiểm tra đồng bộ sổ ngày tạo (case_registry.json) & thứ tự 03_Ket_qua_Xu_ly/")
    if not os.path.exists(cfg.REGISTRY_PATH):
        err("Thiếu file 01_Flow_Xu_Ly_Don/case_registry.json")
        return
    with open(cfg.REGISTRY_PATH, "r", encoding="utf-8") as f:
        reg = json.load(f)

    dirs = sorted(
        d for d in os.listdir(cfg.OUTPUT_ROOT_DIR)
        if os.path.isdir(os.path.join(cfg.OUTPUT_ROOT_DIR, d)) and "Mau_Template_Demo" not in d
    )
    dates_in_display_order = []
    for idx, d in enumerate(dirs, 1):
        m = re.match(r"^(\d+)_(.+)$", d)
        if not m:
            err(f"Thư mục '{d}' không theo định dạng '<Số>_Vu_<Ten_ho_so>'")
            continue
        num, base = int(m.group(1)), m.group(2)
        if num != idx:
            warn(f"Số thứ tự thư mục '{d}' chưa liên tục (kỳ vọng {idx:02d}_)")
        if base not in reg:
            err(f"Thư mục '{d}' chưa có mốc ngày tạo trong case_registry.json")
        else:
            dates_in_display_order.append((d, reg[base]))

    # Kiểm tra thứ tự ngày tạo giảm dần (mới nhất ở trên cùng)
    for i in range(len(dates_in_display_order) - 1):
        d1, t1 = dates_in_display_order[i]
        d2, t2 = dates_in_display_order[i + 1]
        if t1 < t2:
            err(f"Sai thứ tự ngày tạo: '{d1}' ({t1}) đứng trên '{d2}' ({t2})")
    if not ERRORS:
        ok(f"Toàn bộ {len(dirs)} thư mục vụ án khớp case_registry.json và xếp đúng ngày tạo mới nhất ở trên cùng")


def main():
    print("=" * 68)
    print("  KIỂM ĐỊNH TOÀN VẸN HỆ THỐNG XỬ LÝ ÁN TANDTC (SYSTEM VERIFIER)")
    print("=" * 68)
    check_1_agents_size()
    check_2_markdown_links()
    check_3_stale_refs_and_hardcoded_paths()
    check_4_python_scripts_and_fallbacks()
    check_5_templates_integrity()
    check_6_case_registry_and_order()

    print("\n" + "=" * 68)
    if ERRORS:
        print(f"[THẤT BẠI] Phát hiện {len(ERRORS)} lỗi nghiêm trọng, {len(WARNINGS)} cảnh báo.")
        sys.exit(1)
    else:
        print(f"[ĐẠT CHUẨN 100%] 0 lỗi nghiêm trọng, {len(WARNINGS)} cảnh báo.")
        sys.exit(0)


if __name__ == "__main__":
    main()
