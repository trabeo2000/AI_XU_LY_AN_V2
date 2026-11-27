# -*- coding: utf-8 -*-
"""
verify_system.py - Bộ kiểm định toàn vẹn hệ thống (System Integrity Guard).
Chạy lệnh `python 01_Flow_Xu_Ly_Don/scripts/verify_system.py` sau mỗi lần tái cấu trúc
luồng, sửa biểu mẫu, sửa AGENTS.md hoặc cập nhật script để bảo đảm mọi thành phần
khớp nhau 100% và không phát sinh lỗi ngầm.
"""
import contextlib
import io
import json
import os
import py_compile
import re
import sys
from urllib.parse import unquote
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import paths_config as cfg
import docx_utils as du

du.ensure_utf8_stdout()

ERRORS = []
WARNINGS = []


def ok(msg):
    print(f"  [PASS] {msg}")


def err(msg):
    ERRORS.append(msg)
    print(f"  [FAIL] {msg}", file=sys.__stdout__)


def warn(msg):
    WARNINGS.append(msg)
    print(f"  [WARN] {msg}", file=sys.__stdout__)


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
    stale_re = re.compile(
        r"\b(03_To_trinh|04_Don_khang_nghi|07_Tra_loi_don|05_van_ban_phap_ly|06_Quy_trinh"
        r"|01_Dan_su_va_Hop_dong|03_Hon_nhan_Gia_dinh_va_Thua_ke)\b"
    )
    drive_re = re.compile(r"[A-Za-z]:[\\/](AI_XU_LY_AN|Users|Documents|Desktop|OneDrive|Downloads)")

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

    hardcoded_names = ["Trần Hồng Hà", "Nguyễn Văn Tiến", "Đào Thị Minh Thủy", "Phan Thị Vân Hương", "Phan Hương Trà"]
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
        ok("Không hardcode tên Thẩm phán/Lãnh đạo/TKV làm mặc định trong script")


def check_5_templates_integrity():
    print("\n[5] Kiểm tra phôi Word (.docx): dung lượng, VML, cấu trúc DOM & chuẩn placeholder ASCII")
    from docx import Document
    import export_to_trinh
    import export_khang_nghi

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
    local_errs = 0

    for fn, max_kb, check_vml in required:
        fp = os.path.join(cfg.TEMPLATES_DIR, fn)
        if not os.path.exists(fp):
            err(f"Thiếu phôi chuẩn bắt buộc: {fn}")
            local_errs += 1
            continue
        size_kb = os.path.getsize(fp) / 1024
        if size_kb > max_kb:
            warn(f"Phôi {fn} hơi nặng ({size_kb:.1f} KB > {max_kb} KB) — kiểm tra font nhúng")
        with zipfile.ZipFile(fp, "r") as z:
            xml = z.read("word/document.xml").decode("utf-8", errors="ignore")
            if check_vml and ("v:line" not in xml and "v:shape" not in xml and "Straight Connector" not in xml):
                err(f"Phôi {fn} thiếu đường kẻ VML gốc!")
                local_errs += 1
            plain = re.sub(r"<[^>]+>", "", xml)
            tags = sorted(set(any_tag_re.findall(plain)))
            bad_tags = [t for t in tags if not ascii_tag_re.match(t)]
            if bad_tags:
                err(f"Phôi {fn} có placeholder chứa ký tự có dấu/không chuẩn ASCII: {bad_tags}")
                local_errs += 1
            if fn == "base_thong_bao_tld_tandtc.docx":
                tb_tags = set(tags)

    # Kiểm định bất biến cấu trúc DOM của 2 phôi XML gốc
    try:
        export_to_trinh.verify_template_dom(Document(cfg.BASE_TO_TRINH_TEMPLATE))
        export_khang_nghi.verify_template_dom(Document(cfg.BASE_KHANG_NGHI_TEMPLATE))
    except Exception as dom_err:
        err(f"Lỗi cấu trúc DOM phôi gốc: {dom_err}")
        local_errs += 1

    # Đối chiếu 1-1 placeholder của base_thong_bao_tld_tandtc.docx với export_dossier.py
    dossier_py = os.path.join(cfg.SCRIPTS_DIR, "export_dossier.py")
    with open(dossier_py, "r", encoding="utf-8") as f:
        dossier_src = f.read()
    missing_in_code = sorted(t for t in tb_tags if t not in dossier_src)
    if missing_in_code:
        err(f"export_dossier.py thiếu xử lý tag Thông báo: {missing_in_code}")
        local_errs += 1

    if local_errs == 0 and tb_tags:
        ok(
            f"Tất cả 6 phôi Word hợp lệ, VML & DOM nguyên vẹn, 100% tag ASCII & "
            f"{len(tb_tags)}/{len(tb_tags)} tag Thông báo khớp code"
        )


def check_6_case_registry_and_order():
    print("\n[6] Kiểm tra đồng bộ sổ ngày tạo (case_registry.json) & thứ tự 03_Ket_qua_Xu_ly/")
    if not os.path.exists(cfg.REGISTRY_PATH):
        err("Thiếu file 01_Flow_Xu_Ly_Don/case_registry.json")
        return
    with open(cfg.REGISTRY_PATH, "r", encoding="utf-8") as f:
        reg = json.load(f)

    dirs = sorted(
        d for d in os.listdir(cfg.OUTPUT_ROOT_DIR)
        if os.path.isdir(os.path.join(cfg.OUTPUT_ROOT_DIR, d))
        and "Mau_Template_Demo" not in d
        and not d.startswith(".__")
    )
    local_errs = 0
    dates_in_display_order = []
    for idx, d in enumerate(dirs, 1):
        m = re.match(r"^(\d+)_(.+)$", d)
        if not m:
            err(f"Thư mục '{d}' không theo định dạng '<Số>_Vu_<Ten_ho_so>'")
            local_errs += 1
            continue
        num, base = int(m.group(1)), m.group(2)
        if num != idx:
            warn(f"Số thứ tự thư mục '{d}' chưa liên tục (kỳ vọng {idx:02d}_)")
        if base not in reg:
            err(f"Thư mục '{d}' chưa có mốc ngày tạo trong case_registry.json")
            local_errs += 1
        else:
            dates_in_display_order.append((d, reg[base]))

    for i in range(len(dates_in_display_order) - 1):
        d1, t1 = dates_in_display_order[i]
        d2, t2 = dates_in_display_order[i + 1]
        if t1 < t2:
            err(f"Sai thứ tự ngày tạo: '{d1}' ({t1}) đứng trên '{d2}' ({t2})")
            local_errs += 1

    if local_errs == 0:
        ok(f"Toàn bộ {len(dirs)} thư mục vụ án khớp case_registry.json và xếp đúng ngày tạo mới nhất ở trên cùng")


def check_7_dry_run_smoke_tests():
    print("\n[7] Kiểm thử tự động end-to-end (Dry-run, tương thích chéo Schema & bảo toàn định dạng)")
    from docx import Document
    import export_to_trinh
    import export_khang_nghi
    import export_6step_report
    import export_dossier
    import payload_schema as ps
    import sync_legal_knowledge as slk

    demo_tt_path = os.path.join(cfg.SCRIPTS_DIR, "demo_payload.json")
    demo_kn_path = os.path.join(cfg.SCRIPTS_DIR, "demo_khang_nghi_payload.json")
    local_errs = 0

    with open(demo_tt_path, "r", encoding="utf-8") as f:
        tt_data = json.load(f)
    with open(demo_kn_path, "r", encoding="utf-8") as f:
        kn_data = json.load(f)

    sink = io.StringIO()
    with contextlib.redirect_stdout(sink):
        # 7.1 Tờ trình (Toàn diện + Rút gọn TB 383/679)
        try:
            export_to_trinh.build_to_trinh(tt_data, "_smoke_tt.docx", dry_run=True)
            tt_rg = json.loads(json.dumps(tt_data))
            tt_rg["metadata"]["che_do_rut_gon"] = True
            export_to_trinh.build_to_trinh(tt_rg, "_smoke_tt_rg.docx", dry_run=True)
        except Exception as e:
            err(f"Smoke test Tờ trình thất bại: {e}")
            local_errs += 1

        # 7.2 Quyết định Kháng nghị với demo_khang_nghi_payload.json
        try:
            export_khang_nghi.build_khang_nghi(kn_data, "_smoke_kn.docx", dry_run=True)
        except Exception as e:
            err(f"Smoke test Kháng nghị thất bại: {e}")
            local_errs += 1

        # 7.3 Tương thích chéo Schema 2 chiều (chống lỗi dai_dien dict/list & mo_ta/ho_ten)
        try:
            export_khang_nghi.build_khang_nghi(tt_data, "_smoke_cross_kn.docx", dry_run=True, allow_draft=True)
            export_to_trinh.build_to_trinh(kn_data, "_smoke_cross_tt.docx", dry_run=True, allow_draft=True)
        except Exception as e:
            err(f"Smoke test tương thích chéo Schema thất bại: {e}")
            local_errs += 1

        # 7.4 Thông báo giải quyết đơn + xác minh KHÔNG tự bịa bi_don[0] khi thiếu nguoi_de_nghi
        try:
            export_dossier.build_thong_bao(tt_data, "_smoke_tb.docx", dry_run=True)
        except Exception as e:
            err(f"Smoke test Thông báo thất bại: {e}")
            local_errs += 1

        no_sender = json.loads(json.dumps(tt_data))
        no_sender["duong_su"].pop("nguoi_de_nghi", None)
        try:
            export_dossier.build_thong_bao(no_sender, "_smoke_tb_fail.docx", dry_run=True)
            err("Lỗi Zero Hallucination: Thông báo không chặn khi thiếu nguoi_de_nghi!")
            local_errs += 1
        except ValueError:
            pass  # Kỳ vọng phải chặn đúng!

        # 7.5 Báo cáo 6 bước từ JSON chuẩn hóa (kèm 4 trụ cột toàn diện, số học & phản biện ngược)
        try:
            tt_ext = json.loads(json.dumps(tt_data))
            tt_ext["bao_cao_6_buoc"] = {
                "buoc_2_tinh_trang_tha": ["Chưa thi hành án."],
                "buoc_3_van_de_cot_loi": ["Hiệu lực HĐCN ngày 19/5/2020 và người thứ ba ngay tình."],
                "buoc_4_danh_gia_toan_dien": {
                    "to_tung": ["Đủ tư cách tố tụng."],
                    "chung_cu": ["Đã có kết luận giám định."],
                    "noi_dung": ["Đánh giá lại Điều 133 BLDS 2015."],
                    "quyet_dinh_tha": ["Án phí tính lại theo mức bồi thường."],
                },
                "buoc_4_kiem_tra_so_hoc": {
                    "headers": ["Hạng mục", "Bản án tuyên", "Kiểm tra"],
                    "rows": [["Án phí DSST", "Theo giá ngạch", "Cần đối chiếu NQ 326/2016"]],
                },
                "buoc_5_phan_bien_nguoc": ["Kiểm tra ngưỡng vi phạm nghiêm trọng theo Điều 326 BLTTDS."],
            }
            md6 = ps.payload_to_6step_markdown(tt_ext)
            for marker in ("[TRẠNG THÁI HỒ SƠ ĐÃ KHÓA]", "4A. Đánh giá toàn diện", "Bảng kiểm số học", "Phản biện ngược"):
                if marker not in md6:
                    err(f"Báo cáo 6 bước thiếu mục bắt buộc: '{marker}'!")
                    local_errs += 1
            export_6step_report.build_doc_from_markdown(md6, "_smoke_6step.docx", dry_run=True)
        except Exception as e:
            err(f"Smoke test Báo cáo 6 bước thất bại: {e}")
            local_errs += 1

        # 7.5b Văn bản trung gian / Early Exit (Option 7)
        try:
            vbtg_data = json.loads(json.dumps(tt_data))
            vbtg_data["van_ban_trung_gian"] = {
                "loai_van_ban": "phieu_chuyen_don",
                "kinh_gui": ["Tòa án nhân dân tỉnh Bến Tre"],
                "noi_dung": ["Chuyển đơn theo thẩm quyền quy định tại Điều 4 NQ 01/2025/NQ-HĐTP."],
                "nguoi_ky": "Lãnh đạo Vụ",
            }
            export_dossier.build_van_ban_trung_gian(vbtg_data, "_smoke_vbtg.docx", dry_run=True)
        except Exception as e:
            err(f"Smoke test Văn bản trung gian (Option 7) thất bại: {e}")
            local_errs += 1

    # 7.6 Kiểm thử thay tag bảo toàn định dạng nội dòng (kể cả khi tag bị Word tách nhiều run)
    try:
        test_doc = Document()
        p = test_doc.add_paragraph()
        r0 = p.add_run("Kính gửi: ")
        r0.bold = True
        r1 = p.add_run("{{NGUOI_")
        r1.bold = False
        r2 = p.add_run("GUI_DON}}")
        r2.bold = False
        du.replace_tags_preserve_format(p, {"{{NGUOI_GUI_DON}}": "Ông Nguyễn Văn A"})
        if p.text != "Kính gửi: Ông Nguyễn Văn A" or not p.runs[0].bold or p.runs[1].bold:
            err("Lỗi bảo toàn định dạng nội dòng khi thay thế tag trong docx_utils!")
            local_errs += 1
    except Exception as e:
        err(f"Smoke test replace_tags_preserve_format thất bại: {e}")
        local_errs += 1

    # 7.7 Kiểm thử phân loại vòng đời văn bản pháp lý trong sync_legal_knowledge.py
    s_nq = slk.determine_lifecycle_status("04_01/a.pdf", "Nghị quyết số 01-2003-NQ-HĐTP.pdf")
    s_ttds = slk.determine_lifecycle_status("04_01/b.pdf", "1. Bộ luật tố tụng dân sự 2014, sđbs 2025.pdf")
    s_ldd = slk.determine_lifecycle_status("04_01/c.docx", "12. Luật đất đai 2003.docx")
    s_al = slk.determine_lifecycle_status("04_04_An_le_TANDTC/_Ban_goc/An_le_01.pdf", "An_le_01.pdf")
    if "🔵" not in s_nq or "🟢" not in s_ttds or "🟡" not in s_ldd or "ÁN LỆ" not in s_al:
        err(f"Lỗi phân loại vòng đời văn bản pháp lý: NQ={s_nq} | TTDS={s_ttds} | LDD2003={s_ldd} | AL={s_al}")
        local_errs += 1

    if local_errs == 0:
        ok("Toàn bộ bài kiểm thử tự động (Dry-run, Tờ trình rút gọn, Option 7, Báo cáo 6 bước toàn diện, Zero Hallucination, bảo toàn run, vòng đời luật) PASS 100%")


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
    check_7_dry_run_smoke_tests()

    print("\n" + "=" * 68)
    if ERRORS:
        print(f"[THẤT BẠI] Phát hiện {len(ERRORS)} lỗi nghiêm trọng, {len(WARNINGS)} cảnh báo.")
        sys.exit(1)
    else:
        print(f"[ĐẠT CHUẨN 100%] 0 lỗi nghiêm trọng, {len(WARNINGS)} cảnh báo.")
        sys.exit(0)


if __name__ == "__main__":
    main()
