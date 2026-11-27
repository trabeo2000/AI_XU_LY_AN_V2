# -*- coding: utf-8 -*-
"""
Script xuất Quyết định Kháng nghị GĐT / Tái thẩm của Chánh án TANDTC (.docx) chuẩn 100% Bản phát hành gốc.
Tuân thủ AGENTS.md (Bước 6):
- Mã hóa UTF-8 Unicode hoàn toàn bằng python-docx (không dùng PowerShell COM).
- Dựa trên phôi XML gốc base_khang_nghi_tandtc.docx (giữ 100% 2 đường kẻ VML, footer1/footer2, styles, margins).
- Tích hợp chuẩn hóa Schema thống nhất, kiểm tra DOM phôi, File-lock check, Dry-run mode,
  Placeholder/[CẦN ĐIỀN] scan và Cross-consistency check.
"""
import argparse
import json
import os
import sys
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import paths_config as cfg
import docx_utils as du
import payload_schema as ps

du.ensure_utf8_stdout()

BASE_TEMPLATE = cfg.BASE_KHANG_NGHI_TEMPLATE
FONT_NAME = du.FONT_NAME

# Tái xuất các hàm dùng chung để bảo đảm tương thích ngược
set_run_font = du.set_run_font
replace_para_runs_preserve_vml = du.replace_para_runs_preserve_vml
set_no_borders = du.set_no_borders
check_file_lock = du.check_file_lock
collect_all_text = du.collect_all_text


def format_para(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_cm=1.0, left_cm=0,
                before_pt=5, after_pt=0, line_spacing=1.15):
    return du.format_para(
        p, align=align, first_line_cm=first_line_cm, left_cm=left_cm,
        before_pt=before_pt, after_pt=after_pt, line_spacing=line_spacing,
    )


def set_table_cols_dxa(tbl, col_dxas):
    du.set_table_widths_and_indent(tbl, total_dxa=sum(col_dxas), col_dxas=col_dxas, indent_dxa=None)


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


def verify_template_dom(doc):
    """Kiểm định cấu trúc phôi base_khang_nghi_tandtc.docx trước khi ghi dữ liệu theo chỉ số."""
    if len(doc.tables) < 1:
        raise ValueError(f"Phôi Kháng nghị không hợp lệ: cần >= 1 bảng tiêu đề, hiện có {len(doc.tables)}.")
    if len(doc.paragraphs) < 8:
        raise ValueError(f"Phôi Kháng nghị không hợp lệ: cần >= 8 đoạn tiêu đề, hiện có {len(doc.paragraphs)}.")


def populate_header_table(doc, meta):
    """Bảng 1: TÒA ÁN NHÂN DÂN TỐI CAO / Số: ... | Quốc hiệu - Tiêu ngữ / Ngày tháng (giữ 100% VML lines)."""
    tbl0 = doc.tables[0]
    c_left = tbl0.cell(0, 0)
    c_right = tbl0.cell(0, 1)

    replace_para_runs_preserve_vml(
        c_left.paragraphs[0],
        [(meta.get("co_quan_ban_hanh", ps.DEFAULT_CO_QUAN), 12, True, False, False)]
    )
    so_kn = f"Số: {meta.get('so_quyet_dinh', '.../2026/KN-DS')}"
    replace_para_runs_preserve_vml(
        c_left.paragraphs[1],
        [(" ", 14, False, False, False), (so_kn, 14, False, False, True)]
    )

    replace_para_runs_preserve_vml(
        c_right.paragraphs[0],
        [
            ("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", 12, True, False, False),
            ("Độc lập - Tự do - Hạnh phúc", 14, True, False, True),
        ]
    )
    ngay_ky = meta.get("ngay_ky") or "..."
    thang_ky = meta.get("thang_ky") or "..."
    nam_ky = meta.get("nam_ky") or "2026"
    ngay_str = f"{meta.get('dia_danh', 'Hà Nội')}, ngày {ngay_ky} tháng {thang_ky} năm {nam_ky}"
    replace_para_runs_preserve_vml(
        c_right.paragraphs[1],
        [(ngay_str, 14, False, True, False)]
    )


def populate_title_block(doc, meta):
    """Khối tiêu đề QUYẾT ĐỊNH / KHÁNG NGHỊ... / Đối với... / CHÁNH ÁN / TANDTC."""
    paras = doc.paragraphs
    thu_tuc = meta.get("thu_tuc", "GIÁM ĐỐC THẨM").upper()
    default_td2 = "KHÁNG NGHỊ TÁI THẨM" if thu_tuc == "TÁI THẨM" else "KHÁNG NGHỊ GIÁM ĐỐC THẨM"

    replace_para_runs_preserve_vml(paras[2], [(meta.get("tieu_de_1", "QUYẾT ĐỊNH"), 14, True, False, False)])
    replace_para_runs_preserve_vml(paras[3], [(meta.get("tieu_de_2", default_td2), 14, True, False, False)])

    loai_ba = meta.get("loai_ban_an_pt", "dân sự phúc thẩm")
    so_ba = meta.get("so_ba_pt", "...")
    ngay_ba = meta.get("ngay_ba_pt", "...")
    toa_an = meta.get("toa_an_pt", "Tòa án nhân dân...")
    d1 = meta.get("doi_voi_dong_1") or f"Đối với Bản án {loai_ba} số {so_ba} ngày {ngay_ba}"
    d2 = meta.get("doi_voi_dong_2") or f"của {toa_an}"
    replace_para_runs_preserve_vml(paras[4], [(d1, 14, True, False, False)])
    replace_para_runs_preserve_vml(paras[5], [(d2, 14, True, False, False)])
    replace_para_runs_preserve_vml(paras[6], [(meta.get("tham_quyen_dong_1", "CHÁNH ÁN"), 14, True, False, False)])
    replace_para_runs_preserve_vml(paras[7], [(meta.get("tham_quyen_dong_2", ps.DEFAULT_CO_QUAN), 14, True, False, False)])


def add_parties_section(doc, meta, duong_su):
    """Khối Căn cứ pháp lý & Danh sách đương sự (nhãn tư cách IN NGHIÊNG, firstLine=1.0cm)."""
    can_cu = meta.get("can_cu_phap_ly") or ps.default_khang_nghi_can_cu(meta.get("thu_tuc", "GIÁM ĐỐC THẨM"))
    for cc in can_cu:
        add_kn_para(doc, cc, size_pt=14, bold=False, italic=False,
                    first_line_cm=1.0, before_pt=0, after_pt=0, line_spacing=1.15)

    loai_va = meta.get("loai_vu_an", "dân sự")
    qh_tc = meta.get("quan_he_tranh_chap", "Tranh chấp")
    add_rich_kn_para(
        doc,
        [
            (f"Sau khi nghiên cứu hồ sơ vụ án {loai_va} “", 14, False, False),
            (qh_tc, 14, False, True),
            ("” giữa các đương sự:", 14, False, False),
        ],
        first_line_cm=1.0, before_pt=5, after_pt=0, line_spacing=1.15
    )

    def _render_reps(rep_list):
        for dd in rep_list or []:
            dd_role = dd.get("vai_tro") or "Người đại diện"
            dd_desc = dd.get("mo_ta") or dd.get("thong_tin") or ""
            add_rich_kn_para(
                doc,
                [(dd_role, 14, False, True), (f": {dd_desc}", 14, False, False)],
                first_line_cm=1.0, before_pt=5, after_pt=0, line_spacing=1.15
            )

    def render_group(label, items, group_reps=None):
        if not items:
            return
        if len(items) == 1 and label in ("Nguyên đơn", "Bị đơn"):
            it = items[0]
            add_rich_kn_para(
                doc,
                [(f"- {label}:", 14, False, True), (f" {it.get('mo_ta', '')}", 14, False, False)],
                first_line_cm=1.0, before_pt=5, after_pt=0, line_spacing=1.15
            )
            _render_reps(it.get("dai_dien"))
        else:
            add_rich_kn_para(
                doc,
                [(f"- {label}:", 14, False, True)],
                first_line_cm=1.0, before_pt=5, after_pt=0, line_spacing=1.15
            )
            for idx, it in enumerate(items, 1):
                add_kn_para(doc, f"{idx}. {it.get('mo_ta', '')}", size_pt=14,
                            first_line_cm=1.0, before_pt=5, after_pt=0, line_spacing=1.15)
                _render_reps(it.get("dai_dien"))
        _render_reps(group_reps)

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

    p0 = c_left.paragraphs[0]
    format_para(p0, align=WD_ALIGN_PARAGRAPH.LEFT, first_line_cm=0, before_pt=0, after_pt=0, line_spacing=1.0)
    r0 = p0.add_run("Nơi nhận:")
    set_run_font(r0, size_pt=11, bold=True, italic=True)

    for line in noi_nhan:
        p = c_left.add_paragraph()
        format_para(p, align=WD_ALIGN_PARAGRAPH.LEFT, first_line_cm=0, before_pt=0, after_pt=0, line_spacing=1.0)
        r = p.add_run(str(line))
        set_run_font(r, size_pt=11, bold=False, italic=False)

    chuc_danh_lines = ky_duyet.get("chuc_danh_lines") or ["KT. CHÁNH ÁN", "PHÓ CHÁNH ÁN"]
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

    for _ in range(6):
        pe = c_right.add_paragraph()
        format_para(pe, align=WD_ALIGN_PARAGRAPH.CENTER, first_line_cm=0, left_cm=-0.18,
                    before_pt=0, after_pt=0, line_spacing=1.0)

    ho_ten_ky = ky_duyet.get("ho_ten_nguoi_ky") or "[CẦN ĐIỀN TÊN NGƯỜI KÝ]"
    p_name = c_right.add_paragraph()
    format_para(p_name, align=WD_ALIGN_PARAGRAPH.CENTER, first_line_cm=0, left_cm=-0.18,
                before_pt=5, after_pt=0, line_spacing=1.0)
    rn = p_name.add_run(ho_ten_ky)
    set_run_font(rn, size_pt=14, bold=True)


def validate_document(doc, kn_data, allow_draft=False):
    errors, warnings = ps.validate_payload(kn_data, "khang_nghi", allow_draft=allow_draft)
    full_text = collect_all_text(doc)

    placeholders, unfilled = du.scan_unfilled_tokens(full_text)
    if placeholders:
        errors.append(f"Còn sót placeholder chưa điền: {placeholders}")
    if unfilled:
        (warnings if allow_draft else errors).append(f"Còn thẻ thiếu dữ liệu bắt buộc: {unfilled}")

    meta = kn_data.get("metadata", {})
    nguoi_ky = (kn_data.get("ky_duyet") or {}).get("ho_ten_nguoi_ky", "")
    checks = [
        ("Số bản án phúc thẩm", meta.get("so_ba_pt", ""), 3),
        ("Ngày bản án phúc thẩm", meta.get("ngay_ba_pt", ""), 3),
        ("Tòa án phúc thẩm", meta.get("toa_an_pt", ""), 3),
    ]
    if nguoi_ky and "[CẦN ĐIỀN" not in nguoi_ky:
        checks.append(("Tên người ký", nguoi_ky, 1))
    for label, val, min_cnt in checks:
        val_s = str(val or "").strip()
        if not val_s or val_s in ("...", "Tòa án nhân dân..."):
            continue
        count = full_text.count(val_s)
        if count < min_cnt:
            warnings.append(f"{label} ('{val_s}') xuất hiện {count} lần (kỳ vọng >= {min_cnt}) — kiểm tra lại tính nhất quán.")

    return errors, warnings


def build_khang_nghi(data, out_path, dry_run=False, allow_draft=False):
    if not os.path.exists(BASE_TEMPLATE):
        raise FileNotFoundError(f"Không tìm thấy phôi chuẩn: {BASE_TEMPLATE}")

    kn_data = ps.build_khang_nghi_payload(ps.normalize_payload(data))
    meta = kn_data["metadata"]
    doc = Document(BASE_TEMPLATE)
    verify_template_dom(doc)

    # 1. Bảng 1: Cơ quan / Số quyết định | Quốc hiệu / Tiêu ngữ / Ngày tháng
    populate_header_table(doc, meta)

    # 2. Tiêu đề QUYẾT ĐỊNH KHÁNG NGHỊ... / CHÁNH ÁN TANDTC
    populate_title_block(doc, meta)

    # 3. Căn cứ pháp lý & Danh sách đương sự
    add_parties_section(doc, meta, kn_data["duong_su"])

    # 4. NHẬN THẤY:
    add_kn_para(doc, "NHẬN THẤY:", size_pt=14, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
                first_line_cm=0, before_pt=12, after_pt=0, line_spacing=1.15)
    nt = kn_data.get("nhan_thay", [])
    if isinstance(nt, dict):
        if nt.get("mo_dau_ban_an"):
            add_kn_para(doc, nt["mo_dau_ban_an"], size_pt=14, first_line_cm=1.0, before_pt=5, after_pt=0, line_spacing=1.15)
        for qd in nt.get("quyet_dinh_ban_an", []):
            add_kn_para(doc, str(qd), size_pt=14, italic=True, first_line_cm=1.0, before_pt=5, after_pt=0, line_spacing=1.15)
        for dt in nt.get("dien_bien_sau_phuc_tham", []):
            add_kn_para(doc, str(dt), size_pt=14, first_line_cm=1.0, before_pt=5, after_pt=0, line_spacing=1.15)
    elif isinstance(nt, list):
        for p in nt:
            add_kn_para(doc, str(p), size_pt=14, first_line_cm=1.0, before_pt=5, after_pt=0, line_spacing=1.15)

    # 5. XÉT THẤY:
    add_kn_para(doc, "XÉT THẤY:", size_pt=14, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
                first_line_cm=0, before_pt=12, after_pt=12, line_spacing=1.0)
    xt = kn_data.get("xet_thay", [])
    if isinstance(xt, list):
        for doan in xt:
            add_kn_para(doc, str(doan), size_pt=14, first_line_cm=1.0, before_pt=5, after_pt=0, line_spacing=1.15)
    elif isinstance(xt, dict):
        for v in xt.values():
            add_kn_para(doc, str(v), size_pt=14, first_line_cm=1.0, before_pt=5, after_pt=0, line_spacing=1.15)

    # Câu chuyển "Vì các lẽ trên;" (không chèn thêm \t để tránh thụt lề kép)
    add_kn_para(doc, "Vì các lẽ trên;", size_pt=14, first_line_cm=1.0, before_pt=5, after_pt=0, line_spacing=1.15)

    # 6. QUYẾT ĐỊNH:
    add_kn_para(doc, "QUYẾT ĐỊNH:", size_pt=14, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
                first_line_cm=0, before_pt=12, after_pt=12, line_spacing=1.0)
    qds = kn_data.get("quyet_dinh", [])
    if isinstance(qds, list):
        for idx, qd in enumerate(qds):
            is_last = (idx == len(qds) - 1)
            add_kn_para(doc, str(qd), size_pt=14, first_line_cm=1.0,
                        before_pt=0, after_pt=(12 if is_last else 5), line_spacing=1.15)
    elif isinstance(qds, dict):
        for qd in qds.values():
            add_kn_para(doc, str(qd), size_pt=14, first_line_cm=1.0, before_pt=5, after_pt=0, line_spacing=1.15)

    # 7. Bảng 2: Nơi nhận & Chữ ký
    noi_nhan = kn_data.get("noi_nhan") or [
        "- Như Điều 3;",
        "- VKSNDTC (để biết);",
        "- TAND cấp tỉnh;",
        "- Lưu hồ sơ vụ án.",
    ]
    raw_kd = kn_data.get("ky_duyet") or {}
    ky_duyet = {
        "chuc_danh_lines": raw_kd.get("chuc_danh_lines") or ["KT. CHÁNH ÁN", "PHÓ CHÁNH ÁN"],
        "ho_ten_nguoi_ky": raw_kd.get("ho_ten_nguoi_ky") or meta.get("nguoi_ky_khang_nghi") or "[CẦN ĐIỀN TÊN NGƯỜI KÝ]",
    }
    kn_data["ky_duyet"] = ky_duyet
    add_noi_nhan_and_signature_table(doc, noi_nhan, ky_duyet)

    # --- VALIDATION GATE ---
    errors, warnings = validate_document(doc, kn_data, allow_draft=allow_draft)
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
        raise ValueError("Dừng xuất Quyết định Kháng nghị do phát hiện lỗi kiểm tra dữ liệu/placeholder!")

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
    parser.add_argument("--allow-draft", action="store_true", help="Cho phép xuất bản nháp khi còn thẻ [CẦN ĐIỀN...]")
    args = parser.parse_args()

    with open(args.input, "r", encoding="utf-8") as f:
        data = json.load(f)

    build_khang_nghi(data, args.output, dry_run=args.dry_run, allow_draft=args.allow_draft)


if __name__ == "__main__":
    main()
