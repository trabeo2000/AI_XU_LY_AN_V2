# -*- coding: utf-8 -*-
"""
export_dossier.py - Động cơ xuất hồ sơ văn bản toàn diện (All-in-One Judicial Exporter)
Chức năng:
- Nhận payload JSON tổng hợp từ Bước 5.5, chuẩn hóa qua payload_schema (Single Contract).
- Thực thi xuất file tương ứng theo Option 1, 2, 3, 4, 5, 6 trong < 0.5 giây.
- Tự động gom toàn bộ văn bản của cùng 1 hồ sơ vào 03_Ket_qua_Xu_ly/<Số>_Vu_<case_slug>/ (01_ = mới nhất)
  với cơ chế đổi tên có Rollback an toàn khi file/thư mục đang mở.
- Tích hợp --dry-run, File-lock check, kiểm tra sót placeholder/[CẦN ĐIỀN], đối chiếu chéo đa văn bản
  và chỉ dọn file payload tạm sau khi xuất THÀNH CÔNG.
"""
import argparse
from datetime import datetime
import json
import os
import re
import sys
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm

# Setup paths via centralized paths_config (Single Source of Truth)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import paths_config as cfg
import docx_utils as du
import payload_schema as ps

du.ensure_utf8_stdout()

ROOT_DIR = cfg.ROOT_DIR
SCRIPTS_DIR = cfg.SCRIPTS_DIR
OUTPUT_ROOT_DIR = cfg.OUTPUT_ROOT_DIR
REGISTRY_PATH = cfg.REGISTRY_PATH
TEMP_DIR = cfg.TEMP_DIR

os.makedirs(OUTPUT_ROOT_DIR, exist_ok=True)
os.makedirs(TEMP_DIR, exist_ok=True)

import export_to_trinh
import export_khang_nghi
import export_6step_report

# Tái xuất hàm chuyển đổi 6 bước để tương thích ngược
payload_to_6step_markdown = ps.payload_to_6step_markdown


def _strip_num_prefix(folder_name):
    """Tách bỏ tiền tố số thứ tự '01_' để lấy tên chuẩn 'Vu_...'."""
    return re.sub(r"^\d+_", "", folder_name)


def _load_registry():
    if os.path.exists(REGISTRY_PATH):
        try:
            with open(REGISTRY_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def _save_registry(reg):
    with open(REGISTRY_PATH, "w", encoding="utf-8") as f:
        json.dump(reg, f, ensure_ascii=False, indent=2)


def _format_vn_date(iso_str):
    try:
        dt = datetime.strptime(iso_str[:10], "%Y-%m-%d")
        return dt.strftime("%d/%m/%Y")
    except Exception:
        return iso_str[:10] if iso_str else ""


def sync_case_folder_order(target_base_name, dry_run=False):
    """
    Quản lý thứ tự thư mục con trong 03_Ket_qua_Xu_ly/:
    - Định dạng: <Số>_Vu_<Ten_ho_so> (01_ luôn là hồ sơ có ngày tạo gần hiện tại nhất).
    - Cơ chế đổi tên 2 pha CÓ ROLLBACK: nếu bất kỳ thư mục nào bị khóa, khôi phục nguyên trạng 100%,
      tuyệt đối không để kẹt thư mục tên ẩn '.__reorder_*'.
    - Tự động dọn bản ghi mồ côi trong case_registry.json khi thư mục tương ứng không còn tồn tại.
    """
    reg = _load_registry()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    is_new_case = target_base_name not in reg

    existing = {}
    for entry in os.listdir(OUTPUT_ROOT_DIR):
        full = os.path.join(OUTPUT_ROOT_DIR, entry)
        if not os.path.isdir(full) or entry.startswith(".__"):
            continue
        if "Mau_Template_Demo" in entry:
            continue
        base = _strip_num_prefix(entry)
        existing[base] = entry
        if base not in reg:
            mtime = os.path.getmtime(full)
            reg[base] = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S")

    if is_new_case:
        reg[target_base_name] = now_str

    active_bases = set(existing.keys()) | {target_base_name}
    all_bases = sorted(
        active_bases,
        key=lambda b: (reg.get(b, ""), b),
        reverse=True,
    )
    pad = 3 if len(all_bases) >= 99 else 2
    desired_names = {b: f"{idx:0{pad}d}_{b}" for idx, b in enumerate(all_bases, 1)}

    if dry_run:
        return os.path.join(OUTPUT_ROOT_DIR, desired_names[target_base_name])

    # Dọn các key cũ không còn tồn tại trên đĩa
    cleaned_reg = {b: reg[b] for b in all_bases if b in reg}

    renames = [
        (base, cur_name, desired_names[base])
        for base, cur_name in existing.items()
        if cur_name != desired_names[base]
    ]

    # Pha 1: đổi sang tên tạm; nếu lỗi bất kỳ thư mục nào -> rollback ngay toàn bộ Pha 1
    moved_phase1 = []
    phase1_failed = False
    for base, cur_name, new_name in renames:
        cur_path = os.path.join(OUTPUT_ROOT_DIR, cur_name)
        tmp_path = os.path.join(OUTPUT_ROOT_DIR, f".__reorder_{base}")
        try:
            os.rename(cur_path, tmp_path)
            moved_phase1.append((cur_path, tmp_path, os.path.join(OUTPUT_ROOT_DIR, new_name), base))
        except OSError as e:
            print(f"[WARN] Thư mục đang bị khóa ({cur_name}), hủy sắp xếp lại thứ tự và khôi phục an toàn: {e}")
            phase1_failed = True
            break

    if phase1_failed:
        for orig_path, tmp_path, _, _ in reversed(moved_phase1):
            try:
                os.rename(tmp_path, orig_path)
            except OSError as rb_err:
                print(f"[ERROR] Không thể rollback thư mục {tmp_path}: {rb_err}")
        # Trả về thư mục hiện có của target (nếu đã có) hoặc tạo mới theo tên mong muốn
        if target_base_name in existing:
            target_dir = os.path.join(OUTPUT_ROOT_DIR, existing[target_base_name])
        else:
            target_dir = os.path.join(OUTPUT_ROOT_DIR, desired_names[target_base_name])
            os.makedirs(target_dir, exist_ok=True)
        _save_registry(cleaned_reg)
        return target_dir

    # Pha 2: đổi từ tên tạm sang tên đích
    for orig_path, tmp_path, dst_path, base in moved_phase1:
        try:
            os.rename(tmp_path, dst_path)
        except OSError as e:
            print(f"[WARN] Không thể hoàn tất đổi tên {base}, khôi phục tên gốc: {e}")
            try:
                os.rename(tmp_path, orig_path)
            except OSError:
                pass

    target_dir = os.path.join(OUTPUT_ROOT_DIR, desired_names[target_base_name])
    os.makedirs(target_dir, exist_ok=True)
    _save_registry(cleaned_reg)
    return target_dir


def update_output_readme():
    """Cập nhật tự động bảng Mục lục hồ sơ (mới nhất ở trên cùng) trong 03_Ket_qua_Xu_ly/README.md."""
    reg = _load_registry()
    rows = []
    for entry in sorted(os.listdir(OUTPUT_ROOT_DIR)):
        full = os.path.join(OUTPUT_ROOT_DIR, entry)
        if not os.path.isdir(full) or entry.startswith(".__"):
            continue
        cnt = sum(1 for f in os.listdir(full) if os.path.isfile(os.path.join(full, f)))
        m = re.match(r"^(\d+)_(.+)$", entry)
        stt = m.group(1) if m else "--"
        base = m.group(2) if m else entry
        ngay = "Mẫu" if "Mau_Template_Demo" in entry else _format_vn_date(reg.get(base, ""))
        rows.append(f"| {stt} | {ngay} | [`{entry}/`]({entry}/) | {cnt} |")

    readme_text = (
        "# 03_Ket_qua_Xu_ly — KẾT QUẢ XỬ LÝ HỒ SƠ (XẾP NGÀY TẠO MỚI NHẤT Ở TRÊN CÙNG)\n\n"
        "Thư mục tổng hợp toàn bộ sản phẩm văn bản đầu ra của hệ thống (Tờ trình, Quyết định Kháng nghị, Thông báo giải quyết đơn, Báo cáo nghiên cứu 6 bước).\n\n"
        "## Quy tắc tổ chức & Sắp xếp tự động\n"
        "- **Định dạng tên thư mục con:** `<Số>_Vu_<Ten_ho_so>/` (ví dụ: `01_Vu_Ho_Van_Phon_Pham_Thi_Vong/`).\n"
        "- **Thứ tự sắp xếp:** Số **`01_`** luôn là hồ sơ có **ngày tạo gần hiện tại nhất**, giảm dần theo thời gian tạo; thư mục mẫu demo luôn nằm cuối (`99_Mau_Template_Demo/`).\n"
        "- **Cơ chế tự động:** Mỗi khi [`export_dossier.py`](../01_Flow_Xu_Ly_Don/scripts/export_dossier.py) xuất hồ sơ mới ở Bước 6, hệ thống tự động ghi nhận ngày tạo vào [`01_Flow_Xu_Ly_Don/case_registry.json`](../01_Flow_Xu_Ly_Don/case_registry.json), đánh lại số thứ tự để vụ mới nhất nhảy lên vị trí `01_` trên cùng và cập nhật bảng mục lục bên dưới.\n\n"
        "## Danh mục hồ sơ hiện có (Mới nhất → Cũ nhất)\n\n"
        "| STT | Ngày tạo | Thư mục hồ sơ | Số văn bản |\n"
        "| :---: | :---: | :--- | :---: |\n"
        + "\n".join(rows)
        + "\n"
    )
    with open(os.path.join(OUTPUT_ROOT_DIR, "README.md"), "w", encoding="utf-8") as f:
        f.write(readme_text)


def build_thong_bao(payload, output_path, dry_run=False, allow_draft=False):
    """Xuất Thông báo giải quyết đơn theo chuẩn Thẩm phán TANDTC duyệt ký (TB 383/679)."""
    template_path = cfg.BASE_THONG_BAO_TEMPLATE

    norm = ps.normalize_payload(payload)
    errors, warnings = ps.validate_payload(norm, "thong_bao", allow_draft=allow_draft)

    meta = norm["metadata"]
    ds = norm["duong_su"]

    # KHÔNG tự suy diễn bi_don[0]; địa chỉ chỉ tra cứu khớp chính xác theo tên người gửi đơn
    ng_de_nghi = ds.get("nguoi_de_nghi") or "[CẦN ĐIỀN TÊN NGƯỜI GỬI ĐƠN]"
    dia_chi_ng = ds.get("dia_chi_nguoi_gui") or "Theo địa chỉ trong hồ sơ vụ án"

    so_ba = meta.get("so_ba_pt", "...")
    ngay_ba = meta.get("ngay_ba_pt", "...")
    toa_an = meta.get("toa_an_pt", "Tòa án nhân dân...")
    loai_ba = meta.get("loai_ban_an_pt", "dân sự phúc thẩm")
    thu_tuc = meta.get("thu_tuc", "GIÁM ĐỐC THẨM").upper()

    noi_dung_custom = norm.get("noi_dung_thong_bao", [])
    if noi_dung_custom and len(noi_dung_custom) >= 4:
        d1, d2, d3, d4 = noi_dung_custom[0], noi_dung_custom[1], noi_dung_custom[2], noi_dung_custom[3]
    else:
        d1 = f"Tòa án nhân dân tối cao nhận được đơn của ông/bà đề nghị xem xét theo thủ tục {thu_tuc.lower()} đối với Bản án {loai_ba} số {so_ba} ngày {ngay_ba} của {toa_an}."
        d2 = "Sau khi nghiên cứu đơn đề nghị và các tài liệu có trong hồ sơ vụ án, Tòa án nhân dân tối cao nhận thấy các tình tiết khách quan và tài liệu chứng cứ trong hồ sơ vụ án đã được Tòa án cấp sơ thẩm và phúc thẩm thu thập, đánh giá đầy đủ, khách quan theo đúng quy định của pháp luật."
        d3 = f"Các phán quyết của Bản án {loai_ba} số {so_ba} ngày {ngay_ba} của {toa_an} là có căn cứ pháp luật. Nội dung đơn khiếu nại của đương sự không có căn cứ để chấp nhận."
        d4 = f"Vì vậy, không có căn cứ kháng nghị theo thủ tục {thu_tuc.lower()} đối với Bản án {loai_ba} số {so_ba} nêu trên.\nTòa án nhân dân tối cao thông báo để ông/bà biết."

    replacements = {
        "{{SO_THONG_BAO}}": meta.get("so_thong_bao", "        "),
        "{{NGAY}}": meta.get("ngay") or meta.get("ngay_trinh") or "      ",
        "{{THANG}}": meta.get("thang") or meta.get("thang_trinh") or "     ",
        "{{NAM}}": meta.get("nam") or meta.get("nam_trinh") or "2026",
        "{{THU_TUC_VIET_HOA}}": thu_tuc,
        "{{NGUOI_GUI_DON}}": ng_de_nghi,
        "{{DIA_CHI_NGUOI_GUI}}": dia_chi_ng,
        "{{DOAN_1_TIEP_NHAN_DON}}": d1,
        "{{DOAN_2_TINH_TIET_VU_AN}}": d2,
        "{{DOAN_3_LAP_LUAN_BAC_DON}}": d3,
        "{{DOAN_4_THONG_BAO_KET_LUAN}}": d4,
        "{{VU_VKSNDTC}}": meta.get("vu_vksndtc", "Vụ 9 - VKSNDTC"),
        "{{TOA_AN_TINH}}": meta.get("toa_an_tinh", toa_an),
        "{{TOA_AN_HUYEN}}": meta.get("toa_an_huyen", meta.get("toa_an_st", "TAND khu vực có thẩm quyền")),
        "{{CO_QUAN_THADS}}": meta.get("co_quan_thads", "Cơ quan Thi hành án dân sự có thẩm quyền"),
        "{{TEN_THAM_PHAN}}": meta.get("ho_ten_tham_phan") or "[CẦN ĐIỀN TÊN THẨM PHÁN]",
        "{{VU_GDKT}}": meta.get("don_vi_trinh", "Vụ GĐKT II"),
    }

    # Hỗ trợ linh hoạt chức danh ký theo TB 383/679: Thẩm phán TANDTC hoặc Thẩm phán TAND bậc 3
    cd_tp = (meta.get("chuc_danh_tham_phan") or "Thẩm phán TANDTC").strip()
    if "bậc 3" in cd_tp.lower():
        replacements["THẨM PHÁN TÒA ÁN NHÂN DÂN TỐI CAO"] = "THẨM PHÁN TÒA ÁN NHÂN DÂN BẬC 3"
        replacements["Thẩm phán TANDTC "] = "Thẩm phán TAND bậc 3 "

    if os.path.exists(template_path):
        doc = docx.Document(template_path)
        for p in du.iter_document_paragraphs(doc):
            du.replace_tags_preserve_format(p, replacements)
    else:
        doc = docx.Document()
        sec = doc.sections[0]
        sec.top_margin = Cm(2.0)
        sec.bottom_margin = Cm(2.0)
        sec.left_margin = Cm(3.0)
        sec.right_margin = Cm(2.0)

        tbl = doc.add_table(rows=1, cols=2)
        du.set_no_borders(tbl)
        du.set_table_widths_and_indent(tbl, 9060, [4530, 4530], 0)
        c0 = tbl.cell(0, 0)
        export_to_trinh.add_rich_para(c0, [("TÒA ÁN NHÂN DÂN TỐI CAO", 12, True, False), (f"\nSố: {replacements['{{SO_THONG_BAO}}']}/TB-TA", 14, False, False)], first_line_cm=0)
        c1 = tbl.cell(0, 1)
        export_to_trinh.add_rich_para(c1, [("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc", 12, True, False), (f"\nHà Nội, ngày ... tháng ... năm {replacements['{{NAM}}']}", 14, False, True)], first_line_cm=0)

        p_t = doc.add_paragraph()
        du.format_para(p_t, align=WD_ALIGN_PARAGRAPH.CENTER, first_line_cm=0, before_pt=18, after_pt=12)
        r_t = p_t.add_run(f"THÔNG BÁO\nGIẢI QUYẾT ĐƠN ĐỀ NGHỊ {thu_tuc}")
        du.set_run_font(r_t, size_pt=14, bold=True)

        for line in [d1, d2, d3, d4]:
            p = doc.add_paragraph()
            du.format_para(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_cm=1.06, after_pt=6, before_pt=6, line_spacing=1.15)
            r = p.add_run(line)
            du.set_run_font(r, size_pt=14)

    full_text = du.collect_all_text(doc)
    placeholders, unfilled = du.scan_unfilled_tokens(full_text)
    if placeholders:
        errors.append(f"Thông báo còn sót placeholder chưa điền: {placeholders}")
    if unfilled:
        (warnings if allow_draft else errors).append(f"Thông báo còn thẻ thiếu dữ liệu: {unfilled}")

    for w in warnings:
        print(f"  [WARN]  {w}")
    if errors:
        for e in errors:
            print(f"  [ERROR] {e}")
        raise ValueError("Dừng xuất Thông báo: " + "; ".join(errors))

    if dry_run:
        print(f"[DRY-RUN OK] Thông báo giải quyết đơn hợp lệ 100% (chưa ghi file): {output_path}")
        return

    ok, lock_err = du.check_file_lock(output_path)
    if not ok:
        raise PermissionError(lock_err)

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    doc.save(output_path)
    print(f"[SUCCESS] Đã xuất Thông báo giải quyết đơn tại: {output_path}")


VBTG_TITLES = {
    "phieu_chuyen_don": ("PHIẾU CHUYỂN ĐƠN", "Về việc chuyển đơn đề nghị giám đốc thẩm, tái thẩm"),
    "de_xuat_rut_ho_so": ("PHIẾU ĐỀ XUẤT RÚT HỒ SƠ VỤ ÁN", "Về việc đề nghị rút hồ sơ vụ án để nghiên cứu giải quyết đơn"),
    "yeu_cau_bo_sung": ("THÔNG BÁO", "Về việc yêu cầu bổ sung tài liệu, chứng cứ kèm theo đơn"),
    "bao_cao_co_quan_chuyen_don": ("CÔNG VĂN BÁO CÁO", "Về kết quả giải quyết đơn do cơ quan có thẩm quyền chuyển đến"),
}


def build_van_ban_trung_gian(payload, output_path, dry_run=False, allow_draft=False):
    """Xuất Văn bản trung gian / Early Exit (Option 7) chuẩn thể thức hành chính tư pháp TANDTC."""
    norm = ps.normalize_payload(payload)
    errors, warnings = ps.validate_payload(norm, "van_ban_trung_gian", allow_draft=allow_draft)

    meta = norm["metadata"]
    vbtg = norm.get("van_ban_trung_gian") or {}
    loai = (vbtg.get("loai_van_ban") or "phieu_chuyen_don").strip().lower()
    default_title, default_sub = VBTG_TITLES.get(loai, ("VĂN BẢN XỬ LÝ ĐƠN", ""))

    tieu_de = vbtg.get("tieu_de") or default_title
    trich_yeu = vbtg.get("trich_yeu") or default_sub
    so_vb = vbtg.get("so_van_ban") or "        /TANDTC-Vụ II"
    ngay = meta.get("ngay_ky") or meta.get("ngay_trinh") or "    "
    thang = meta.get("thang_ky") or meta.get("thang_trinh") or "    "
    nam = meta.get("nam_ky") or meta.get("nam_trinh") or "2026"
    dia_danh = meta.get("dia_danh") or "Hà Nội"
    don_vi = meta.get("don_vi_trinh") or ps.DEFAULT_DON_VI

    doc = docx.Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2.0)
    sec.bottom_margin = Cm(2.0)
    sec.left_margin = Cm(3.0)
    sec.right_margin = Cm(2.0)

    # Bảng 1: Quốc hiệu - Tiêu ngữ - Số văn bản - Ngày tháng
    tbl = doc.add_table(rows=1, cols=2)
    du.set_no_borders(tbl)
    du.set_table_widths_and_indent(tbl, 9060, [3822, 5238], 0)
    c0, c1 = tbl.cell(0, 0), tbl.cell(0, 1)
    du.clear_cell_paragraphs(c0)
    du.clear_cell_paragraphs(c1)
    export_to_trinh.add_rich_para(
        c0,
        [
            (meta.get("co_quan_cap_tren", ps.DEFAULT_CO_QUAN), 12, False, False),
            (f"\n{don_vi}", 12, True, False),
            (f"\nSố: {so_vb}", 13, False, False),
        ],
        align=WD_ALIGN_PARAGRAPH.CENTER,
        first_line_cm=0,
        after_pt=4,
    )
    export_to_trinh.add_rich_para(
        c1,
        [
            ("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", 12, True, False),
            ("\nĐộc lập - Tự do - Hạnh phúc", 13, True, False),
            (f"\n{dia_danh}, ngày {ngay} tháng {thang} năm {nam}", 13, False, True),
        ],
        align=WD_ALIGN_PARAGRAPH.CENTER,
        first_line_cm=0,
        after_pt=4,
    )

    # Tiêu đề & Trích yếu
    p_t = doc.add_paragraph()
    du.format_para(p_t, align=WD_ALIGN_PARAGRAPH.CENTER, first_line_cm=0, before_pt=14, after_pt=4)
    du.set_run_font(p_t.add_run(tieu_de), size_pt=14, bold=True)
    if trich_yeu:
        p_sub = doc.add_paragraph()
        du.format_para(p_sub, align=WD_ALIGN_PARAGRAPH.CENTER, first_line_cm=0, before_pt=0, after_pt=10)
        du.set_run_font(p_sub.add_run(trich_yeu), size_pt=13, italic=True)

    # Kính gửi
    kinh_gui = vbtg.get("kinh_gui") or "[CẦN ĐIỀN NƠI NHẬN CHÍNH]"
    kg_list = kinh_gui if isinstance(kinh_gui, list) else [str(kinh_gui)]
    for idx, kg in enumerate(kg_list):
        prefix = "Kính gửi: " if idx == 0 else "          "
        export_to_trinh.add_rich_para(
            doc,
            [(prefix, 14, idx == 0, False), (str(kg), 14, False, False)],
            align=WD_ALIGN_PARAGRAPH.LEFT,
            first_line_cm=1.0,
            after_pt=6,
        )

    # Nội dung chính
    noi_dung = vbtg.get("noi_dung") or []
    if isinstance(noi_dung, str):
        noi_dung = [noi_dung]
    for para in noi_dung:
        p = doc.add_paragraph()
        du.format_para(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_cm=1.06, before_pt=4, after_pt=6, line_spacing=1.15)
        du.set_run_font(p.add_run(str(para)), size_pt=14)

    # Bảng 2: Nơi nhận & Chữ ký
    tbl_sig = doc.add_table(rows=1, cols=2)
    du.set_no_borders(tbl_sig)
    du.set_table_widths_and_indent(tbl_sig, 9060, [4800, 4260], 0)
    cl, cr = tbl_sig.cell(0, 0), tbl_sig.cell(0, 1)
    du.clear_cell_paragraphs(cl)
    du.clear_cell_paragraphs(cr)

    noi_nhan = vbtg.get("noi_nhan") or ["- Như trên;", f"- Lưu: VT, {don_vi}."]
    export_to_trinh.add_rich_para(cl, [("Nơi nhận:", 12, True, True)], align=WD_ALIGN_PARAGRAPH.LEFT, first_line_cm=0, after_pt=2)
    for nn in noi_nhan:
        export_to_trinh.add_rich_para(cl, [(str(nn), 11, False, False)], align=WD_ALIGN_PARAGRAPH.LEFT, first_line_cm=0, after_pt=1)

    chuc_danh_ky = vbtg.get("chuc_danh_ky") or (
        "TL. CHÁNH ÁN\nVỤ TRƯỞNG" if loai != "de_xuat_rut_ho_so" else "THƯ KÝ VIÊN"
    )
    nguoi_ky = vbtg.get("nguoi_ky") or meta.get("ho_ten_tkv") or "[CẦN ĐIỀN NGƯỜI KÝ]"
    export_to_trinh.add_rich_para(
        cr,
        [(chuc_danh_ky, 12, True, False)],
        align=WD_ALIGN_PARAGRAPH.CENTER,
        first_line_cm=0,
        after_pt=48,
    )
    export_to_trinh.add_rich_para(
        cr,
        [(nguoi_ky, 14, True, False)],
        align=WD_ALIGN_PARAGRAPH.CENTER,
        first_line_cm=0,
        after_pt=4,
    )

    full_text = du.collect_all_text(doc)
    placeholders, unfilled = du.scan_unfilled_tokens(full_text)
    if placeholders:
        errors.append(f"Văn bản trung gian còn sót placeholder: {placeholders}")
    if unfilled:
        (warnings if allow_draft else errors).append(f"Văn bản trung gian còn thẻ thiếu dữ liệu: {unfilled}")

    for w in warnings:
        print(f"  [WARN]  {w}")
    if errors:
        for e in errors:
            print(f"  [ERROR] {e}")
        raise ValueError("Dừng xuất Văn bản trung gian: " + "; ".join(errors))

    if dry_run:
        print(f"[DRY-RUN OK] Văn bản trung gian ({loai}) hợp lệ 100% (chưa ghi file): {output_path}")
        return

    ok, lock_err = du.check_file_lock(output_path)
    if not ok:
        raise PermissionError(lock_err)

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    doc.save(output_path)
    print(f"[SUCCESS] Đã xuất Văn bản trung gian ({loai}) tại: {output_path}")


def _verify_cross_document_consistency(exported_paths, norm_data):
    """Đối chiếu chéo nhất quán số/ngày bản án giữa nhiều văn bản cùng bộ hồ sơ."""
    if len(exported_paths) < 2:
        return
    meta = norm_data.get("metadata", {})
    key_vals = [
        ("Số bản án PT", meta.get("so_ba_pt", "")),
        ("Ngày bản án PT", meta.get("ngay_ba_pt", "")),
    ]
    for path in exported_paths:
        if not os.path.exists(path):
            continue
        txt = du.collect_all_text(docx.Document(path))
        for label, val in key_vals:
            val_s = str(val or "").strip()
            if val_s and val_s != "..." and val_s not in txt:
                print(f"[WARN] Đối chiếu chéo: {os.path.basename(path)} không thấy {label} ('{val_s}').")


def main():
    parser = argparse.ArgumentParser(description="Động cơ xuất hồ sơ văn bản toàn diện (All-in-One Judicial Exporter)")
    parser.add_argument("--payload", required=True, help="Đường dẫn file JSON chứa dữ liệu vụ án")
    parser.add_argument("--option", required=True, choices=["1", "2", "3", "4", "5", "6", "7", "all"], help="Lựa chọn loại văn bản (1-7 hoặc all)")
    parser.add_argument("--case-slug", default="", help="Tên viết tắt vụ việc (dùng đặt tên thư mục con và file)")
    parser.add_argument("--dry-run", action="store_true", help="Chỉ kiểm tra cấu trúc dữ liệu, không ghi file")
    parser.add_argument("--allow-draft", action="store_true", help="Cho phép xuất bản nháp khi còn thẻ [CẦN ĐIỀN...]")
    parser.add_argument("--keep-temp", action="store_true", help="Giữ lại file payload tạm sau khi chạy xong")
    args = parser.parse_args()

    if not os.path.exists(args.payload):
        print(f"[ERROR] Không tìm thấy file payload: {args.payload}")
        sys.exit(1)

    with open(args.payload, "r", encoding="utf-8") as f:
        raw_data = json.load(f)

    data = ps.normalize_payload(raw_data)
    meta = data["metadata"]
    ds = data["duong_su"]

    case_slug = du.slugify_vn(args.case_slug.strip())
    if not case_slug:
        ng_don = ds["nguyen_don"][0].get("ho_ten", "") if ds.get("nguyen_don") else ""
        bi_don = ds["bi_don"][0].get("ho_ten", "") if ds.get("bi_don") else ""
        if ng_don and bi_don:
            clean_nd = re.sub(r"^(Ông|Bà|Anh|Chị|Cụ)\s+", "", ng_don, flags=re.IGNORECASE)
            clean_bd = re.sub(r"^(Ông|Bà|Anh|Chị|Cụ)\s+", "", bi_don, flags=re.IGNORECASE)
            case_slug = du.slugify_vn(f"{clean_nd}_{clean_bd}") or "Vu_an"
        else:
            case_slug = "Vu_an"

    case_base_name = _strip_num_prefix(case_slug)
    if not case_base_name.startswith("Vu_"):
        case_base_name = f"Vu_{case_base_name}"
    case_dir = sync_case_folder_order(case_base_name, dry_run=args.dry_run)

    file_to_trinh = os.path.join(case_dir, f"To_trinh_{case_slug}.docx")
    file_khang_nghi = os.path.join(case_dir, f"Quyet_dinh_Khang_nghi_{case_slug}.docx")
    file_thong_bao = os.path.join(case_dir, f"Thong_bao_Giai_quyet_don_{case_slug}.docx")
    file_6step = os.path.join(case_dir, f"Bao_cao_Nghien_cuu_6_buoc_{case_slug}.docx")
    loai_vbtg = du.slugify_vn((data.get("van_ban_trung_gian") or {}).get("loai_van_ban") or "Van_ban_trung_gian")
    file_vbtg = os.path.join(case_dir, f"{loai_vbtg}_{case_slug}.docx")

    opt = args.option
    exported_files = []
    export_succeeded = False

    try:
        # Option 1: Tờ trình
        if opt in ("1", "3", "6", "all"):
            print("[1] Đang xuất Tờ trình Chánh án TANDTC...")
            export_to_trinh.build_to_trinh(data, file_to_trinh, dry_run=args.dry_run, allow_draft=args.allow_draft)
            exported_files.append(file_to_trinh)

        # Option 2: Quyết định Kháng nghị
        if opt in ("2", "3", "6", "all"):
            print("[2] Đang xuất Quyết định Kháng nghị...")
            kn_data = ps.build_khang_nghi_payload(data)
            export_khang_nghi.build_khang_nghi(kn_data, file_khang_nghi, dry_run=args.dry_run, allow_draft=args.allow_draft)
            exported_files.append(file_khang_nghi)

        # Option 4: Thông báo trả lời đơn
        if opt in ("4", "6", "all"):
            print("[3] Đang xuất Thông báo giải quyết đơn...")
            build_thong_bao(data, file_thong_bao, dry_run=args.dry_run, allow_draft=args.allow_draft)
            exported_files.append(file_thong_bao)

        # Option 5: Báo cáo 6 bước
        if opt in ("5", "6", "all"):
            print("[4] Đang xuất Báo cáo toàn diện 6 bước...")
            full_md = ps.payload_to_6step_markdown(data)
            export_6step_report.build_doc_from_markdown(
                full_md,
                file_6step,
                title_case=meta.get("tieu_de_vu_an") or meta.get("quan_he_tranh_chap", "BÁO CÁO TOÀN DIỆN 6 BƯỚC"),
                subtitle=f"Bản án phúc thẩm số: {meta.get('so_ba_pt', '')}",
                dry_run=args.dry_run,
                don_vi_trinh=meta.get("don_vi_trinh", ps.DEFAULT_DON_VI),
                thu_tuc=meta.get("thu_tuc", "GIÁM ĐỐC THẨM"),
            )
            exported_files.append(file_6step)

        # Option 7: Văn bản trung gian / Early Exit (Phiếu chuyển đơn / Rút hồ sơ / Yêu cầu bổ sung / Báo cáo CQ chuyển đơn)
        if opt == "7" or (opt in ("6", "all") and data.get("van_ban_trung_gian")):
            print("[5] Đang xuất Văn bản trung gian / Early Exit...")
            build_van_ban_trung_gian(data, file_vbtg, dry_run=args.dry_run, allow_draft=args.allow_draft)
            exported_files.append(file_vbtg)

        if not args.dry_run:
            _verify_cross_document_consistency(exported_files, data)
            update_output_readme()

        export_succeeded = True
        mode_label = "DRY-RUN HỢP LỆ" if args.dry_run else "HOÀN TẤT THÀNH CÔNG"
        print("\n" + "=" * 50)
        print(f"[{mode_label}] {len(exported_files)} văn bản Word (.docx) tại {case_dir}:")
        for f in exported_files:
            print(f"  - {f}")
        print("=" * 50)

    finally:
        # CHỈ xóa file payload tạm khi xuất THÀNH CÔNG (giữ nguyên để debug/chạy lại nếu gặp lỗi)
        if export_succeeded and not args.keep_temp and not args.dry_run and os.path.exists(args.payload):
            try:
                os.remove(args.payload)
                print(f"[CLEANUP] Đã xóa an toàn file payload tạm: {args.payload}")
            except Exception:
                pass


if __name__ == "__main__":
    main()
