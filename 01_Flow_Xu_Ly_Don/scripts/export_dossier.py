# -*- coding: utf-8 -*-
"""
export_dossier.py - Động cơ xuất hồ sơ văn bản toàn diện (All-in-One Judicial Exporter)
Chức năng:
- Nhận payload JSON tổng hợp từ Bước 5.5.
- Thực thi xuất file tương ứng theo Option 1, 2, 3, 4, 5, 6 trong < 0.5 giây.
- Tự động gom toàn bộ văn bản của cùng 1 hồ sơ vào 03_Ket_qua_Xu_ly/<Số>_Vu_<case_slug>/ (01_ = mới nhất).
- Tích hợp đầy đủ --dry-run, kiểm tra file-lock, quét sót placeholder và tự động dọn file tạm.
"""
import argparse
from datetime import datetime
import json
import os
import re
import sys
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

if sys.stdout.encoding != "utf-8":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

# Setup paths via centralized paths_config (Single Source of Truth)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import paths_config as cfg

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
    - Tự động cập nhật case_registry.json và bảng mục lục trong 03_Ket_qua_Xu_ly/README.md.
    Trả về đường dẫn thư mục đích cho hồ sơ target_base_name.
    """
    reg = _load_registry()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    is_new_case = target_base_name not in reg

    # Thu thập danh sách thư mục vụ án hiện có (trừ thư mục mẫu 99_Mau_Template_Demo / _Mau_Template_Demo)
    existing = {}
    for entry in os.listdir(OUTPUT_ROOT_DIR):
        full = os.path.join(OUTPUT_ROOT_DIR, entry)
        if not os.path.isdir(full):
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

    # Sắp xếp toàn bộ vụ án theo ngày tạo giảm dần (mới nhất ở vị trí 01)
    all_bases = sorted(
        set(existing.keys()) | {target_base_name},
        key=lambda b: (reg.get(b, ""), b),
        reverse=True,
    )
    pad = 3 if len(all_bases) >= 99 else 2
    desired_names = {b: f"{idx:0{pad}d}_{b}" for idx, b in enumerate(all_bases, 1)}

    if dry_run:
        return os.path.join(OUTPUT_ROOT_DIR, desired_names[target_base_name])

    # Thực hiện đổi tên 2 bước (qua tên tạm) để tránh đụng tên giữa các số thứ tự
    renames = []
    for base, cur_name in existing.items():
        new_name = desired_names[base]
        if cur_name != new_name:
            renames.append((base, cur_name, new_name))

    temp_map = []
    for base, cur_name, new_name in renames:
        cur_path = os.path.join(OUTPUT_ROOT_DIR, cur_name)
        tmp_path = os.path.join(OUTPUT_ROOT_DIR, f".__reorder_{base}")
        try:
            os.rename(cur_path, tmp_path)
            temp_map.append((tmp_path, os.path.join(OUTPUT_ROOT_DIR, new_name), base))
        except OSError as e:
            print(f"[WARN] Không thể đổi tên thư mục đang mở ({cur_name}): {e}")

    for tmp_path, dst_path, base in temp_map:
        try:
            os.rename(tmp_path, dst_path)
        except OSError as e:
            print(f"[WARN] Không thể hoàn tất đổi tên {base}: {e}")

    target_dir = os.path.join(OUTPUT_ROOT_DIR, desired_names[target_base_name])
    os.makedirs(target_dir, exist_ok=True)
    _save_registry(reg)
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


def _run_has_media(run_el):
    return (
        run_el.find(qn("w:pict")) is not None
        or run_el.find(qn("w:drawing")) is not None
        or run_el.find(".//{urn:schemas-microsoft-com:vml}line") is not None
        or run_el.find(".//{urn:schemas-microsoft-com:vml}shape") is not None
    )


def _replace_tags_preserve_format(p, replacements, default_size_pt=None):
    """Thay thế các tag {{...}} trong đoạn văn nhưng bảo toàn 100% VML/pict/drawing và định dạng run."""
    full_text = p.text
    new_text = full_text
    changed = False
    for tag, val in replacements.items():
        if tag in new_text:
            new_text = new_text.replace(tag, str(val))
            changed = True
    if not changed:
        return

    text_runs = [r for r in p.runs if not _run_has_media(r._r)]
    if text_runs:
        text_runs[0].text = new_text
        if default_size_pt is not None and text_runs[0].font.size is None:
            text_runs[0].font.size = Pt(default_size_pt)
        if not text_runs[0].font.name:
            text_runs[0].font.name = "Times New Roman"
        for extra in text_runs[1:]:
            extra.text = ""
    else:
        r = p.add_run(new_text)
        export_to_trinh.set_run_font(r, size_pt=default_size_pt or 14)


def build_thong_bao(payload, output_path, dry_run=False):
    """Xuất Thông báo giải quyết đơn theo chuẩn Thẩm phán TANDTC duyệt ký (TB 383/679)."""
    template_path = os.path.join(ROOT_DIR, "01_Flow_Xu_Ly_Don", "templates", "base_thong_bao_tld_tandtc.docx")

    meta = payload.get("metadata", {})
    ds = payload.get("duong_su", {})

    ng_de_nghi = ds.get("nguoi_de_nghi") or ds.get("nguoi_gui_don") or ""
    dia_chi_ng = ds.get("dia_chi_nguoi_gui") or ""
    if not ng_de_nghi and ds.get("bi_don"):
        ng_de_nghi = ds["bi_don"][0].get("ho_ten", "")
        if not dia_chi_ng:
            dia_chi_ng = ds["bi_don"][0].get("dia_chi", "")
    elif not ng_de_nghi and ds.get("nguyen_don"):
        ng_de_nghi = ds["nguyen_don"][0].get("ho_ten", "")
        if not dia_chi_ng:
            dia_chi_ng = ds["nguyen_don"][0].get("dia_chi", "")

    so_ba = meta.get("so_ba_pt", "...")
    ngay_ba = meta.get("ngay_ba_pt", "...")
    toa_an = meta.get("toa_an_pt", "Tòa án nhân dân...")
    loai_ba = meta.get("loai_ban_an_pt", "dân sự phúc thẩm")
    thu_tuc = meta.get("thu_tuc", "GIÁM ĐỐC THẨM").upper()

    noi_dung_custom = payload.get("noi_dung_thong_bao", [])
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
        "{{NGUOI_GUI_DON}}": ng_de_nghi or "Đương sự gửi đơn",
        "{{DIA_CHI_NGUOI_GUI}}": dia_chi_ng or "Theo địa chỉ trong hồ sơ vụ án",
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

    if os.path.exists(template_path):
        doc = docx.Document(template_path)
        for tbl in doc.tables:
            for row in tbl.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        _replace_tags_preserve_format(p, replacements)
        for p in doc.paragraphs:
            _replace_tags_preserve_format(p, replacements, default_size_pt=14)
    else:
        doc = docx.Document()
        sec = doc.sections[0]
        sec.top_margin = Cm(2.0)
        sec.bottom_margin = Cm(2.0)
        sec.left_margin = Cm(3.0)
        sec.right_margin = Cm(2.0)

        tbl = doc.add_table(rows=1, cols=2)
        export_to_trinh.set_no_borders(tbl)
        export_to_trinh.set_table_widths_and_indent(tbl, 9060, [4530, 4530], 0)
        c0 = tbl.cell(0, 0)
        export_to_trinh.add_rich_para(c0, [("TÒA ÁN NHÂN DÂN TỐI CAO", 12, True, False), (f"\nSố: {replacements['{{SO_THONG_BAO}}']}/TB-TA", 14, False, False)], first_line_cm=0)
        c1 = tbl.cell(0, 1)
        export_to_trinh.add_rich_para(c1, [("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc", 12, True, False), (f"\nHà Nội, ngày ... tháng ... năm {replacements['{{NAM}}']}", 14, False, True)], first_line_cm=0)

        p_t = doc.add_paragraph()
        export_to_trinh.format_para(p_t, align=WD_ALIGN_PARAGRAPH.CENTER, first_line_cm=0, before_pt=18, after_pt=12)
        r_t = p_t.add_run(f"THÔNG BÁO\nGIẢI QUYẾT ĐƠN ĐỀ NGHỊ {thu_tuc}")
        export_to_trinh.set_run_font(r_t, size_pt=14, bold=True)

        for line in [d1, d2, d3, d4]:
            p = doc.add_paragraph()
            export_to_trinh.format_para(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_cm=1.06, after_pt=6, before_pt=6, line_spacing=1.15)
            r = p.add_run(line)
            export_to_trinh.set_run_font(r, size_pt=14)

    # Validation: quét placeholder còn sót
    full_text = export_to_trinh.collect_all_text(doc)
    leftover = sorted(set(re.findall(r"\{\{[^}]+\}\}", full_text)))
    if leftover:
        raise ValueError(f"Thông báo còn sót placeholder chưa điền: {leftover}")

    if dry_run:
        print(f"[DRY-RUN OK] Thông báo giải quyết đơn hợp lệ 100% (chưa ghi file): {output_path}")
        return

    ok, lock_err = export_to_trinh.check_file_lock(output_path)
    if not ok:
        raise PermissionError(lock_err)

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    doc.save(output_path)
    print(f"[SUCCESS] Đã xuất Thông báo giải quyết đơn tại: {output_path}")


def payload_to_6step_markdown(data):
    """Chuyển đổi payload JSON chuẩn sang Markdown sạch cho Báo cáo 6 bước."""
    meta = data.get("metadata", {})
    md_parts = []
    md_parts.append("# THÔNG TIN ĐƯƠNG SỰ & BẢN ÁN BỊ KHIẾU NẠI")
    md_parts.append(f"- **Vụ việc:** {meta.get('tieu_de_vu_an') or meta.get('quan_he_tranh_chap', '')}")
    md_parts.append(f"- **Bản án phúc thẩm:** {meta.get('so_ba_pt', '')} ngày {meta.get('ngay_ba_pt') or meta.get('ngay_tuyen_pt', '')} của {meta.get('toa_an_pt', '')}")
    if meta.get("so_ba_st") or meta.get("toa_an_st"):
        md_parts.append(f"- **Bản án sơ thẩm:** {meta.get('so_ba_st', '')} ngày {meta.get('ngay_ba_st') or meta.get('ngay_tuyen_st', '')} của {meta.get('toa_an_st', '')}")
    md_parts.append("")

    md_parts.append("# I. TÓM TẮT NỘI DUNG VỤ ÁN")
    for block in data.get("muc_1_tom_tat", []):
        if isinstance(block, str):
            md_parts.append(block)
        elif isinstance(block, dict):
            if block.get("tieu_de"):
                md_parts.append(f"## {block['tieu_de']}")
            for para in block.get("doan_van", []):
                md_parts.append(str(para))

    md_parts.append("\n# II. QUÁ TRÌNH GIẢI QUYẾT TẠI CÁC CẤP TÒA ÁN")
    m2 = data.get("muc_2_qua_trinh") or data.get("muc_2_noi_dung_an") or {}
    md_parts.append(f"## {m2.get('so_tham_tieu_de', 'Cấp sơ thẩm')}")
    for p in (m2.get("so_tham_quyet_dinh") or m2.get("so_tham_noi_dung") or []):
        md_parts.append(str(p))
    for note in m2.get("so_tham_ghi_chu", []):
        md_parts.append(str(note))
    md_parts.append(f"## {m2.get('phuc_tham_tieu_de', 'Cấp phúc thẩm')}")
    for p in (m2.get("phuc_tham_quyet_dinh") or m2.get("phuc_tham_noi_dung") or []):
        md_parts.append(str(p))
    for ts in m2.get("tien_su_sau_phuc_tham", []):
        md_parts.append(str(ts))

    md_parts.append("\n# III. NỘI DUNG ĐỀ NGHỊ & ĐÁNH GIÁ CĂN CỨ")
    m3 = data.get("muc_3_de_nghi") or data.get("muc_3_nhan_dinh") or {}
    if isinstance(m3, dict):
        if m3.get("mo_dau"):
            md_parts.append(str(m3["mo_dau"]))
        for ly_do in (m3.get("cac_ly_do") or m3.get("cac_noi_dung_khieu_nai") or []):
            md_parts.append(str(ly_do) if str(ly_do).startswith("-") else f"- {ly_do}")
        for key_list in ("vi_pham_to_tung", "vi_pham_noi_dung"):
            for item in m3.get(key_list, []):
                md_parts.append(f"- {item}")

    md_parts.append("\n# IV. ĐỀ XUẤT GIẢI QUYẾT")
    m4 = data.get("muc_4_de_xuat") or data.get("muc_4_y_kien_de_xuat") or {}
    if isinstance(m4, dict):
        for item in (m4.get("noi_dung_phan_tich") or m4.get("noi_dung_khang_nghi") or []):
            md_parts.append(item.get("text", "") if isinstance(item, dict) else str(item))
        de_xuat = m4.get("de_xuat_chot") or m4.get("de_xuat_tkv") or m4.get("ket_luan_de_xuat") or ""
        if de_xuat:
            md_parts.append(f"\n## Quan điểm đề xuất: {de_xuat}")
        for tp in m4.get("du_thao_y_kien_tham_phan", []):
            md_parts.append(str(tp))

    return "\n".join(md_parts)


def main():
    parser = argparse.ArgumentParser(description="Động cơ xuất hồ sơ văn bản toàn diện (All-in-One Judicial Exporter)")
    parser.add_argument("--payload", required=True, help="Đường dẫn file JSON chứa dữ liệu vụ án")
    parser.add_argument("--option", required=True, choices=["1", "2", "3", "4", "5", "6", "all"], help="Lựa chọn loại văn bản (1-6 hoặc all)")
    parser.add_argument("--case-slug", default="", help="Tên viết tắt vụ việc (dùng đặt tên thư mục con và file)")
    parser.add_argument("--dry-run", action="store_true", help="Chỉ kiểm tra cấu trúc dữ liệu, không ghi file")
    parser.add_argument("--keep-temp", action="store_true", help="Giữ lại file payload tạm sau khi chạy xong")
    args = parser.parse_args()

    if not os.path.exists(args.payload):
        print(f"[ERROR] Không tìm thấy file payload: {args.payload}")
        sys.exit(1)

    with open(args.payload, "r", encoding="utf-8") as f:
        data = json.load(f)

    case_slug = args.case_slug.strip()
    if not case_slug:
        ds = data.get("duong_su", {})
        ng_don = ds.get("nguyen_don", [{}])[0].get("ho_ten", "") if ds.get("nguyen_don") else ""
        bi_don = ds.get("bi_don", [{}])[0].get("ho_ten", "") if ds.get("bi_don") else ""
        if ng_don and bi_don:
            case_slug = f"{ng_don}_{bi_don}".replace(" ", "_").replace("Ông_", "").replace("Bà_", "")
        else:
            case_slug = "Vu_an"

    # Định danh thư mục con theo hồ sơ vụ việc (tự xếp số 01_ = ngày tạo mới nhất lên đầu)
    case_base_name = _strip_num_prefix(case_slug)
    if not case_base_name.startswith("Vu_"):
        case_base_name = f"Vu_{case_base_name}"
    case_dir = sync_case_folder_order(case_base_name, dry_run=args.dry_run)

    file_to_trinh = os.path.join(case_dir, f"To_trinh_{case_slug}.docx")
    file_khang_nghi = os.path.join(case_dir, f"Quyet_dinh_Khang_nghi_{case_slug}.docx")
    file_thong_bao = os.path.join(case_dir, f"Thong_bao_Giai_quyet_don_{case_slug}.docx")
    file_6step = os.path.join(case_dir, f"Bao_cao_Nghien_cuu_6_buoc_{case_slug}.docx")

    opt = args.option
    exported_files = []

    try:
        # Option 1: Tờ trình
        if opt in ("1", "3", "6", "all"):
            print("[1] Đang xuất Tờ trình Chánh án TANDTC...")
            export_to_trinh.build_to_trinh(data, file_to_trinh, dry_run=args.dry_run)
            exported_files.append(file_to_trinh)

        # Option 2: Quyết định Kháng nghị
        if opt in ("2", "3", "6", "all"):
            print("[2] Đang xuất Quyết định Kháng nghị...")
            kn_data = data.copy()
            m2 = data.get("muc_2_qua_trinh") or data.get("muc_2_noi_dung_an") or {}
            m3 = data.get("muc_3_nhan_dinh") or {}
            m4 = data.get("muc_4_de_xuat") or data.get("muc_4_y_kien_de_xuat") or {}
            if "nhan_thay" not in kn_data:
                kn_data["nhan_thay"] = {
                    "mo_dau_ban_an": m2.get("phuc_tham_tieu_de") or m2.get("so_tham_tieu_de", ""),
                    "quyet_dinh_ban_an": m2.get("phuc_tham_quyet_dinh") or m2.get("so_tham_quyet_dinh", []),
                    "dien_bien_sau_phuc_tham": m2.get("tien_su_sau_phuc_tham", []),
                }
            if "xet_thay" not in kn_data:
                xt = list(m3.get("vi_pham_to_tung", [])) + list(m3.get("vi_pham_noi_dung", []))
                if not xt:
                    xt = [x.get("text", "") if isinstance(x, dict) else str(x) for x in m4.get("noi_dung_phan_tich", [])]
                kn_data["xet_thay"] = xt
            if "quyet_dinh" not in kn_data:
                qd = m4.get("noi_dung_khang_nghi") or ([m4["de_xuat_chot"]] if m4.get("de_xuat_chot") else [])
                kn_data["quyet_dinh"] = qd
            export_khang_nghi.build_khang_nghi(kn_data, file_khang_nghi, dry_run=args.dry_run)
            exported_files.append(file_khang_nghi)

        # Option 4: Thông báo trả lời đơn
        if opt in ("4", "6", "all"):
            print("[3] Đang xuất Thông báo giải quyết đơn...")
            build_thong_bao(data, file_thong_bao, dry_run=args.dry_run)
            exported_files.append(file_thong_bao)

        # Option 5: Báo cáo 6 bước
        if opt in ("5", "6", "all"):
            print("[4] Đang xuất Báo cáo toàn diện 6 bước...")
            meta = data.get("metadata", {})
            full_md = payload_to_6step_markdown(data)
            export_6step_report.build_doc_from_markdown(
                full_md,
                file_6step,
                title_case=meta.get("tieu_de_vu_an") or meta.get("quan_he_tranh_chap", "BÁO CÁO TOÀN DIỆN 6 BƯỚC"),
                subtitle=f"Bản án phúc thẩm số: {meta.get('so_ba_pt', '')}",
                dry_run=args.dry_run,
            )
            exported_files.append(file_6step)

        if not args.dry_run:
            update_output_readme()

        mode_label = "DRY-RUN HỢP LỆ" if args.dry_run else "HOÀN TẤT THÀNH CÔNG"
        print("\n" + "=" * 50)
        print(f"[{mode_label}] {len(exported_files)} văn bản Word (.docx) tại {case_dir}:")
        for f in exported_files:
            print(f"  - {f}")
        print("=" * 50)

    finally:
        # Tự động dọn dẹp file tạm payload sau khi xuất thật (giữ lại khi --dry-run hoặc --keep-temp)
        if not args.keep_temp and not args.dry_run and os.path.exists(args.payload):
            try:
                os.remove(args.payload)
                print(f"[CLEANUP] Đã xóa an toàn file payload tạm: {args.payload}")
            except Exception:
                pass


if __name__ == "__main__":
    main()
