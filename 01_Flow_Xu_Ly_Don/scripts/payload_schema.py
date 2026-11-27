# -*- coding: utf-8 -*-
"""
payload_schema.py - Chuẩn hóa cấu trúc dữ liệu (Single Payload Contract), kiểm định
chặt chẽ trước khi xuất Word và chuyển đổi đồng nhất sang Kháng nghị / Báo cáo 6 bước.

Giải quyết triệt để:
- Xung đột cấu trúc `dai_dien` (dict vs list) và `mo_ta` vs `ho_ten` giữa Tờ trình và Kháng nghị.
- Loại bỏ hoàn toàn hành vi tự suy diễn sai `bi_don[0]` làm Người gửi đơn.
- Thống nhất các tên key alias lịch sử về 1 schema chuẩn duy nhất.
- Thống nhất hàm dựng Markdown Báo cáo 6 bước giữa export_dossier.py và export_6step_report.py.
"""
import copy
import re


DEFAULT_CO_QUAN = "TÒA ÁN NHÂN DÂN TỐI CAO"
DEFAULT_DON_VI = "VỤ GIÁM ĐỐC KIỂM TRA II"


def _normalize_official_spelling(text):
    """Chuẩn hóa dấu thanh Quốc hiệu / Tên cơ quan theo Nghị định 30/2020/NĐ-CP."""
    if not isinstance(text, str):
        return text
    return (
        text.replace("TOÀ ÁN", "TÒA ÁN")
        .replace("Toà án", "Tòa án")
        .replace("CỘNG HOÀ", "CỘNG HÒA")
    )


def normalize_dai_dien_list(raw_dd):
    """
    Chuẩn hóa `dai_dien` (hỗ trợ cả dict đơn lẻ, list dict hoặc str) về chuẩn list[dict]:
      [{"vai_tro": ..., "thong_tin": ..., "dia_chi": ..., "mo_ta": ...}, ...]
    """
    if not raw_dd:
        return []
    if isinstance(raw_dd, dict):
        items = [raw_dd]
    elif isinstance(raw_dd, list):
        items = raw_dd
    else:
        items = [{"mo_ta": str(raw_dd)}]

    result = []
    for item in items:
        if isinstance(item, str):
            val = item.strip()
            if val:
                result.append({
                    "vai_tro": "Người đại diện",
                    "thong_tin": val,
                    "dia_chi": "",
                    "mo_ta": val,
                })
            continue
        if not isinstance(item, dict):
            continue

        vai_tro = (item.get("vai_tro") or "Người đại diện").strip()
        ho_ten = (item.get("ho_ten") or "").strip()
        nam_sinh = str(item.get("nam_sinh") or "").strip()
        thong_tin = (item.get("thong_tin") or "").strip()
        dia_chi = (item.get("dia_chi") or "").strip()
        mo_ta = (item.get("mo_ta") or "").strip()

        if not mo_ta:
            if thong_tin:
                if dia_chi and dia_chi.lower() not in thong_tin.lower():
                    sep = " " if thong_tin.endswith((";", ".", ",")) else "; "
                    mo_ta = f"{thong_tin}{sep}địa chỉ: {dia_chi}"
                else:
                    mo_ta = thong_tin
            else:
                parts = []
                if ho_ten:
                    parts.append(ho_ten)
                if nam_sinh:
                    parts.append(f"sinh năm {nam_sinh}")
                if dia_chi:
                    parts.append(f"địa chỉ: {dia_chi}")
                mo_ta = ", ".join(parts)

        if not thong_tin:
            if ho_ten:
                thong_tin = f"{ho_ten}, sinh năm {nam_sinh};" if nam_sinh else f"{ho_ten};"
            else:
                thong_tin = mo_ta

        if mo_ta or thong_tin:
            result.append({
                "vai_tro": vai_tro,
                "thong_tin": thong_tin,
                "dia_chi": dia_chi,
                "mo_ta": mo_ta,
            })
    return result


def normalize_party_item(raw_item):
    """
    Chuẩn hóa 1 đương sự (hỗ trợ cả kiểu Tờ trình {ho_ten, nam_sinh, dia_chi}
    và kiểu Kháng nghị {mo_ta}) về cấu trúc thống nhất.
    """
    if isinstance(raw_item, str):
        txt = raw_item.strip()
        return {
            "ho_ten": "",
            "nam_sinh": "",
            "dia_chi": "",
            "mo_ta": txt,
            "dai_dien": [],
        }
    if not isinstance(raw_item, dict):
        txt = str(raw_item).strip()
        return {"ho_ten": "", "nam_sinh": "", "dia_chi": "", "mo_ta": txt, "dai_dien": []}

    ho_ten = (raw_item.get("ho_ten") or "").strip()
    nam_sinh = str(raw_item.get("nam_sinh") or "").strip()
    dia_chi = (raw_item.get("dia_chi") or "").strip()
    mo_ta = (raw_item.get("mo_ta") or "").strip()

    if not mo_ta:
        parts = []
        if ho_ten:
            parts.append(ho_ten)
        if nam_sinh:
            parts.append(f"sinh năm {nam_sinh}")
        if dia_chi:
            parts.append(f"địa chỉ: {dia_chi}")
        mo_ta = ", ".join(parts)

    dai_dien = normalize_dai_dien_list(raw_item.get("dai_dien"))
    return {
        "ho_ten": ho_ten,
        "nam_sinh": nam_sinh,
        "dia_chi": dia_chi,
        "mo_ta": mo_ta,
        "dai_dien": dai_dien,
    }


def _extract_address_from_mo_ta(mo_ta):
    if not mo_ta:
        return ""
    m = re.search(r"(?:địa chỉ|cư trú tại)\s*:\s*(.+)$", mo_ta, flags=re.IGNORECASE)
    return m.group(1).strip().rstrip(";") if m else ""


def lookup_party_address(duong_su, person_name):
    """
    Tra cứu chính xác địa chỉ của người gửi đơn từ danh sách đương sự đã chuẩn hóa
    (KHÔNG tự suy diễn từ bi_don[0]).
    """
    if not person_name or not isinstance(duong_su, dict):
        return ""
    needle = person_name.strip().lower()
    if not needle:
        return ""

    # Loại bỏ xưng hô phổ biến khi đối chiếu linh hoạt nhưng chính xác tên riêng
    clean_needle = re.sub(r"^(ông|bà|anh|chị|cụ)\s+", "", needle).strip()
    for group_key in ("nguyen_don", "bi_don", "nguoi_lien_quan"):
        for item in duong_su.get(group_key, []):
            if not isinstance(item, dict):
                continue
            ho_ten = (item.get("ho_ten") or "").lower()
            mo_ta = (item.get("mo_ta") or "").lower()
            if (ho_ten and (needle in ho_ten or (clean_needle and clean_needle in ho_ten))) or (
                mo_ta and (needle in mo_ta or (clean_needle and clean_needle in mo_ta))
            ):
                if item.get("dia_chi"):
                    return item["dia_chi"].rstrip(";")
                extracted = _extract_address_from_mo_ta(item.get("mo_ta", ""))
                if extracted:
                    return extracted
    return ""


def normalize_payload(raw_data):
    """
    Tạo bản sao sâu và chuẩn hóa toàn bộ payload JSON về 1 schema thống nhất.
    Tương thích ngược 100% với cả demo_payload.json (Tờ trình) và demo_khang_nghi_payload.json.
    """
    data = copy.deepcopy(raw_data) if isinstance(raw_data, dict) else {}

    # 1. Metadata
    meta = dict(data.get("metadata") or {})
    meta["co_quan_cap_tren"] = _normalize_official_spelling(
        meta.get("co_quan_cap_tren") or meta.get("co_quan_ban_hanh") or DEFAULT_CO_QUAN
    )
    meta["co_quan_ban_hanh"] = _normalize_official_spelling(
        meta.get("co_quan_ban_hanh") or meta["co_quan_cap_tren"]
    )
    meta["don_vi_trinh"] = (meta.get("don_vi_trinh") or DEFAULT_DON_VI).strip()

    # Ngày bản án ST / PT
    meta["ngay_ba_pt"] = meta.get("ngay_ba_pt") or meta.get("ngay_tuyen_pt") or ""
    meta["ngay_ba_st"] = meta.get("ngay_ba_st") or meta.get("ngay_tuyen_st") or ""

    # Ngày ký / ngày trình
    ngay_chung = meta.get("ngay_trinh") or meta.get("ngay_ky") or meta.get("ngay") or ""
    thang_chung = meta.get("thang_trinh") or meta.get("thang_ky") or meta.get("thang") or ""
    nam_chung = meta.get("nam_trinh") or meta.get("nam_ky") or meta.get("nam") or "2026"
    meta["ngay_trinh"] = meta.get("ngay_trinh") or ngay_chung
    meta["thang_trinh"] = meta.get("thang_trinh") or thang_chung
    meta["nam_trinh"] = meta.get("nam_trinh") or nam_chung
    meta["ngay_ky"] = meta.get("ngay_ky") or ngay_chung
    meta["thang_ky"] = meta.get("thang_ky") or thang_chung
    meta["nam_ky"] = meta.get("nam_ky") or nam_chung

    # Thủ tục GĐT / Tái thẩm & Chức danh Thẩm phán
    raw_tt = (meta.get("thu_tuc") or "").strip().upper()
    if not raw_tt:
        td2 = (meta.get("tieu_de_2") or "").upper()
        raw_tt = "TÁI THẨM" if "TÁI THẨM" in td2 else "GIÁM ĐỐC THẨM"
    meta["thu_tuc"] = raw_tt
    meta["chuc_danh_tham_phan"] = (meta.get("chuc_danh_tham_phan") or "Thẩm phán TANDTC").strip()
    meta["che_do_rut_gon"] = bool(meta.get("che_do_rut_gon", False))
    data["metadata"] = meta

    # 2. Đương sự
    raw_ds = data.get("duong_su") or {}
    ds = {
        "nguyen_don": [normalize_party_item(x) for x in (raw_ds.get("nguyen_don") or [])],
        "bi_don": [normalize_party_item(x) for x in (raw_ds.get("bi_don") or [])],
        "nguoi_lien_quan": [normalize_party_item(x) for x in (raw_ds.get("nguoi_lien_quan") or [])],
        "dai_dien_nguyen_don": normalize_dai_dien_list(raw_ds.get("dai_dien_nguyen_don")),
        "dai_dien_bi_don": normalize_dai_dien_list(raw_ds.get("dai_dien_bi_don")),
        "dai_dien_nguoi_lien_quan": normalize_dai_dien_list(raw_ds.get("dai_dien_nguoi_lien_quan")),
        "nhan_nguoi_de_nghi": raw_ds.get("nhan_nguoi_de_nghi")
        or ("Người đề nghị tái thẩm" if raw_tt == "TÁI THẨM" else "Người đề nghị giám đốc thẩm"),
        # KHÔNG tự suy diễn bi_don[0] — chỉ lấy giá trị khai báo tường minh
        "nguoi_de_nghi": (
            raw_ds.get("nguoi_de_nghi")
            or raw_ds.get("nguoi_gui_don")
            or meta.get("nguoi_de_nghi")
            or meta.get("nguoi_gui_don")
            or ""
        ).strip(),
        "dia_chi_nguoi_gui": (
            raw_ds.get("dia_chi_nguoi_gui")
            or meta.get("dia_chi_nguoi_gui")
            or ""
        ).strip(),
    }
    if ds["nguoi_de_nghi"] and not ds["dia_chi_nguoi_gui"]:
        ds["dia_chi_nguoi_gui"] = lookup_party_address(ds, ds["nguoi_de_nghi"])
    data["duong_su"] = ds

    # 3. Mục II (thống nhất các alias lịch sử)
    raw_m2 = data.get("muc_2_qua_trinh") or data.get("muc_2_noi_dung_an") or {}
    if isinstance(raw_m2, dict):
        data["muc_2_qua_trinh"] = {
            "so_tham_tieu_de": raw_m2.get("so_tham_tieu_de", "1. Tòa án cấp sơ thẩm tuyên xử:"),
            "so_tham_quyet_dinh": list(raw_m2.get("so_tham_quyet_dinh") or raw_m2.get("so_tham_noi_dung") or []),
            "so_tham_ghi_chu": list(raw_m2.get("so_tham_ghi_chu") or []),
            "phuc_tham_tieu_de": raw_m2.get("phuc_tham_tieu_de", "2. Tòa án cấp phúc thẩm tuyên xử:"),
            "phuc_tham_quyet_dinh": list(raw_m2.get("phuc_tham_quyet_dinh") or raw_m2.get("phuc_tham_noi_dung") or []),
            "tien_su_sau_phuc_tham": list(raw_m2.get("tien_su_sau_phuc_tham") or []),
        }

    # 4. Mục III (thống nhất các alias lịch sử)
    raw_m3 = data.get("muc_3_de_nghi") or data.get("muc_3_nhan_dinh") or {}
    if isinstance(raw_m3, dict):
        ly_do = (
            raw_m3.get("cac_ly_do")
            or raw_m3.get("cac_noi_dung_khieu_nai")
            or raw_m3.get("cac_noi_dung_khiếu_nai")
            or []
        )
        data["muc_3_de_nghi"] = {
            "mo_dau": raw_m3.get("mo_dau") or "",
            "cac_ly_do": list(ly_do),
            "vi_pham_to_tung": list(raw_m3.get("vi_pham_to_tung") or []),
            "vi_pham_noi_dung": list(raw_m3.get("vi_pham_noi_dung") or []),
        }

    # 5. Mục IV (thống nhất các alias lịch sử)
    raw_m4 = data.get("muc_4_de_xuat") or data.get("muc_4_y_kien_de_xuat") or {}
    if isinstance(raw_m4, dict):
        m4_norm = dict(raw_m4)
        m4_norm["noi_dung_phan_tich"] = list(
            raw_m4.get("noi_dung_phan_tich") or raw_m4.get("noi_dung_khang_nghi") or []
        )
        m4_norm["de_xuat_chot"] = (
            raw_m4.get("de_xuat_chot")
            or raw_m4.get("de_xuat_chinh")
            or raw_m4.get("de_xuat_tkv")
            or raw_m4.get("ket_luan_de_xuat")
            or ""
        )
        m4_norm["du_thao_y_kien_tham_phan"] = list(raw_m4.get("du_thao_y_kien_tham_phan") or [])
        m4_norm["cac_muc_phan_tich"] = list(raw_m4.get("cac_muc_phan_tich") or [])
        data["muc_4_de_xuat"] = m4_norm

    return data


def default_khang_nghi_can_cu(thu_tuc="GIÁM ĐỐC THẨM"):
    """Trả về căn cứ pháp lý chuẩn thẩm quyền Chánh án TANDTC theo BLTTDS 2015 (sđbs 2025)."""
    if str(thu_tuc).strip().upper() == "TÁI THẨM":
        return [
            "Căn cứ Điều 352, khoản 1 Điều 354 và Điều 355 Bộ luật Tố tụng dân sự năm 2015 (sửa đổi, bổ sung theo Luật số 85/2025/QH15);"
        ]
    return [
        "Căn cứ Điều 326, khoản 1 Điều 331 và khoản 2 Điều 332 Bộ luật Tố tụng dân sự năm 2015 (sửa đổi, bổ sung theo Luật số 85/2025/QH15);"
    ]


def build_khang_nghi_payload(norm_data):
    """Chuyển đổi sâu và an toàn từ payload tổng hợp sang payload chuẩn cho Quyết định Kháng nghị."""
    kn = copy.deepcopy(norm_data)
    meta = kn.setdefault("metadata", {})
    thu_tuc = meta.get("thu_tuc", "GIÁM ĐỐC THẨM").upper()

    if not meta.get("tieu_de_2"):
        meta["tieu_de_2"] = "KHÁNG NGHỊ TÁI THẨM" if thu_tuc == "TÁI THẨM" else "KHÁNG NGHỊ GIÁM ĐỐC THẨM"
    if not meta.get("can_cu_phap_ly"):
        meta["can_cu_phap_ly"] = default_khang_nghi_can_cu(thu_tuc)

    m2 = kn.get("muc_2_qua_trinh") or {}
    m3 = kn.get("muc_3_de_nghi") or {}
    m4 = kn.get("muc_4_de_xuat") or {}

    if "nhan_thay" not in kn or not kn["nhan_thay"]:
        kn["nhan_thay"] = {
            "mo_dau_ban_an": m2.get("phuc_tham_tieu_de") or m2.get("so_tham_tieu_de", ""),
            "quyet_dinh_ban_an": m2.get("phuc_tham_quyet_dinh") or m2.get("so_tham_quyet_dinh", []),
            "dien_bien_sau_phuc_tham": m2.get("tien_su_sau_phuc_tham", []),
        }

    if "xet_thay" not in kn or not kn["xet_thay"]:
        xt = list(m3.get("vi_pham_to_tung", [])) + list(m3.get("vi_pham_noi_dung", []))
        if not xt:
            xt = [
                x.get("text", "") if isinstance(x, dict) else str(x)
                for x in m4.get("noi_dung_phan_tich", [])
            ]
        kn["xet_thay"] = [x for x in xt if str(x).strip()]

    if "quyet_dinh" not in kn or not kn["quyet_dinh"]:
        qd_list = m4.get("noi_dung_khang_nghi") or []
        if not qd_list and m4.get("de_xuat_chot"):
            so_ba = meta.get("so_ba_pt", "...")
            ngay_ba = meta.get("ngay_ba_pt", "...")
            toa_an = meta.get("toa_an_pt", "Tòa án cấp phúc thẩm")
            loai_ba = meta.get("loai_ban_an_pt", "dân sự phúc thẩm")
            qd_list = [
                f"1. Kháng nghị đối với Bản án {loai_ba} số {so_ba} ngày {ngay_ba} của {toa_an}.",
                f"2. {m4['de_xuat_chot']}",
            ]
        kn["quyet_dinh"] = qd_list

    return kn


def _render_md_table(headers, rows):
    if not headers or not rows:
        return []
    lines = [
        "| " + " | ".join(str(h).replace("|", "/") for h in headers) + " |",
        "| " + " | ".join(["---"] * len(headers)) + " |",
    ]
    for r in rows:
        cells = [str(c).replace("\n", " ").replace("|", "/") for c in r]
        while len(cells) < len(headers):
            cells.append("")
        lines.append("| " + " | ".join(cells[: len(headers)]) + " |")
    return lines


def payload_to_6step_markdown(raw_data):
    """
    Hàm duy nhất (Single Source of Truth) chuyển payload JSON sang Markdown Báo cáo 6 bước.
    Được dùng chung bởi cả export_dossier.py và export_6step_report.py.
    """
    data = normalize_payload(raw_data)
    meta = data.get("metadata", {})
    ds = data.get("duong_su", {})
    bc6 = data.get("bao_cao_6_buoc") or {}

    so_ba_pt = meta.get("so_ba_pt", "")
    ngay_ba_pt = meta.get("ngay_ba_pt", "")
    toa_an_pt = meta.get("toa_an_pt", "")
    so_ba_st = meta.get("so_ba_st", "")
    ngay_ba_st = meta.get("ngay_ba_st", "")
    toa_an_st = meta.get("toa_an_st", "")
    thu_tuc = meta.get("thu_tuc", "GIÁM ĐỐC THẨM")
    ng_gui = ds.get("nguoi_de_nghi") or "[Chưa xác định]"
    vu_viec = meta.get("tieu_de_vu_an") or meta.get("quan_he_tranh_chap", "")

    md = []
    # Khối Callout Trạng thái hồ sơ đã khóa
    trang_thai = bc6.get("trang_thai_khoa")
    if isinstance(trang_thai, list) and trang_thai:
        md.append("> **[TRẠNG THÁI HỒ SƠ ĐÃ KHÓA]**")
        for line in trang_thai:
            md.append(f"> {line}")
    else:
        ba_str = f"Phúc thẩm số {so_ba_pt} ngày {ngay_ba_pt} ({toa_an_pt})"
        if so_ba_st or toa_an_st:
            ba_str = f"Sơ thẩm số {so_ba_st} ngày {ngay_ba_st} ({toa_an_st}); " + ba_str
        md.extend([
            "> **[TRẠNG THÁI HỒ SƠ ĐÃ KHÓA]**",
            f"> - **Vụ việc / Đương sự gửi đơn:** {vu_viec} | **Người gửi đơn:** {ng_gui}",
            f"> - **Bản án / QĐ bị khiếu nại:** {ba_str}",
            f"> - **Thủ tục xem xét:** {thu_tuc}",
            f"> - **Cán bộ nghiên cứu / Thẩm phán:** {meta.get('ho_ten_tkv', '')} / {meta.get('ho_ten_tham_phan', '')}",
        ])
    md.append("")

    # Bước 1: Định danh, đương sự & luật áp dụng
    md.append("# BƯỚC 1: ĐỊNH DANH HỒ SƠ, ĐƯƠNG SỰ & KHÓA HIỆU LỰC PHÁP LUẬT")
    md.append(f"- **Quan hệ tranh chấp:** {vu_viec}")
    if ds.get("nguyen_don"):
        md.append("- **Nguyên đơn:** " + "; ".join(p.get("mo_ta") or p.get("ho_ten", "") for p in ds["nguyen_don"]))
    if ds.get("bi_don"):
        md.append("- **Bị đơn:** " + "; ".join(p.get("mo_ta") or p.get("ho_ten", "") for p in ds["bi_don"]))
    if ds.get("nguoi_lien_quan"):
        md.append("- **Người có quyền lợi, nghĩa vụ liên quan:** " + "; ".join(p.get("mo_ta") or p.get("ho_ten", "") for p in ds["nguoi_lien_quan"]))
    for line in bc6.get("buoc_1_phap_ly", []):
        md.append(str(line))

    # Tóm tắt nội dung vụ án (Mục I)
    md.append("\n# I. TÓM TẮT NỘI DUNG VỤ ÁN & QUÁ TRÌNH GIẢI QUYẾT")
    for block in data.get("muc_1_tom_tat", []):
        if isinstance(block, str):
            md.append(block)
        elif isinstance(block, dict):
            if block.get("tieu_de"):
                md.append(f"## {block['tieu_de']}")
            for para in block.get("doan_van", []):
                md.append(str(para))

    m2 = data.get("muc_2_qua_trinh") or {}
    if m2:
        md.append(f"## {m2.get('so_tham_tieu_de', '1. Tòa án cấp sơ thẩm')}")
        for p in m2.get("so_tham_quyet_dinh", []):
            md.append(str(p))
        for note in m2.get("so_tham_ghi_chu", []):
            md.append(str(note))
        md.append(f"## {m2.get('phuc_tham_tieu_de', '2. Tòa án cấp phúc thẩm')}")
        for p in m2.get("phuc_tham_quyet_dinh", []):
            md.append(str(p))
        for ts in m2.get("tien_su_sau_phuc_tham", []):
            md.append(str(ts))

    # Bước 2 & 3: Điều kiện thụ lý, Tình trạng THA, Timeline, Vấn đề cốt lõi & Yêu cầu khiếu nại
    md.append("\n# BƯỚC 2 & 3: ĐIỀU KIỆN THỤ LÝ, TÌNH TRẠNG THA, DÒNG THỜI GIAN & VẤN ĐỀ CỐT LÕI")
    for line in bc6.get("buoc_2_thu_ly", []):
        md.append(str(line))

    tha_info = bc6.get("buoc_2_tinh_trang_tha")
    if isinstance(tha_info, str) and tha_info.strip():
        md.append("## Tình trạng Thi hành án & Người thứ ba ngay tình")
        md.append(tha_info.strip())
    elif isinstance(tha_info, list) and tha_info:
        md.append("## Tình trạng Thi hành án & Người thứ ba ngay tình")
        for item in tha_info:
            s = str(item)
            md.append(s if s.startswith("-") else f"- {s}")

    timeline = bc6.get("buoc_3_timeline")
    if isinstance(timeline, dict) and timeline.get("headers") and timeline.get("rows"):
        md.append("## Dòng thời gian tố tụng (Timeline)")
        md.extend(_render_md_table(timeline["headers"], timeline["rows"]))
    elif isinstance(timeline, list) and timeline:
        md.append("## Dòng thời gian tố tụng (Timeline)")
        for item in timeline:
            md.append(str(item) if str(item).startswith("-") else f"- {item}")

    vd_cot_loi = bc6.get("buoc_3_van_de_cot_loi")
    if isinstance(vd_cot_loi, dict) and vd_cot_loi.get("headers") and vd_cot_loi.get("rows"):
        md.append("## Cây vấn đề pháp lý cốt lõi của vụ án (Độc lập với đơn)")
        md.extend(_render_md_table(vd_cot_loi["headers"], vd_cot_loi["rows"]))
    elif isinstance(vd_cot_loi, list) and vd_cot_loi:
        md.append("## Cây vấn đề pháp lý cốt lõi của vụ án (Độc lập với đơn)")
        for item in vd_cot_loi:
            s = str(item)
            md.append(s if s.startswith("-") else f"- {s}")

    m3 = data.get("muc_3_de_nghi") or {}
    if m3.get("mo_dau"):
        md.append(str(m3["mo_dau"]))
    for ly_do in m3.get("cac_ly_do", []):
        s = str(ly_do)
        md.append(s if s.startswith("-") else f"- {s}")
    for vp in m3.get("vi_pham_to_tung", []):
        md.append(f"- **Vi phạm tố tụng:** {vp}")
    for vp in m3.get("vi_pham_noi_dung", []):
        md.append(f"- **Vi phạm nội dung:** {vp}")

    # Bước 4: Đánh giá toàn diện (Độc lập với đơn) & Ma trận đối chiếu yêu cầu trong đơn
    md.append("\n# BƯỚC 4: ĐÁNH GIÁ TOÀN DIỆN BẢN ÁN (ĐỘC LẬP VỚI ĐƠN) & ĐỐI CHIẾU YÊU CẦU TRONG ĐƠN")
    dg_toan_dien = bc6.get("buoc_4_danh_gia_toan_dien")
    if isinstance(dg_toan_dien, dict) and dg_toan_dien.get("headers") and dg_toan_dien.get("rows"):
        md.append("## 4A. Đánh giá toàn diện việc giải quyết vụ án của Tòa án hai cấp (4 trụ cột)")
        md.extend(_render_md_table(dg_toan_dien["headers"], dg_toan_dien["rows"]))
    elif isinstance(dg_toan_dien, dict) and dg_toan_dien:
        md.append("## 4A. Đánh giá toàn diện việc giải quyết vụ án của Tòa án hai cấp (4 trụ cột)")
        pillar_labels = (
            ("to_tung", "Trụ cột 1 — Tố tụng & Thẩm quyền"),
            ("chung_cu", "Trụ cột 2 — Thu thập & Đánh giá Chứng cứ"),
            ("noi_dung", "Trụ cột 3 — Áp dụng Pháp luật Nội dung"),
            ("quyet_dinh_tha", "Trụ cột 4 — Phần Quyết định & Khả năng Thi hành án"),
        )
        for key, title in pillar_labels:
            val = dg_toan_dien.get(key)
            if isinstance(val, str) and val.strip():
                md.append(f"### {title}")
                md.append(val.strip())
            elif isinstance(val, list) and val:
                md.append(f"### {title}")
                for it in val:
                    s = str(it)
                    md.append(s if s.startswith("-") else f"- {s}")
    elif isinstance(dg_toan_dien, list) and dg_toan_dien:
        md.append("## 4A. Đánh giá toàn diện việc giải quyết vụ án của Tòa án hai cấp (4 trụ cột)")
        for it in dg_toan_dien:
            s = str(it)
            md.append(s if s.startswith("-") else f"- {s}")

    so_hoc = bc6.get("buoc_4_kiem_tra_so_hoc")
    if isinstance(so_hoc, dict) and so_hoc.get("headers") and so_hoc.get("rows"):
        md.append("## Bảng kiểm số học (Án phí – Lãi suất – Diện tích – Giá trị)")
        md.extend(_render_md_table(so_hoc["headers"], so_hoc["rows"]))
    elif isinstance(so_hoc, list) and so_hoc:
        md.append("## Bảng kiểm số học (Án phí – Lãi suất – Diện tích – Giá trị)")
        for it in so_hoc:
            s = str(it)
            md.append(s if s.startswith("-") else f"- {s}")

    for vp_ngoai in bc6.get("buoc_4_vi_pham_ngoai_don", []):
        s = str(vp_ngoai)
        md.append(s if s.startswith("-") else f"- {s}")

    ma_tran = bc6.get("buoc_4_ma_tran")
    if isinstance(ma_tran, dict) and ma_tran.get("headers") and ma_tran.get("rows"):
        md.append("## 4B. Ma trận đối chiếu yêu cầu trong đơn của đương sự")
        md.extend(_render_md_table(ma_tran["headers"], ma_tran["rows"]))

    m4 = data.get("muc_4_de_xuat") or {}
    if m4.get("nhan_dinh_chung"):
        md.append(str(m4["nhan_dinh_chung"]))
    for item in m4.get("noi_dung_phan_tich", []):
        md.append(item.get("text", "") if isinstance(item, dict) else str(item))
    for sub in m4.get("cac_muc_phan_tich", []):
        if isinstance(sub, dict):
            if sub.get("tieu_de"):
                md.append(f"### {sub['tieu_de']}")
            for p in sub.get("doan_van", []):
                md.append(str(p))

    # Bước 5: Đánh giá căn cứ, Phản biện ngược & Đề xuất giải quyết
    md.append("\n# BƯỚC 5: ĐÁNH GIÁ CĂN CỨ, PHẢN BIỆN NGƯỢC & ĐỀ XUẤT PHƯƠNG ÁN XỬ LÝ")
    danh_gia = bc6.get("buoc_5_danh_gia")
    if isinstance(danh_gia, dict) and danh_gia.get("headers") and danh_gia.get("rows"):
        md.extend(_render_md_table(danh_gia["headers"], danh_gia["rows"]))
    elif isinstance(danh_gia, list):
        for line in danh_gia:
            md.append(str(line))

    phan_bien = bc6.get("buoc_5_phan_bien_nguoc")
    if isinstance(phan_bien, str) and phan_bien.strip():
        md.append("## Phản biện ngược (Kiểm tra độ bền lập luận)")
        md.append(phan_bien.strip())
    elif isinstance(phan_bien, list) and phan_bien:
        md.append("## Phản biện ngược (Kiểm tra độ bền lập luận)")
        for item in phan_bien:
            s = str(item)
            md.append(s if s.startswith("-") else f"- {s}")

    de_xuat = m4.get("de_xuat_chot", "")
    if de_xuat:
        md.append(f"## Quan điểm đề xuất của Thư ký viên\n{de_xuat}")
    if m4.get("du_thao_y_kien_tham_phan"):
        md.append("## Dự thảo ý kiến Thẩm phán")
        for tp in m4["du_thao_y_kien_tham_phan"]:
            md.append(str(tp))

    return "\n".join(md)


def validate_payload(norm_data, doc_type, allow_draft=False):
    """
    Kiểm định tính chặt chẽ của payload trước khi xuất văn bản.
    Trả về (errors, warnings).
    """
    errors = []
    warnings = []
    meta = norm_data.get("metadata", {})
    ds = norm_data.get("duong_su", {})

    def _is_missing(val):
        s = str(val or "").strip()
        return (not s) or s in ("...", "TAND...", "Tòa án nhân dân...") or ("[CẦN ĐIỀN" in s)

    if _is_missing(meta.get("so_ba_pt")):
        errors.append("Thiếu số bản án/quyết định phúc thẩm (`metadata.so_ba_pt`).")
    if _is_missing(meta.get("ngay_ba_pt")):
        errors.append("Thiếu ngày tuyên bản án/quyết định phúc thẩm (`metadata.ngay_ba_pt`).")
    if _is_missing(meta.get("toa_an_pt")):
        errors.append("Thiếu tên Tòa án phúc thẩm (`metadata.toa_an_pt`).")

    total_parties = len(ds.get("nguyen_don", [])) + len(ds.get("bi_don", [])) + len(ds.get("nguoi_lien_quan", []))
    if total_parties == 0:
        warnings.append("Danh sách đương sự (`duong_su`) đang trống.")

    if doc_type == "to_trinh":
        if _is_missing(meta.get("ho_ten_tkv")):
            (warnings if allow_draft else errors).append("Thiếu họ tên Thư ký viên (`metadata.ho_ten_tkv`).")
        if _is_missing(meta.get("ho_ten_tham_phan")):
            (warnings if allow_draft else errors).append("Thiếu họ tên Thẩm phán (`metadata.ho_ten_tham_phan`).")
        if not ds.get("nguoi_de_nghi"):
            warnings.append("Chưa khai báo tường minh người đề nghị GĐT/TT (`duong_su.nguoi_de_nghi`).")

    elif doc_type == "khang_nghi":
        ky_duyet = norm_data.get("ky_duyet") or {}
        nguoi_ky = ky_duyet.get("ho_ten_nguoi_ky") or meta.get("nguoi_ky_khang_nghi") or ""
        if _is_missing(nguoi_ky):
            (warnings if allow_draft else errors).append(
                "Thiếu họ tên người ký Quyết định Kháng nghị (`ky_duyet.ho_ten_nguoi_ky`)."
            )
        if not norm_data.get("xet_thay"):
            (warnings if allow_draft else errors).append("Thiếu nội dung phần XÉT THẤY (`xet_thay`).")
        if not norm_data.get("quyet_dinh"):
            (warnings if allow_draft else errors).append("Thiếu nội dung phần QUYẾT ĐỊNH (`quyet_dinh`).")

    elif doc_type == "thong_bao":
        if _is_missing(ds.get("nguoi_de_nghi")):
            (warnings if allow_draft else errors).append(
                "Thiếu tên người gửi đơn (`duong_su.nguoi_de_nghi`) — hệ thống không tự suy diễn từ Bị đơn."
            )
        if _is_missing(meta.get("ho_ten_tham_phan")):
            (warnings if allow_draft else errors).append("Thiếu tên Thẩm phán ký Thông báo (`metadata.ho_ten_tham_phan`).")

    elif doc_type == "van_ban_trung_gian":
        vbtg = norm_data.get("van_ban_trung_gian") or {}
        if not vbtg.get("kinh_gui"):
            (warnings if allow_draft else errors).append("Thiếu nơi nhận chính (`van_ban_trung_gian.kinh_gui`).")
        if not vbtg.get("noi_dung"):
            (warnings if allow_draft else errors).append("Thiếu nội dung văn bản trung gian (`van_ban_trung_gian.noi_dung`).")

    return errors, warnings
