# -*- coding: utf-8 -*-
"""
Trích xuất phôi chuẩn base_khang_nghi_tandtc.docx trực tiếp từ Bản phát hành gốc
nhằm bảo toàn 100% 2 đường kẻ ngang VML dưới Tên cơ quan & Tiêu ngữ, styles.xml,
Normal style (14pt, giãn dòng 1.15), lề trang A4 và footer1.xml/footer2.xml.
"""
import glob
import os
import shutil
import sys
from docx import Document
from docx.oxml.ns import qn

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import paths_config as cfg

_candidates = glob.glob(os.path.join(cfg.OUTPUT_ROOT_DIR, "*Vu_Lam_Thi_Deo_Vu_Thanh_Do", "DT KN Đèo - Đô*.docx"))
SRC_DOCX = _candidates[0] if _candidates else os.path.join(cfg.OUTPUT_ROOT_DIR, "08_Vu_Lam_Thi_Deo_Vu_Thanh_Do", "DT KN Đèo - Đô - sửa 14.4 (TP Thuỷ + PCA Tiến sửa) - BẢN PHÁT HÀNH.docx")
DST_DIR = cfg.TEMPLATES_DIR
DST_DOCX = os.path.join(DST_DIR, "base_khang_nghi_tandtc.docx")


def build_base_khang_nghi():
    os.makedirs(DST_DIR, exist_ok=True)
    if not os.path.exists(SRC_DOCX):
        if os.path.exists(DST_DOCX):
            print(f"[OK] Phôi chuẩn đã tồn tại sẵn tại: {DST_DOCX}")
            return
        raise FileNotFoundError(f"Không tìm thấy file nguồn gốc: {SRC_DOCX}")
    shutil.copy2(SRC_DOCX, DST_DOCX)

    doc = Document(DST_DOCX)
    body = doc.element.body

    # Giữ lại chính xác 9 phần tử đầu của Bản phát hành gốc:
    # [0] Table 0 (TÒA ÁN NHÂN DÂN TỐI CAO / Số: ... | Quốc hiệu - Tiêu ngữ / Hà Nội, ngày ... + 2 VML lines)
    # [1] Dòng trống đệm 1
    # [2] Dòng trống đệm 2
    # [3] QUYẾT ĐỊNH
    # [4] KHÁNG NGHỊ GIÁM ĐỐC THẨM
    # [5] Đối với Bản án dân sự phúc thẩm số ... ngày ...
    # [6] của Tòa án nhân dân ...
    # [7] CHÁNH ÁN
    # [8] TÒA ÁN NHÂN DÂN TỐI CAO
    # và phần tử cuối <w:sectPr>
    children = list(body)
    keep_count = 9
    for el in children[keep_count:]:
        if el.tag == qn("w:sectPr"):
            continue
        body.remove(el)

    doc.save(DST_DOCX)
    print(f"[OK] Đã tạo phôi chuẩn Quyết định Kháng nghị XML gốc tại: {DST_DOCX}")


if __name__ == "__main__":
    build_base_khang_nghi()
