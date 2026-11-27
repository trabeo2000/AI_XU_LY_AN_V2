# -*- coding: utf-8 -*-
"""
Trích xuất phôi chuẩn base_to_trinh_tandtc.docx trực tiếp từ file gốc TANDTC
nhằm bảo toàn 100% docDefaults, styles.xml, footer1.xml, lề trang A4 và
2 đường kẻ ngang VML dưới Tên cơ quan & Tiêu ngữ.
"""
import glob
import os
import shutil
import sys
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import paths_config as cfg

_candidates = glob.glob(os.path.join(cfg.OUTPUT_ROOT_DIR, "*Vu_Nguyen_Ngoc_Hien_Nguyen_Kien_Thanh", "To_trinh_*.docx"))
SRC_DOCX = _candidates[0] if _candidates else os.path.join(cfg.OUTPUT_ROOT_DIR, "07_Vu_Nguyen_Ngoc_Hien_Nguyen_Kien_Thanh", "To_trinh_Nguyễn_Ngọc_Hiền_Nguyễn_Kiến_Thành.docx")
DST_DIR = cfg.TEMPLATES_DIR
DST_DOCX = os.path.join(DST_DIR, "base_to_trinh_tandtc.docx")


def build_base_template():
    os.makedirs(DST_DIR, exist_ok=True)
    if not os.path.exists(SRC_DOCX):
        if os.path.exists(DST_DOCX):
            print(f"[OK] Phôi chuẩn đã tồn tại sẵn tại: {DST_DOCX}")
            return
        raise FileNotFoundError(f"Không tìm thấy file nguồn gốc: {SRC_DOCX}")
    shutil.copy2(SRC_DOCX, DST_DOCX)

    doc = Document(DST_DOCX)
    body = doc.element.body

    # Giữ lại chính xác 9 phần tử đầu:
    # [0] Table 0 (Quốc hiệu, Tiêu ngữ, Tên cơ quan, Ngày tháng + 2 đường VML)
    # [1] Dòng trống đệm
    # [2] TỜ TRÌNH
    # [3] CHÁNH ÁN TÒA ÁN NHÂN DÂN TỐI CAO
    # [4] Về vụ án dân sự “...”
    # [5] (Bản án dân sự phúc thẩm số ...
    # [6] ngày ... của Tòa án nhân dân ...)
    # [7] --------------------------------
    # [8] Table 1 (Bảng chia đôi: Khung TKV bên trái | Đương sự bên phải)
    # và phần tử cuối <w:sectPr>
    children = list(body)
    keep_count = 9
    for el in children[keep_count:]:
        if el.tag == qn("w:sectPr"):
            continue
        body.remove(el)

    # Chuẩn hóa nhẹ Table 1: chỉ giữ 1 hàng duy nhất (xóa hàng trống thứ 2 nếu có)
    tbl1 = doc.tables[1]._tbl
    trs = tbl1.findall(qn("w:tr"))
    for extra_tr in trs[1:]:
        tbl1.remove(extra_tr)

    # Chuẩn hóa lề Footer = 0.50 cm (284 dxa) cho đẹp và đồng nhất
    sectPr = body.find(qn("w:sectPr"))
    if sectPr is not None:
        pgMar = sectPr.find(qn("w:pgMar"))
        if pgMar is not None:
            pgMar.set(qn("w:footer"), "284")

    doc.save(DST_DOCX)
    print(f"[OK] Đã tạo phôi chuẩn XML gốc tại: {DST_DOCX}")


if __name__ == "__main__":
    build_base_template()
