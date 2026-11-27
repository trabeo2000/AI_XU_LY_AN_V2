# -*- coding: utf-8 -*-
"""
Trích xuất phôi chuẩn base_khang_nghi_tandtc.docx trực tiếp từ Bản phát hành gốc
nhằm bảo toàn 100% 2 đường kẻ ngang VML dưới Tên cơ quan & Tiêu ngữ, styles.xml,
Normal style (14pt, giãn dòng 1.15), lề trang A4 và footer1.xml/footer2.xml.
"""
import argparse
import glob
import os
import shutil
import sys
from docx import Document
from docx.oxml.ns import qn

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import paths_config as cfg
import docx_utils as du

du.ensure_utf8_stdout()

_candidates = glob.glob(os.path.join(cfg.OUTPUT_ROOT_DIR, "*Vu_Lam_Thi_Deo_Vu_Thanh_Do", "DT KN Đèo - Đô*.docx"))
DEFAULT_SRC_DOCX = _candidates[0] if _candidates else ""
DST_DIR = cfg.TEMPLATES_DIR
DST_DOCX = cfg.BASE_KHANG_NGHI_TEMPLATE


def build_base_khang_nghi(src_docx=None):
    os.makedirs(DST_DIR, exist_ok=True)
    source = src_docx or DEFAULT_SRC_DOCX
    if not source or not os.path.exists(source):
        if os.path.exists(DST_DOCX):
            print(f"[OK] Phôi chuẩn đã tồn tại sẵn tại: {DST_DOCX}")
            return
        raise FileNotFoundError(
            "Không tìm thấy file nguồn gốc. Truyền đường dẫn qua `--source <file.docx>`."
        )
    shutil.copy2(source, DST_DOCX)

    doc = Document(DST_DOCX)
    body = doc.element.body

    children = list(body)
    keep_count = 9
    for el in children[keep_count:]:
        if el.tag == qn("w:sectPr"):
            continue
        body.remove(el)

    doc.save(DST_DOCX)
    print(f"[OK] Đã tạo phôi chuẩn Quyết định Kháng nghị XML gốc tại: {DST_DOCX}")


def main():
    parser = argparse.ArgumentParser(description="Tái tạo phôi chuẩn base_khang_nghi_tandtc.docx từ Bản phát hành gốc")
    parser.add_argument("--source", "-s", help="Đường dẫn file Quyết định Kháng nghị gốc (tùy chọn)")
    args = parser.parse_args()
    build_base_khang_nghi(src_docx=args.source)


if __name__ == "__main__":
    main()
