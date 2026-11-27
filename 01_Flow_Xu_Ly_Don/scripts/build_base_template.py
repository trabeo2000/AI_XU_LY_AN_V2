# -*- coding: utf-8 -*-
"""
Trích xuất phôi chuẩn base_to_trinh_tandtc.docx trực tiếp từ file gốc TANDTC
nhằm bảo toàn 100% docDefaults, styles.xml, footer1.xml, lề trang A4 và
2 đường kẻ ngang VML dưới Tên cơ quan & Tiêu ngữ.
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

_candidates = glob.glob(os.path.join(cfg.OUTPUT_ROOT_DIR, "*Vu_Nguyen_Ngoc_Hien_Nguyen_Kien_Thanh", "To_trinh_*.docx"))
DEFAULT_SRC_DOCX = _candidates[0] if _candidates else ""
DST_DIR = cfg.TEMPLATES_DIR
DST_DOCX = cfg.BASE_TO_TRINH_TEMPLATE


def build_base_template(src_docx=None):
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

    tbl1 = doc.tables[1]._tbl
    trs = tbl1.findall(qn("w:tr"))
    for extra_tr in trs[1:]:
        tbl1.remove(extra_tr)

    sectPr = body.find(qn("w:sectPr"))
    if sectPr is not None:
        pgMar = sectPr.find(qn("w:pgMar"))
        if pgMar is not None:
            pgMar.set(qn("w:footer"), "284")

    doc.save(DST_DOCX)
    print(f"[OK] Đã tạo phôi chuẩn XML gốc tại: {DST_DOCX}")


def main():
    parser = argparse.ArgumentParser(description="Tái tạo phôi chuẩn base_to_trinh_tandtc.docx từ file Word mẫu gốc")
    parser.add_argument("--source", "-s", help="Đường dẫn file Tờ trình Word gốc (tùy chọn)")
    args = parser.parse_args()
    build_base_template(src_docx=args.source)


if __name__ == "__main__":
    main()
