# -*- coding: utf-8 -*-
"""
universal_visual_filler.py - Động cơ xuất văn bản theo Phôi Mẫu Word Trực Quan (Visual Template)
- Hỗ trợ Tờ trình Chánh án, Quyết định Kháng nghị và Thông báo giải quyết đơn.
- Tốc độ xuất: < 0.2 giây, bảo toàn 100% đường kẻ VML và định dạng nội dòng từng run.
- Duyệt đệ quy bảng lồng đa cấp và toàn bộ header/footer.
- Tích hợp --dry-run, File-lock check và chặn xuất file khi còn sót {{placeholder}} / [CẦN ĐIỀN...].
"""
import argparse
import json
import os
import sys
import time
import docx

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import paths_config as cfg  # noqa: F401
import docx_utils as du

du.ensure_utf8_stdout()

check_file_lock = du.check_file_lock
collect_all_text = du.collect_all_text


def replace_tag_in_para(p, tag, value):
    """Tương thích ngược: thay 1 tag đơn lẻ bảo toàn định dạng từng run và VML."""
    return du.replace_tags_preserve_format(p, {tag: value})


def fill_template(template_path, data, output_path, dry_run=False, allow_draft=False):
    t0 = time.time()
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Không tìm thấy file phôi mẫu: {template_path}")

    doc = docx.Document(template_path)
    replacements = {
        str(tag): str(val)
        for tag, val in data.items()
        if isinstance(val, (str, int, float))
    }

    for p in du.iter_document_paragraphs(doc):
        du.replace_tags_preserve_format(p, replacements)

    placeholders, unfilled = du.scan_unfilled_tokens(collect_all_text(doc))
    if placeholders:
        msg = f"Còn {len(placeholders)} placeholder chưa được điền: {placeholders}"
        if allow_draft:
            print(f"[WARN] {msg}")
        else:
            raise ValueError(msg)
    if unfilled:
        msg = f"Còn thẻ thiếu dữ liệu bắt buộc: {unfilled}"
        if allow_draft:
            print(f"[WARN] {msg}")
        else:
            raise ValueError(msg)

    if dry_run:
        print(f"[DRY-RUN OK] Kiểm tra phôi trực quan hoàn tất (chưa ghi file): {output_path}")
        return

    ok, lock_err = check_file_lock(output_path)
    if not ok:
        raise PermissionError(lock_err)

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    doc.save(output_path)
    elapsed = time.time() - t0
    print(f"[XUẤT THÀNH CÔNG] Trong {elapsed:.3f} giây!")
    print(f"  + Phôi mẫu : {template_path}")
    print(f"  + Đầu ra   : {output_path}")


def main():
    parser = argparse.ArgumentParser(description="Universal Visual Template Filler for TANDTC")
    parser.add_argument("--template", required=True, help="Đường dẫn file Word mẫu trực quan (.docx)")
    parser.add_argument("--data", required=True, help="Đường dẫn file JSON dữ liệu")
    parser.add_argument("--output", required=True, help="Đường dẫn file Word đầu ra (.docx)")
    parser.add_argument("--dry-run", action="store_true", help="Chỉ kiểm tra dữ liệu, không ghi file")
    parser.add_argument("--allow-draft", action="store_true", help="Cho phép xuất bản nháp dù còn placeholder/[CẦN ĐIỀN]")
    args = parser.parse_args()

    with open(args.data, "r", encoding="utf-8") as f:
        data = json.load(f)

    fill_template(args.template, data, args.output, dry_run=args.dry_run, allow_draft=args.allow_draft)


if __name__ == "__main__":
    main()
