# ⚖️ 04_04 — TUYỂN TẬP ÁN LỆ TÒA ÁN NHÂN DÂN TỐI CAO (TANDTC)

Thư mục lưu trữ hệ thống **Án lệ chính thức do Hội đồng Thẩm phán TANDTC công bố** (từ Án lệ số 01 trở đi), phục vụ đối chiếu bắt buộc tại **Bước 1B**, **Bước 4** và **Bước 5** trong quy trình xử lý đơn GĐT/Tái thẩm.

## Hướng dẫn nạp file Án lệ gốc vào kho
1. Chép các file Án lệ gốc (`.pdf`, `.docx`, `.doc`) vào thư mục [`_Ban_goc/`](_Ban_goc/).
2. Chạy lệnh đồng bộ tự động để chuyển đổi toàn bộ sang Markdown 2 tầng siêu tốc (có mục lục tra cứu nhanh & gắn nhãn `⚖️ [ÁN LỆ TANDTC]`):
   ```bash
   python 01_Flow_Xu_Ly_Don/scripts/sync_legal_knowledge.py --sync
   ```
3. Kiểm tra lại toàn vẹn kho bằng:
   ```bash
   python 01_Flow_Xu_Ly_Don/scripts/verify_system.py
   ```
