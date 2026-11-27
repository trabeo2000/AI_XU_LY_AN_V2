# AI_XU_LY_AN
Hệ thống AI Hỗ trợ Thẩm tra viên & Thư ký Tòa án nhân dân tối cao (TANDTC) xử lý đơn khiếu nại, đề nghị giám đốc thẩm, tái thẩm.

## Bộ công cụ hỗ trợ siêu tốc (100% Portable & Offline):
1. **Động cơ xuất hồ sơ toàn diện All-in-One (`export_dossier.py`)**:
   - Xuất trọn gói Tờ trình, Kháng nghị, Thông báo giải quyết đơn và Báo cáo 6 bước (hỗ trợ cả 2 phương án cùng lúc) trong < 0.5s qua 1 lệnh CLI duy nhất.
2. **Bộ đọc hồ sơ & tài liệu siêu tốc (`fast_reader.py`)**:
   - Trích xuất Word (`.docx`), Excel (`.xlsx`), PDF số trong 0.1 - 0.2s.
   - Xử lý PDF scan 100-200 trang bằng cơ chế định vị trúng đích & ghép trang thông minh (`--stitch 2`), giảm 80% thời gian đọc.
3. **Bộ xuất Tờ trình Chánh án TANDTC (`export_to_trinh.py`)**:
   - Xuất file `.docx` chuẩn 100% phôi gốc (giữ 2 đường kẻ VML, lề, font Times New Roman, styles) trong < 1s.
4. **Bộ xuất Quyết định Kháng nghị GĐT / Tái thẩm (`export_khang_nghi.py`)**:
   - Xuất file `.docx` chuẩn bản phát hành gốc trong < 1s.
5. **Bộ xuất Báo cáo Toàn diện 6 bước (`export_6step_report.py`)**:
   - Xuất file `.docx` Báo cáo nghiên cứu hồ sơ từ Bước 1 đến Bước 6 (bảng ma trận, callout box, header TANDTC) trong < 1s.
6. **Động cơ Phôi Mẫu Word Trực Quan (`universal_visual_filler.py`)**:
   - Tự do mở trực tiếp phôi Word mẫu (`demo_phoi_to_trinh_truc_quan.docx`, `demo_phoi_khang_nghi_truc_quan.docx`, `demo_phoi_thong_bao_truc_quan.docx`) để chỉnh sửa kiểu dáng, thêm bớt mục bằng chuột; máy tự động rót dữ liệu trong < 0.1s mà không cần sửa bất kỳ dòng code nào.

7. **Kiểm định toàn vẹn & cấu hình đường dẫn tập trung (`paths_config.py`, `verify_system.py`)**:
   - Tập trung toàn bộ đường dẫn tương đối tại `paths_config.py`; chạy `python 01_Flow_Xu_Ly_Don/scripts/verify_system.py` sau mỗi lần tái cấu trúc để xác thực 100% link, phôi Word, placeholder, cú pháp Python và thứ tự hồ sơ.

## Cấu trúc thư mục Workspace:
- `01_Flow_Xu_Ly_Don/`: Quy trình xử lý đơn, các biểu mẫu chuẩn, `case_registry.json` và script công cụ lõi. *(Xem sơ đồ luồng dữ liệu tại [01_Flow_Xu_Ly_Don/README.md](01_Flow_Xu_Ly_Don/README.md) và [So_do_kien_truc_luong_file.svg](01_Flow_Xu_Ly_Don/So_do_kien_truc_luong_file.svg))*.
- `02_Ho_so/`: Chứa hồ sơ tài liệu các vụ việc cần xử lý.
- `03_Ket_qua_Xu_ly/`: Thư mục tổng hợp đầu ra (.docx) phân chia thành từng thư mục con theo từng hồ sơ vụ án (`<Số>_Vu_<Ten_vu>/`, tự động xếp `01_` ngày tạo mới nhất ở trên cùng), chứa trọn bộ Tờ trình Chánh án, Dự thảo Quyết định Kháng nghị GĐT/Tái thẩm, Thông báo giải quyết đơn (Trả lời đơn - TLĐ) và Báo cáo nghiên cứu 6 bước.
- `04_Van_ban_phap_ly/`: Thư viện văn bản pháp luật, Nghị quyết HĐTP (gồm NQ 01/2025/NQ-HĐTP), án lệ, giải đáp nghiệp vụ.
- `05_Quy_trinh_va_Pham_vi_Chuc_danh/`: Quy trình làm việc nội bộ (QC 75, TB 383, TB 679), cẩm nang nghiệp vụ và phạm vi thẩm quyền các cấp chức danh.
