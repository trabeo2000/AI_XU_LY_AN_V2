# 03_Ket_qua_Xu_ly — KẾT QUẢ XỬ LÝ HỒ SƠ (XẾP NGÀY TẠO MỚI NHẤT Ở TRÊN CÙNG)

Thư mục tổng hợp toàn bộ sản phẩm văn bản đầu ra của hệ thống (Tờ trình, Quyết định Kháng nghị, Thông báo giải quyết đơn, Báo cáo nghiên cứu 6 bước).

## Quy tắc tổ chức & Sắp xếp tự động
- **Định dạng tên thư mục con:** `<Số>_Vu_<Ten_ho_so>/` (ví dụ: `01_Vu_Ho_Van_Phon_Pham_Thi_Vong/`).
- **Thứ tự sắp xếp:** Số **`01_`** luôn là hồ sơ có **ngày tạo gần hiện tại nhất**, giảm dần theo thời gian tạo; thư mục mẫu demo luôn nằm cuối (`99_Mau_Template_Demo/`).
- **Cơ chế tự động:** Mỗi khi [`export_dossier.py`](../01_Flow_Xu_Ly_Don/scripts/export_dossier.py) xuất hồ sơ mới ở Bước 6, hệ thống tự động ghi nhận ngày tạo vào [`01_Flow_Xu_Ly_Don/case_registry.json`](../01_Flow_Xu_Ly_Don/case_registry.json), đánh lại số thứ tự để vụ mới nhất nhảy lên vị trí `01_` trên cùng và cập nhật bảng mục lục bên dưới.

## Danh mục hồ sơ hiện có (Mới nhất → Cũ nhất)

| STT | Ngày tạo | Thư mục hồ sơ | Số văn bản |
| :---: | :---: | :--- | :---: |
| 01 | 27/11/2026 | [`01_Vu_Duong_Thi_Ty_La_Thi_Loan_Khong_KN/`](01_Vu_Duong_Thi_Ty_La_Thi_Loan_Khong_KN/) | 2 |
| 02 | 27/11/2026 | [`02_Vu_Nguyen_Thai_Y_Nguyen_Thi_Nguyet/`](02_Vu_Nguyen_Thai_Y_Nguyen_Thi_Nguyet/) | 4 |
| 03 | 27/11/2026 | [`03_Vu_Le_Thi_Tim_Nguyen_Thi_Tuyet/`](03_Vu_Le_Thi_Tim_Nguyen_Thi_Tuyet/) | 4 |
| 04 | 27/11/2026 | [`04_Vu_Duong_Thi_Ty_La_Thi_Loan/`](04_Vu_Duong_Thi_Ty_La_Thi_Loan/) | 4 |
| 05 | 27/11/2026 | [`05_Vu_Le_Thi_Hoan_Thai_Quyen_Anh/`](05_Vu_Le_Thi_Hoan_Thai_Quyen_Anh/) | 3 |
| 06 | 27/11/2026 | [`06_Vu_Ho_Van_Phon_Pham_Thi_Vong/`](06_Vu_Ho_Van_Phon_Pham_Thi_Vong/) | 1 |
| 07 | 26/11/2026 | [`07_Vu_Tran_Thi_Lien_Nguyen_Thi_Huong/`](07_Vu_Tran_Thi_Lien_Nguyen_Thi_Huong/) | 7 |
| 08 | 26/11/2026 | [`08_Vu_Vo_Van_Ut_Vo_Thi_Ut/`](08_Vu_Vo_Van_Ut_Vo_Thi_Ut/) | 5 |
| 09 | 26/11/2026 | [`09_Vu_Bui_Van_Cuoc_Pham_Thi_Qui/`](09_Vu_Bui_Van_Cuoc_Pham_Thi_Qui/) | 2 |
| 10 | 26/11/2026 | [`10_Vu_Le_Thi_Tuyet_Mai_Au_Thi_Kieu_Trang/`](10_Vu_Le_Thi_Tuyet_Mai_Au_Thi_Kieu_Trang/) | 5 |
| 11 | 25/11/2026 | [`11_Vu_Phan_Van_Bay_Tran_Van_Mung/`](11_Vu_Phan_Van_Bay_Tran_Van_Mung/) | 1 |
| 12 | 05/11/2026 | [`12_Vu_Nguyen_Ngoc_Hien_Nguyen_Kien_Thanh/`](12_Vu_Nguyen_Ngoc_Hien_Nguyen_Kien_Thanh/) | 4 |
| 13 | 07/04/2026 | [`13_Vu_Lam_Thi_Deo_Vu_Thanh_Do/`](13_Vu_Lam_Thi_Deo_Vu_Thanh_Do/) | 1 |
| 99 | Mẫu | [`99_Mau_Template_Demo/`](99_Mau_Template_Demo/) | 1 |
