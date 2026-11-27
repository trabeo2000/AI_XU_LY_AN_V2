# 03_Ket_qua_Xu_ly — KẾT QUẢ XỬ LÝ HỒ SƠ (XẾP NGÀY TẠO MỚI NHẤT Ở TRÊN CÙNG)

Thư mục tổng hợp toàn bộ sản phẩm văn bản đầu ra của hệ thống (Tờ trình, Quyết định Kháng nghị, Thông báo giải quyết đơn, Báo cáo nghiên cứu 6 bước).

## Quy tắc tổ chức & Sắp xếp tự động
- **Định dạng tên thư mục con:** `<Số>_Vu_<Ten_ho_so>/` (ví dụ: `01_Vu_Ho_Van_Phon_Pham_Thi_Vong/`).
- **Thứ tự sắp xếp:** Số **`01_`** luôn là hồ sơ có **ngày tạo gần hiện tại nhất**, giảm dần theo thời gian tạo; thư mục mẫu demo luôn nằm cuối (`99_Mau_Template_Demo/`).
- **Cơ chế tự động:** Mỗi khi [`export_dossier.py`](../01_Flow_Xu_Ly_Don/scripts/export_dossier.py) xuất hồ sơ mới ở Bước 6, hệ thống tự động ghi nhận ngày tạo vào [`01_Flow_Xu_Ly_Don/case_registry.json`](../01_Flow_Xu_Ly_Don/case_registry.json), đánh lại số thứ tự để vụ mới nhất nhảy lên vị trí `01_` trên cùng và cập nhật bảng mục lục bên dưới.

## Danh mục hồ sơ hiện có (Mới nhất → Cũ nhất)

| STT | Ngày tạo | Thư mục hồ sơ | Số văn bản |
| :---: | :---: | :--- | :---: |
| 01 | 27/11/2026 | [`01_Vu_Le_Thi_Hoan_Thai_Quyen_Anh/`](01_Vu_Le_Thi_Hoan_Thai_Quyen_Anh/) | 3 |
| 02 | 27/11/2026 | [`02_Vu_Ho_Van_Phon_Pham_Thi_Vong/`](02_Vu_Ho_Van_Phon_Pham_Thi_Vong/) | 1 |
| 03 | 26/11/2026 | [`03_Vu_Tran_Thi_Lien_Nguyen_Thi_Huong/`](03_Vu_Tran_Thi_Lien_Nguyen_Thi_Huong/) | 2 |
| 04 | 26/11/2026 | [`04_Vu_Vo_Van_Ut_Vo_Thi_Ut/`](04_Vu_Vo_Van_Ut_Vo_Thi_Ut/) | 5 |
| 05 | 26/11/2026 | [`05_Vu_Bui_Van_Cuoc_Pham_Thi_Qui/`](05_Vu_Bui_Van_Cuoc_Pham_Thi_Qui/) | 2 |
| 06 | 26/11/2026 | [`06_Vu_Le_Thi_Tuyet_Mai_Au_Thi_Kieu_Trang/`](06_Vu_Le_Thi_Tuyet_Mai_Au_Thi_Kieu_Trang/) | 5 |
| 07 | 25/11/2026 | [`07_Vu_Phan_Van_Bay_Tran_Van_Mung/`](07_Vu_Phan_Van_Bay_Tran_Van_Mung/) | 1 |
| 08 | 05/11/2026 | [`08_Vu_Nguyen_Ngoc_Hien_Nguyen_Kien_Thanh/`](08_Vu_Nguyen_Ngoc_Hien_Nguyen_Kien_Thanh/) | 4 |
| 09 | 07/04/2026 | [`09_Vu_Lam_Thi_Deo_Vu_Thanh_Do/`](09_Vu_Lam_Thi_Deo_Vu_Thanh_Do/) | 1 |
| 99 | Mẫu | [`99_Mau_Template_Demo/`](99_Mau_Template_Demo/) | 1 |
