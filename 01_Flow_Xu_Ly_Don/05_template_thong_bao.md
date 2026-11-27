# 05. TEMPLATE THÔNG BÁO GIẢI QUYẾT ĐƠN ĐỀ NGHỊ GIÁM ĐỐC THẨM / TÁI THẨM (SINGLE SOURCE OF TRUTH)

Tài liệu này là **Single Source of Truth duy nhất** tổng hợp toàn bộ cấu trúc, thông số định dạng XML gốc, phông chữ, cỡ chữ, căn lề, kích thước bảng (`cm` / `dxa`), đường kẻ ngang phân cách VML, cấu trúc trường thông tin đầu vào và quy trình xuất file Word `.docx` chuẩn 100% cho **Thông báo giải quyết đơn đề nghị Giám đốc thẩm / Tái thẩm (Trả lời đơn - TLĐ)** do Thẩm phán TANDTC trực tiếp duyệt ký theo tinh thần cải cách thủ tục hành chính tư pháp tại **Thông báo số 383/TB-TANDTC và Thông báo số 679/TB-TANDTC**.

> 🔴 **QUY TẮC BẮT BUỘC (MANDATORY ENGINE):**
> - Khi xuất Thông báo giải quyết đơn `.docx` ở Bước 6 (lưu vào [`03_Ket_qua_Xu_ly/`](../03_Ket_qua_Xu_ly/) theo thư mục con từng vụ án `<Số>_Vu_<ten_vu>/`), hệ thống hỗ trợ 2 chế độ:
>   1. **Chế độ Phôi Mẫu Trực Quan (Visual Template):** Sử dụng [`universal_visual_filler.py`](scripts/universal_visual_filler.py) rót dữ liệu vào [`demo_phoi_thong_bao_truc_quan.docx`](templates/demo_phoi_thong_bao_truc_quan.docx). Thẩm tra viên có thể tự do mở file Word này lên chỉnh sửa kiểu dáng, font chữ, thêm bớt mục bằng chuột mà không cần sửa code.
>   2. **Chế độ Exporter Tự Động Siêu Tốc:** Sử dụng [`export_dossier.py`](scripts/export_dossier.py) (lựa chọn **Option 4** hoặc **Option 6**) nạp phôi gốc [`base_thong_bao_tld_tandtc.docx`](templates/base_thong_bao_tld_tandtc.docx) để rót dữ liệu trong `< 0.2 giây`.
> - **TUYỆT ĐỐI NGHIÊM CẤM** dựng chay từ trang Word trống hoặc dùng PowerShell COM scripting.

---

## 1. Thiết Lập Trang In, Phông Chữ & Căn Lề Chuẩn Tư Pháp (Page Setup & Typography)

| Thuộc tính | Thông số chuẩn 100% (Trích xuất từ phôi mẫu gốc) |
| :--- | :--- |
| **Khổ giấy** | A4 dọc — Rộng `21.0 cm` × Cao `29.7 cm` |
| **Lề trang (Margins)** | Trên `2.0 cm` · Dưới `2.0 cm` · Trái `3.0 cm` (chuẩn đóng gáy hồ sơ) · Phải `2.0 cm` |
| **Phông chữ toàn văn bản** | **Times New Roman** (chuẩn Unicode TCVN) |
| **Cỡ chữ mặc định thân bài** | **`14pt`**, căn đều 2 bên (`Justify`), thụt đầu dòng **`1.06 cm`**, giãn dòng **`1.15 lines`** |
| **Khoảng cách đoạn (Spacing)** | `Space Before = 6pt`, `Space After = 6pt` (riêng đoạn mở đầu `Before = 12pt`, đoạn kết `After = 12pt`) |

---

## 2. Quy Định Chi Tiết 5 Khối Cấu Trúc Thành Phần

### 🏛️ Khối 1: Bảng Quốc Hiệu – Tiêu Ngữ – Cơ Quan Ban Hành – Ngày Tháng *(Bảng 1)*
- **Cấu trúc bảng:** 1 hàng × 2 cột, **KHÔNG VIỀN** (`borders = none`).
  - **Cột 1 (Trái - Rộng `6.44 cm`):**
    - Dòng 1: `TÒA ÁN NHÂN DÂN TỐI CAO` — **`12pt`, IN ĐẬM, IN HOA**, căn giữa (`Center`).
    - Dòng 2: Đường kẻ ngang VML phân cách màu đen nét liền căn giữa bên dưới.
    - Dòng 3: `Số: {{SO_THONG_BAO}}/TB-TA` — **`14pt`, thường (Regular)**, căn giữa (`Center`).
  - **Cột 2 (Phải - Rộng `9.44 cm`):**
    - Dòng 1: `CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM` — **`12pt`, IN ĐẬM, IN HOA**, căn giữa (`Center`).
    - Dòng 2: `Độc lập - Tự do - Hạnh phúc` — **`14pt`, IN ĐẬM**, căn giữa (`Center`).
    - Dòng 3: Đường kẻ ngang VML phân cách màu đen nét liền căn giữa bên dưới.
    - Dòng 4: *`Hà Nội, ngày {{NGAY}} tháng {{THANG}} năm {{NAM}}`* — **`14pt`, IN NGHIÊNG**, căn giữa (`Center`).

---

### 📄 Khối 2: Tiêu Đề Văn Bản *(Căn giữa toàn trang)*
1. **`THÔNG BÁO`** — **`14pt`, IN ĐẬM, IN HOA**, căn giữa (`Center`), `Space Before = 18pt`, `Space After = 0pt`.
2. **`GIẢI QUYẾT ĐƠN ĐỀ NGHỊ {{THU_TUC_VIET_HOA}}`** *(Mặc định: `GIÁM ĐỐC THẨM`)* — **`14pt`, IN ĐẬM, IN HOA**, căn giữa (`Center`), `Space Before = 0pt`, `Space After = 12pt`, giãn dòng `1.5 lines`.

---

### 📬 Khối 3: Bảng Kính Gửi & Địa Chỉ Người Nhận Đơn *(Bảng 2)*
- **Cấu trúc bảng:** 2 hàng × 2 cột, **KHÔNG VIỀN** (`borders = none`), giãn dòng `1.15 lines`.
  - **Cột nhãn (Trái - Rộng `4.74 cm`, Căn phải - Align Right):**
    - Hàng 1: `Kính gửi:` — **`14pt`, thường**.
    - Hàng 2: `Địa chỉ:` — **`14pt`, thường**.
  - **Cột nội dung (Phải - Rộng `11.26 cm`, Căn trái - Align Left):**
    - Hàng 1: `{{NGUOI_GUI_DON}}` — **`14pt`, thường/đậm**.
    - Hàng 2: `{{DIA_CHI_NGUOI_GUI}}` — **`14pt`, thường** *(bao gồm địa chỉ hiện tại và địa chỉ cũ nếu có thay đổi trong hồ sơ)*.

---

### 📝 Khối 4: Thân Bài Lập Luận Bác Đơn (4 Đoạn Chuẩn Nghiệp Vụ)
- **Định dạng chung:** Font **`Times New Roman 14pt`**, Căn đều 2 bên (`Justify`), Thụt lề đầu dòng **`1.06 cm`**, Giãn dòng **`1.15 lines`**.
- **4 Đoạn nội dung logic:**
  1. **Đoạn 1 (Xác nhận nhận đơn & Bản án bị khiếu nại):** Nêu rõ TANDTC nhận được đơn ngày nào, đương sự đề nghị xem xét bản án sơ thẩm/phúc thẩm số bao nhiêu, ngày nào của Tòa án nào.
  2. **Đoạn 2 (Tóm tắt quan hệ tranh chấp & Tình tiết cốt lõi vụ án):** Khái quát nguồn gốc tài sản, các giao dịch phân chia, sự thỏa thuận của các bên đương sự trong quá khứ.
  3. **Đoạn 3 (Lập luận bác đơn & Đánh giá chứng cứ):** Phân tích chi tiết tại sao phán quyết của Tòa án cấp sơ thẩm/phúc thẩm là có căn cứ, đúng pháp luật; chỉ rõ yêu cầu khiếu nại của đương sự là không có cơ sở chấp nhận.
  4. **Đoạn 4 (Kết luận & Thông báo):**
     - *"Vì vậy, không có căn cứ kháng nghị theo thủ tục giám đốc thẩm đối với Bản án dân sự phúc thẩm số... nêu trên."*
     - *"Tòa án nhân dân tối cao thông báo để ông/bà được biết."*

---

### ✍️ Khối 5: Nơi Nhận & Thẩm Quyền Ký Thẩm Phán TANDTC *(Bảng 3)*
- **Cấu trúc bảng:** 1 hàng × 2 cột, **KHÔNG VIỀN** (`borders = none`).
  - **Cột 1 (Nơi nhận - Rộng `8.02 cm`, Căn đều/Trái):**
    - Dòng 1: *`Nơi nhận:`* — **`12pt`, IN ĐẬM, IN NGHIÊNG**.
    - Dòng 2 đến 9: Danh sách các cơ quan nhận — **`11pt`, thường**, giãn dòng `1.1 lines`:
      - `- Như kính gửi;`
      - `- Đ/c Chánh án TANDTC (để b/c);`
      - `- Đ/c PCA TANDTC phụ trách (để b/c);`
      - `- {{VU_VKSNDTC}} (để biết);`
      - `- {{TOA_AN_TINH}} (kèm hồ sơ vụ án);`
      - `- {{TOA_AN_HUYEN}};`
      - `- {{CO_QUAN_THADS}};`
      - `- Lưu: VP, Thẩm phán TANDTC {{TEN_THAM_PHAN}}, {{VU_GDKT}} (03 bản).`
  - **Cột 2 (Chữ ký Thẩm phán - Rộng `9.05 cm`, Căn giữa - Center):**
    - Dòng 1: **`TUQ. CHÁNH ÁN`** — **`13pt`, IN ĐẬM, IN HOA**, căn giữa.
    - Dòng 2: **`THẨM PHÁN TÒA ÁN NHÂN DÂN TỐI CAO`** *(hoặc **`THẨM PHÁN TÒA ÁN NHÂN DÂN BẬC 3`** — `export_dossier.py` tự động chuyển đổi đồng bộ cả dòng chữ ký và dòng `Lưu: VP, ...` khi `metadata.chuc_danh_tham_phan` chứa `bậc 3`)*. — **`12pt`, IN ĐẬM, IN HOA**, căn giữa, **`Space After = 84pt`** (tạo khoảng trống ký tên 6-7 dòng).
    - Dòng 3: **`{{TEN_THAM_PHAN}}`** — **`14pt`, IN ĐẬM**, căn giữa.
