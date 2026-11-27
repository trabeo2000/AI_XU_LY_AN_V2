# 03. TEMPLATE TỜ TRÌNH CHÁNH ÁN TÒA ÁN NHÂN DÂN TỐI CAO (SINGLE SOURCE OF TRUTH)

Tài liệu này là **Single Source of Truth duy nhất** tổng hợp toàn bộ cấu trúc, thông số định dạng XML gốc (`word/document.xml`, `word/styles.xml`, `word/footer1.xml`), phông chữ, cỡ chữ, căn lề, kích thước bảng (`cm` / `dxa`), 2 đường kẻ ngang VML, cấu trúc JSON đầu vào và mã/quy trình xuất file Word `.docx` chuẩn 100% cho **Tờ trình Chánh án Tòa án nhân dân tối cao (Vụ Giám đốc kiểm tra II)**.

> 🔴 **QUY TẮC BẮT BUỘC (MANDATORY ENGINE):**
> - Khi xuất Tờ trình `.docx` ở Bước 6 (lưu vào [`03_Ket_qua_Xu_ly/`](../03_Ket_qua_Xu_ly/) theo thư mục con từng vụ án `<Số>_Vu_<ten_vu>/`), hệ thống hỗ trợ 2 chế độ:
>   1. **Chế độ Phôi Mẫu Trực Quan (Khuyên dùng):** Sử dụng [`universal_visual_filler.py`](scripts/universal_visual_filler.py) rót dữ liệu vào [`demo_phoi_to_trinh_truc_quan.docx`](templates/demo_phoi_to_trinh_truc_quan.docx). Thẩm tra viên có thể tự do mở file Word này lên chỉnh sửa kiểu dáng, font chữ, thêm bớt mục bằng chuột mà không cần sửa code.
>   2. **Chế độ Exporter Tự Động:** Sử dụng script [`export_to_trinh.py`](scripts/export_to_trinh.py) nạp phôi XML gốc [`base_to_trinh_tandtc.docx`](templates/base_to_trinh_tandtc.docx) và tự động dựng bảng biểu.
> - **TUYỆT ĐỐI NGHIÊM CẤM** dựng chay từ trang Word trống hoặc dùng PowerShell COM scripting.

---

## 1. Thiết Lập Trang In, Phông Chữ Mặc Định & Chân Trang (Page Setup & Typography)

| Thuộc tính | Thông số chuẩn 100% (trích xuất từ XML gốc) |
| :--- | :--- |
| **Khổ giấy** | A4 dọc — Rộng `21.0 cm` (`11906 dxa`) × Cao `29.7 cm` (`16838 dxa`) |
| **Lề trang (Margins)** | Trên `2.0 cm` (`1134 dxa`) · Dưới `2.0 cm` (`1134 dxa`) · Trái `3.0 cm` (`1701 dxa`) · Phải `2.0 cm` (`1134 dxa`) |
| **Header / Footer distance** | Header `1.27 cm` (`720 dxa`) · Footer `0.50 cm` (`284 dxa`) · Gutter `0 cm` |
| **Phông chữ toàn văn bản** | **Times New Roman** (khai báo đồng bộ cả 4 thuộc tính `w:ascii`, `w:hAnsi`, `w:cs`, `w:eastAsia`) |
| **Mặc định tài liệu (`docDefaults`)** | Cỡ `13pt`, căn đều 2 bên (`both`/`justify`), `Space After = 6pt`, giãn dòng đơn (`Single 1.0`), thụt đầu dòng `1.27 cm` |
| **Đoạn văn nội dung chính** | Cỡ **`14pt`**, căn đều 2 bên (`justify`), thụt đầu dòng **`1.25 cm`** (`709 dxa`), `Space Before = 0pt`, `Space After = 6pt`, giãn dòng đơn (`Single 1.0`), màu chữ `#000000` |
| **Chân trang (Footer)** | Field `PAGE` tại `word/footer1.xml`, căn giữa (`center`), `Times New Roman`, cỡ **`14pt`**, màu `#000000`, `Different first page (titlePg) = False` |

---

## 2. Quy Định Chi Tiết 7 Thành Phần & Định Dạng Từng Dòng

### 🏛️ Thành phần 1: Bảng Quốc Hiệu – Tiêu Ngữ – Đơn Vị Trình – Ngày Tháng *(Bảng 1)*
- **Cấu trúc bảng:** 2 cột × 2 hàng, **KHÔNG VIỀN** (`borders = none`), tổng chiều rộng **`16.76 cm`** (`9503 dxa`), `tblInd = 0`.
  - **Cột 1 (Trái):** `6.74 cm` (`3822 dxa`)
  - **Cột 2 (Phải):** `10.02 cm` (`5681 dxa`)
- **Hàng 1 – Cột 1** *(căn giữa, `firstLine = 0`, `After = 6pt`)*:
  - Dòng 1: `TOÀ ÁN NHÂN DÂN TỐI CAO` — **`13pt`, thường (không đậm), IN HOA**.
  - Dòng 2: **`{{DON_VI_TRINH}}`** *(mặc định: `VỤ GIÁM ĐỐC KIỂM TRA II`)* — **`13pt`, IN ĐẬM, IN HOA**, kèm **đường kẻ ngang VML màu đen (`Straight Connector 3`, nét `0.5pt`, `from="28.85pt,16.2pt" to="150.85pt,16.2pt"`)** căn giữa bên dưới (bảo toàn từ phôi XML gốc).
- **Hàng 1 – Cột 2** *(căn giữa, `firstLine = 0`, `After = 6pt`)*:
  - Dòng 1: **`CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM`** — **`13pt`, IN ĐẬM, IN HOA**.
  - Dòng 2: **`Độc lập - Tự do - Hạnh phúc`** — **`14pt`, IN ĐẬM**, kèm **đường kẻ ngang VML màu đen (`Straight Connector 4`, nét `0.25pt`, `from="70.05pt,1.1pt" to="202.05pt,1.1pt"`)** căn giữa bên dưới (bảo toàn từ phôi XML gốc).
- **Hàng 2 – Cột 1:** Để trống.
- **Hàng 2 – Cột 2:** *`Hà Nội, ngày {{NGAY}} tháng {{THANG}} năm {{NAM}}`* — **`14pt`, IN NGHIÊNG**, căn giữa, `firstLine = 0`, `Before = 12pt`, `After = 6pt`.

---

### 📄 Thành phần 2: Khối Tên Văn Bản & Trích Yếu Vụ Án *(Căn giữa, `firstLine = 0`)*
*(Sau Bảng 1 có 1 dòng trống đệm `after = 0pt`, `line = 1.15`)*
1. **`TỜ TRÌNH`** — **`14pt`, IN ĐẬM, IN HOA**, căn giữa, `Before = 12pt`, `After = 0pt`.
2. **`{{NGUOI_NHAN_TRINH}}`** *(mặc định: `CHÁNH ÁN TÒA ÁN NHÂN DÂN TỐI CAO`)* — **`14pt`, IN ĐẬM, IN HOA**, căn giữa, `Before = 0pt`, `After = 0pt`.
3. `Về vụ án {{LOAI_VU_AN}} ` (**`14pt`, đứng thường**) + *`“{{QUAN_HE_TRANH_CHAP}}”`* (**`14pt`, IN NGHIÊNG**) — căn giữa, `After = 0pt`.
4. Dòng 1 ngoặc đơn: `(Bản án {{LOAI_BAN_AN_PT}} số {{SO_BA_PT}}` — **`14pt`, đứng thường**, căn giữa, `After = 0pt`.
5. Dòng 2 ngoặc đơn: `ngày {{NGAY_BA_PT}} của {{TOA_AN_PT}})` — **`14pt`, đứng thường**, căn giữa, `After = 0pt`.
6. Đường phân cách ngang: `--------------------------------` *(đúng 32 dấu gạch ngang `-`)* — **`14pt`, đứng thường**, căn giữa, **giãn dòng `1.5` (`line = 360`)**, `After = 0pt`.

---

### 📊 Thành phần 3: Bảng Chia Đôi Đặc Thù TANDTC — Khung Tiến Độ Trình (Trái) | Đương Sự (Phải) *(Bảng 2)*
- **Bảng ngoài:** 2 cột × 1 hàng, cố định layout (`tblLayout = fixed`), **4 cạnh ngoài và đường ngang `nil`**, **CHỈ CÓ 1 đường kẻ dọc ở giữa** (`insideV: single, sz=4 (0.5pt), color=#000000`).
  - **Chế độ chuẩn (thẳng lề trang `3cm / 2cm` — mặc định):** Tổng rộng **`16.25 cm`** (`9214 dxa`), `tblInd = 0`:
    - Cột trái: **`7.50 cm`** (`4253 dxa`)
    - Cột phải: **`8.75 cm`** (`4961 dxa`)
  - **Chế độ rộng (`wide_party_table = true`, dùng khi vụ án rất nhiều đương sự):** Tổng rộng **`17.50 cm`** (`9923 dxa`), `tblInd = -1.50 cm` (`-851 dxa`):
    - Cột trái: **`9.00 cm`** (`5104 dxa`)
    - Cột phải: **`8.50 cm`** (`4819 dxa`)
- **Cột Trái — Bảng con lồng 1×1 có khung viền đen kín (`0.5pt`):**
  - 1 đoạn trống phía trên bảng con.
  - Bảng con 1 cột × 1 hàng, viền đơn đen kín 4 cạnh (`single, sz=4, color=#000000`), rộng `7.05 cm` (`3995 dxa`) *(hoặc `8.49 cm` / `4815 dxa` khi bật chế độ rộng)*, chiều cao tối thiểu `1.79 cm` (`1013 dxa`).
  - Nội dung trong khung: **Times New Roman, cỡ `12pt` (`sz=24`), đứng thường**, căn trái, `firstLine = 0`, `After = 6pt`, gồm đúng **7 dòng cố định + 2 dòng trống cuối khung**:
    ```text
    - TKV: {{HO_TEN_TKV}}
    - {{CHUC_DANH_THAM_PHAN}}: {{HO_TEN_THAM_PHAN}}
    - Ngày báo cáo VT:
    - Ngày báo cáo PCA:
    - Ngày báo cáo Tổ TP:
    - Ngày trình CA:
    - Lý do trình:
    [dòng trống]
    [dòng trống]
    ```
  - *(Chức danh Thẩm phán linh hoạt: `Thẩm phán TANDTC` hoặc `Thẩm phán TAND bậc 3`).*
- **Cột Phải — Danh Sách Đương Sự (nằm song song bên phải đường kẻ dọc):**
  - Toàn bộ cỡ **`14pt`**, căn đều 2 bên (`justify`), `firstLine = 0`, `After = 6pt`:
  - 1 đoạn trống đầu ô (`14pt`).
  - `Có các đương sự là:` (`14pt`, đứng thường).
  - `- ` (`14pt`, thường) + **`Nguyên đơn:`** (`14pt`, **IN ĐẬM**) + ` Họ tên, sinh năm ...;` (`14pt`, thường) / `Địa chỉ: ...` (`14pt`, thường).
  - `- ` (`14pt`, thường) + **`Bị đơn:`** (`14pt`, **IN ĐẬM**, `Before = 12pt`) + ` Họ tên, sinh năm ...;` / `Địa chỉ: ...`.
  - *(Khối động — chỉ hiện khi có)*: `- ` (`14pt`, thường) + **`Người có quyền lợi, nghĩa vụ liên quan:`** (`14pt`, **IN ĐẬM**, `Before = 12pt`) $\rightarrow$ đánh số `1.`, `2.`, ... (`14pt`, thường); nếu có người đại diện theo ủy quyền thì in nghiêng nhãn `- ` + *`Người đại diện theo ủy quyền của ...`* (`14pt`, **IN NGHIÊNG**) + `: Ông/Bà ...` (`14pt`, thường).
  - *`- Người đề nghị giám đốc thẩm`* *(hoặc `tái thẩm`)* (`14pt`, **IN NGHIÊNG**, `Before = 12pt`) + `: {{NGUOI_DE_NGHI}}` (`14pt`, thường).

---

### 📝 Thành phần 4: Mục `I. TÓM TẮT NỘI DUNG VỤ ÁN`
- **Tiêu đề mục:** **`I. TÓM TẮT NỘI DUNG VỤ ÁN`** — `14pt` *(hoặc `13pt` qua tham số `roman_heading_size_pt`)*, **IN ĐẬM, IN HOA**, thụt đầu dòng `1.25 cm`, `Before = 12pt`, `After = 6pt`.
- **Yêu cầu nội dung bắt buộc (100% chuẩn xác theo hồ sơ, cấm bịa đặt):**
  - BẮT BUỘC ghi rõ đầy đủ 3 nhóm yêu cầu tố tụng:
    1. *Yêu cầu khởi kiện của nguyên đơn:* Ghi rõ nội dung yêu cầu, xác định rõ có thay đổi, bổ sung, rút yêu cầu trong quá trình giải quyết vụ án (tại sơ thẩm, phúc thẩm) hay không.
    2. *Yêu cầu phản tố của bị đơn:* Xác định có yêu cầu phản tố theo Điều 200, 202 BLTTDS hay không; nếu không có phải khẳng định rõ bị đơn không có yêu cầu phản tố (chỉ có ý kiến không đồng ý hoặc nội dung kháng cáo).
    3. *Yêu cầu độc lập của người có quyền lợi, nghĩa vụ liên quan:* Xác định có yêu cầu độc lập theo Điều 201, 202 BLTTDS hay không; nếu không có phải khẳng định rõ và nêu ý kiến/nguyện vọng của họ cùng hướng giải quyết của Tòa án.
- **Trình bày linh hoạt 2 chế độ:**
  - *Dạng phân mục (nhiều đương sự / có yêu cầu độc lập):* Tiêu đề từng nhóm đương sự (`1. Theo Đơn khởi kiện...`, `2. Bị đơn là...`, `3. Người có quyền lợi, nghĩa vụ liên quan...`) định dạng **`14pt`, IN NGHIÊNG**, thụt `1.25 cm`.
  - *Dạng tường thuật liền mạch (vụ án đơn giản 2 bên):* Viết thẳng các đoạn văn diễn biến giao dịch & lời trình bày 2 bên nhưng vẫn phải bảo đảm đầy đủ 3 nhóm yêu cầu tố tụng nêu trên.
- **Đoạn văn nội dung:** **`14pt`, đứng thường**, căn đều, thụt `1.25 cm`, `After = 6pt`, luôn kèm trích dẫn số bút lục `(BL ...)`.

---

### ⚖️ Thành phần 5: Mục `II. QUÁ TRÌNH GIẢI QUYẾT`
- **Tiêu đề mục:** **`II. QUÁ TRÌNH GIẢI QUYẾT`** — **IN ĐẬM, IN HOA**, thụt `1.25 cm`, `Before = 6pt`, `After = 6pt`.
- **1. Bản án/Quyết định sơ thẩm:**
  - Dòng dẫn: **`1. Tại Bản án {{LOAI_BA_ST}} số {{SO_BA_ST}} ngày {{NGAY_BA_ST}}, {{TOA_AN_ST}} quyết định (tóm tắt):`** — **`14pt`, IN ĐẬM**, căn đều, thụt `1.25 cm`.
  - Nội dung tóm tắt quyết định sơ thẩm (`1.`, `2.`, `2.1.`...): **`14pt`, IN NGHIÊNG**, căn đều, thụt `1.25 cm`.
  - Câu chuẩn cố định: `Toà án cấp sơ thẩm còn quyết định về [chi phí tố tụng,] án phí và quyền kháng cáo của các đương sự.` — `14pt`, đứng thường.
  - Thông tin kháng cáo / kháng nghị phúc thẩm — `14pt`, đứng thường.
- **2. Bản án/Quyết định phúc thẩm & Tiền sử giải quyết đơn:**
  - Dòng dẫn: **`2. Tại Bản án {{LOAI_BA_PT}} số {{SO_BA_PT}} ngày {{NGAY_BA_PT}}, {{TOA_AN_PT}} quyết định: ...`** — **`14pt`, IN ĐẬM**, căn đều, thụt `1.25 cm`.
  - Tiền sử đơn & Thông báo giải quyết của VKSND cấp cao / VKSNDTC / TAND cấp cao (nếu có) — `14pt`, đứng thường.

---

### 📌 Thành phần 6: Mục `III. ĐỀ NGHỊ CỦA ĐƯƠNG SỰ`
- **Tiêu đề mục:** **`III. ĐỀ NGHỊ CỦA ĐƯƠNG SỰ`** — **IN ĐẬM, IN HOA**, thụt `1.25 cm`, `Before = 6pt`, `After = 6pt`.
- **Đoạn mở đầu:** `Ngày {{NGAY_DON}}, {{NGUOI_DE_NGHI}} có đơn đề nghị xem xét theo thủ tục {{THU_TUC}} đối với Bản án {{LOAI_BA_PT}} với lý do:` — `14pt`, đứng thường, thụt `1.25 cm`.
- **Danh sách lý do khiếu nại/đề nghị:** Gạch đầu dòng `- ...;`, `14pt`, đứng thường, căn đều, thụt `1.25 cm`, `After = 6pt`.

---

### ✍️ Thành phần 7: Mục `IV. ĐỀ XUẤT GIẢI QUYẾT ĐƠN ĐỀ NGHỊ`, Bảng Chữ Ký & Các Cấp Phê Duyệt
- **Tiêu đề mục:** **`IV. ĐỀ XUẤT GIẢI QUYẾT ĐƠN ĐỀ NGHỊ`** — **IN ĐẬM, IN HOA**, thụt `1.25 cm`, `Before = 6pt`, `After = 6pt`.
- **`1. Ý kiến đề xuất của cán bộ nghiên cứu:`** — **`14pt`, IN ĐẬM**, thụt `1.25 cm`, `After = 6pt`.
  - Nội dung phân tích pháp lý: **`14pt`, đứng thường**, căn đều, thụt `1.25 cm`, `After = 6pt`, trích dẫn đầy đủ `(BL ...)` (nếu đã rút hồ sơ) và điều luật đúng phiên bản đã khóa ở Bước 1. **Trình bày rõ 2 lớp nhận định:** *(a) Đánh giá toàn diện việc giải quyết vụ án của Tòa án hai cấp (Tố tụng – Chứng cứ – Nội dung – Số học/THA, không phụ thuộc lý do trong đơn)* và *(b) Đánh giá, đối chiếu cụ thể từng lý do đương sự nêu trong đơn*.
    - Nếu đánh số `1)`, `2)`, `3)`...: số **`1)` IN ĐẬM**, nội dung đứng thường.
    - Nếu tách tiểu ý `- Đối với ...:`: cụm dẫn *`Đối với ...`* **IN NGHIÊNG**, nội dung đứng thường.
  - Đoạn chốt đề xuất: **`Thư ký viên đề xuất`** (**`14pt`, IN ĐẬM**) + `: {{DE_XUAT_CHOT}}` (`14pt`, đứng thường), `After = 0pt`.
  - *(Tùy chọn cho hướng Trả lời đơn)*: `Kính trình Thẩm phán Tòa án nhân dân tối cao xem xét, quyết định.` (`14pt`, đứng thường, `After = 0pt`).
  - **Bảng 3 — Bảng Chữ Ký Thư Ký Viên *(BẮT BUỘC cho cả 2 hướng Kháng nghị & Trả lời đơn)*:**
    - Bảng 2 cột × 1 hàng, **KHÔNG VIỀN (`none`)**, tổng rộng `15.98 cm` (`9060 dxa`), mỗi cột **`7.99 cm`** (`4530 dxa`), `vAlign = center`.
    - Cột trái: Để trống.
    - Cột phải (căn giữa, `firstLine = 0`, `14pt`, **IN ĐẬM**):
      - Dòng 1: **`Thư ký viên`**
      - 3 dòng trống ký tên
      - Dòng 5: **`{{HO_TEN_TKV}}`**
- **`2. Ý kiến của {{CHUC_DANH_THAM_PHAN}} {{HO_TEN_THAM_PHAN}} ngày      /     /{{NAM}}:`** — **`14pt`, IN ĐẬM**, thụt `1.25 cm`, `Before = 12pt`, `After = 6pt`.
  - Nội dung: Điền sẵn **Dự thảo Đề xuất Ý kiến của Thẩm phán** (do AI soạn ở Bước 5 và User duyệt ở Bước 5.5 — `14pt`, đứng thường, căn đều, thụt `1.25 cm`) HOẶC để 3 dòng trống nếu User chọn để Thẩm phán tự ghi tay.
- **2 Chế độ phân cấp duyệt cuối Tờ trình (theo TB 383 & TB 679):**
  - **Chế độ Rút gọn (`che_do_rut_gon: true` — Hướng Trả lời đơn):** Tự động **lược bỏ các mục 3, 4, 5**, dừng ngay tại Mục 2 để Thẩm phán duyệt ký trực tiếp Thông báo giải quyết đơn.
  - **Chế độ Toàn diện (`che_do_rut_gon: false` — Mặc định cho Hướng Kháng nghị / Vụ phức tạp họp Tổ 05 Thẩm phán):** Tự động in đầy đủ các mục Ý kiến Lãnh đạo cấp trên:
    - **`3. Ý kiến của Lãnh đạo Vụ (Vụ trưởng / Phó Vụ trưởng) ngày      /     /{{NAM}}:`** — **`14pt`, IN ĐẬM**, thụt `1.25 cm`, `Before = 12pt`, `After = 6pt` *(+ 3 dòng trống)*.
    - **`4. Ý kiến của Phó Chánh án Tòa án nhân dân tối cao ngày      /     /{{NAM}}:`** — **`14pt`, IN ĐẬM**, thụt `1.25 cm`, `Before = 12pt`, `After = 6pt` *(+ 3 dòng trống)*.
    - **`5. Ý kiến của Chánh án Tòa án nhân dân tối cao ngày      /     /{{NAM}}:`** — **`14pt`, IN ĐẬM**, thụt `1.25 cm`, `Before = 12pt`, `After = 6pt` *(+ 3 dòng trống)*.

---

## 3. Chuẩn Cấu Trúc Dữ Liệu JSON Đầu Vào (Step 5.5 $\rightarrow$ Step 6 Schema)

Sau khi User xác nhận **`"OK"`** và chọn **Option 1** hoặc **Option 3** ở Bước 5.5/6, AI xuất file JSON theo đúng cấu trúc dưới đây để truyền vào `export_to_trinh.py`:

```json
{
  "metadata": {
    "co_quan_cap_tren": "TOÀ ÁN NHÂN DÂN TỐI CAO",
    "don_vi_trinh": "VỤ GIÁM ĐỐC KIỂM TRA II",
    "dia_danh": "Hà Nội",
    "ngay_trinh": "05",
    "thang_trinh": "11",
    "nam_trinh": "2026",
    "tieu_de_van_ban": "TỜ TRÌNH",
    "nguoi_nhan_trinh": "CHÁNH ÁN TÒA ÁN NHÂN DÂN TỐI CAO",
    "loai_vu_an": "dân sự",
    "quan_he_tranh_chap": "Tranh chấp ...",
    "loai_ban_an_pt": "dân sự phúc thẩm",
    "so_ba_pt": ".../2024/DS-PT",
    "ngay_ba_pt": ".../.../2024",
    "toa_an_pt": "Tòa án nhân dân tỉnh ...",
    "ho_ten_tkv": "...",
    "chuc_danh_ky_tkv": "Thư ký viên",
    "chuc_danh_tham_phan": "Thẩm phán TANDTC",
    "ho_ten_tham_phan": "...",
    "ngay_bao_cao_vt": "",
    "ngay_bao_cao_pca": "",
    "ngay_bao_cao_to_tp": "",
    "ngay_trinh_ca": "",
    "ly_do_trinh": "",
    "wide_party_table": false,
    "che_do_rut_gon": false,
    "roman_heading_size_pt": 14
  },
  "duong_su": {
    "nguyen_don": [{ "ho_ten": "Ông/Bà ...", "nam_sinh": "1985", "dia_chi": "..." }],
    "bi_don": [{ "ho_ten": "Ông/Bà ...", "nam_sinh": "1986", "dia_chi": "..." }],
    "nguoi_lien_quan": [],
    "nhan_nguoi_de_nghi": "Người đề nghị giám đốc thẩm",
    "nguoi_de_nghi": "Ông/Bà ..."
  },
  "muc_1_tom_tat": [
    { "tieu_de": "1. Theo Đơn khởi kiện...", "doan_van": ["... (BL ...)."] }
  ],
  "muc_2_qua_trinh": {
    "so_tham_tieu_de": "1. Tại Bản án dân sự sơ thẩm số ... ngày ..., Tòa án nhân dân ... quyết định (tóm tắt):",
    "so_tham_quyet_dinh": ["1. ..."],
    "so_tham_ghi_chu": ["Toà án cấp sơ thẩm còn quyết định về án phí và quyền kháng cáo của các đương sự."],
    "phuc_tham_tieu_de": "2. Tại Bản án dân sự phúc thẩm số ... ngày ..., Tòa án nhân dân ... quyết định: ...",
    "phuc_tham_quyet_dinh": [],
    "tien_su_sau_phuc_tham": ["Tại Thông báo số ... ngày ..., Viện kiểm sát nhân dân ..."]
  },
  "muc_3_de_nghi": {
    "mo_dau": "Ngày ..., ông/bà ... có đơn đề nghị xem xét theo thủ tục giám đốc thẩm đối với Bản án dân sự phúc thẩm với lý do:",
    "cac_ly_do": ["- ...;"]
  },
  "muc_4_de_xuat": {
    "noi_dung_phan_tich": ["1) ... (BL ...)."],
    "de_xuat_chot": "Kháng nghị giám đốc thẩm theo hướng ... / Trả lời đơn đề nghị giám đốc thẩm của ...",
    "kinh_trinh": "",
    "du_thao_y_kien_tham_phan": ["(Dự thảo ý kiến Thẩm phán — hoặc để mảng rỗng [] nếu để trống ghi tay)"]
  }
}
```

---

## 4. Quy Trình Chạy Script Xuất File & Kiểm Soát Chất Lượng (Step 6 Quality Gate)

Luôn chạy 2 bước tuần tự (**Dry-run kiểm tra trước $\rightarrow$ Xuất file chính thức sau**) bằng Python trong `.venv`:

```powershell
# Bước 1: Dry-run kiểm tra lỗi placeholder & tính nhất quán dữ liệu (chưa ghi file)
python 01_Flow_Xu_Ly_Don\scripts\export_to_trinh.py --input "<path_to_payload.json>" --output "03_Ket_qua_Xu_ly\<Số>_Vu_<ten_vu>\<ten_file>.docx" --dry-run

# Bước 2: Xuất file .docx chính thức (có File-lock check chống ghi đè file đang mở trong Word)
python 01_Flow_Xu_Ly_Don\scripts\export_to_trinh.py --input "<path_to_payload.json>" --output "03_Ket_qua_Xu_ly\<Số>_Vu_<ten_vu>\<ten_file>.docx"
```

**4 chốt kiểm soát tự động bên trong [`export_to_trinh.py`](scripts/export_to_trinh.py):**
1. **Bảo toàn VML (`replace_para_text_preserve_vml`):** Thay chữ ở Bảng 1 nhưng giữ nguyên 100% 2 đường kẻ ngang `<v:line>` / `<w:pict>`.
2. **File-lock Check:** Phát hiện nếu file `.docx` đích đang mở trong Microsoft Word để dừng an toàn.
3. **Placeholder Scan:** Quét 100% đoạn văn và bảng lồng, đảm bảo **0 placeholder** `{{...}}` bị sót.
4. **Cross-consistency Check:** Xác thực số Bản án phúc thẩm, ngày tuyên án phúc thẩm, tên Thư ký viên, tên Thẩm phán và tên Người đề nghị khớp 100% xuyên suốt toàn bộ Tờ trình.
