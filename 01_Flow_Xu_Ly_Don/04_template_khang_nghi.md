# 04. TEMPLATE QUYẾT ĐỊNH KHÁNG NGHỊ GĐT / TÁI THẨM TANDTC (SINGLE SOURCE OF TRUTH)

Tài liệu này là **Single Source of Truth duy nhất** tổng hợp toàn bộ cấu trúc, thông số định dạng XML gốc (`word/document.xml`, `word/styles.xml`, `word/footer1.xml`, `word/footer2.xml`), phông chữ, cỡ chữ, căn lề, kích thước bảng (`cm` / `dxa`), 2 đường kẻ ngang VML, cấu trúc JSON đầu vào và quy trình xuất file Word `.docx` chuẩn 100% theo **Bản phát hành** cho **Quyết định Kháng nghị Giám đốc thẩm / Tái thẩm của Chánh án Tòa án nhân dân tối cao**.

> 🔴 **QUY TẮC BẮT BUỘC (MANDATORY ENGINE):**
> - Khi xuất Quyết định Kháng nghị `.docx` ở Bước 6 (lưu vào [`03_Ket_qua_Xu_ly/`](../03_Ket_qua_Xu_ly/) theo thư mục con từng vụ án `<Số>_Vu_<ten_vu>/`), hệ thống hỗ trợ 2 chế độ:
>   1. **Chế độ Phôi Mẫu Trực Quan (Khuyên dùng):** Sử dụng [`universal_visual_filler.py`](scripts/universal_visual_filler.py) rót dữ liệu vào [`demo_phoi_khang_nghi_truc_quan.docx`](templates/demo_phoi_khang_nghi_truc_quan.docx). Thẩm tra viên có thể tự do mở file Word này lên chỉnh sửa kiểu dáng, thêm bớt mục bằng chuột trực tiếp.
>   2. **Chế độ Exporter Tự Động:** Sử dụng script [`export_khang_nghi.py`](scripts/export_khang_nghi.py) dựa trên phôi XML gốc [`base_khang_nghi_tandtc.docx`](templates/base_khang_nghi_tandtc.docx).
> - **TUYỆT ĐỐI NGHIÊM CẤM** dựng chay từ file Word trống hoặc dùng PowerShell COM scripting.

---

## 1. Thiết Lập Trang In, Phông Chữ Mặc Định & Chân Trang (Page Setup & Typography)

| Thuộc tính | Thông số chuẩn 100% (trích xuất từ XML Bản phát hành gốc) |
| :--- | :--- |
| **Khổ giấy** | A4 dọc — Rộng `21.0 cm` (`11906 dxa`) × Cao `29.7 cm` (`16838 dxa`) |
| **Lề trang (Margins)** | Trên `2.0 cm` (`1134 dxa`) · Dưới `2.0 cm` (`1134 dxa`) · Trái `3.0 cm` (`1701 dxa`) · Phải `2.0 cm` (`1134 dxa`) |
| **Header / Footer distance** | Header `1.27 cm` (`720 dxa`) · Footer `1.27 cm` (`720 dxa`) · Gutter `0 cm` |
| **Phông chữ toàn văn bản** | **Times New Roman** (đồng bộ cả 4 thuộc tính `w:ascii`, `w:hAnsi`, `w:cs`, `w:eastAsia`) |
| **Đoạn văn nội dung chính** | Cỡ **`14pt`**, căn đều 2 bên (`justify`), thụt đầu dòng **`1.0 cm`** (`567 dxa`), giãn dòng **`1.15`** (`line = 276`), `Space Before = 5pt`, `Space After = 0pt` |
| **Chân trang (Footer)** | Field `PAGE` tại `word/footer1.xml` (trang chẵn) & `word/footer2.xml` (mặc định), `Different first page = False`, bảo toàn nguyên bản từ phôi XML gốc |

---

## 2. Quy Định Chi Tiết 7 Thành Phần & Định Dạng Từng Dòng

### 🏛️ Thành phần 1: Bảng Quốc Hiệu – Số Ký Hiệu – Tiêu Ngữ – Ngày Tháng *(Bảng 1)*
- **Cấu trúc bảng:** 2 cột × 1 hàng, **KHÔNG VIỀN**, tổng rộng `100%` (`5000 pct`).
  - **Cột trái:** `40%` (`6.40 cm` / `3629 dxa`).
  - **Cột phải:** `60%` (`9.60 cm` / `5443 dxa`).
- **Cột trái** *(căn giữa, giãn dòng `1.0`, `firstLine = 0`)*:
  - Đoạn 1: **`TÒA ÁN NHÂN DÂN TỐI CAO`** — **`12pt`, IN ĐẬM, IN HOA**, kèm **đường kẻ ngang VML shape (`margin-left: 45.2pt; margin-top: 17.7pt; width: 90.1pt`)** bảo toàn từ phôi XML gốc.
  - Đoạn 2: `\nSố: {{SO_QUYET_DINH}}` *(VD: `Số: 28/2026/KN-DS`)* — **`14pt`, đứng thường**.
- **Cột phải** *(căn giữa, giãn dòng `1.0`, `firstLine = 0`)*:
  - Đoạn 1: **`CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM`** (**`12pt`, IN ĐẬM, IN HOA**) + xuống dòng mềm `\n` + **`Độc lập - Tự do - Hạnh phúc`** (**`14pt`, IN ĐẬM**).
  - Đoạn 2: *`Hà Nội, ngày {{NGAY}} tháng {{THANG}} năm {{NAM}}`* — **`14pt`, IN NGHIÊNG**, `Before = 6pt`, kèm **đường kẻ ngang VML `Line 5`** (`from="71.95pt,1.6pt" to="197.95pt,1.6pt"`) bảo toàn từ phôi XML gốc.

---

### 📄 Thành phần 2: Khối Tên Quyết Định, Đối Tượng Kháng Nghị & Thẩm Quyền *(Căn giữa, `firstLine = 0`)*
*(Sau Bảng 1 có 2 dòng trống đệm `line = 1.0`, `14pt`)*
1. **`QUYẾT ĐỊNH`** — **`14pt`, IN ĐẬM, IN HOA**, căn giữa, `line = 1.0`.
2. **`KHÁNG NGHỊ {{LOAI_THU_TUC_HOA}}`** *(mặc định: `KHÁNG NGHỊ GIÁM ĐỐC THẨM` hoặc `KHÁNG NGHỊ TÁI THẨM`)* — **`14pt`, IN ĐẬM, IN HOA**, căn giữa, `line = 1.0`.
3. **`Đối với Bản án {{LOAI_BAN_AN_PT}} số {{SO_BA_PT}} ngày {{NGAY_BA_PT}}`** — **`14pt`, IN ĐẬM**, căn giữa, `Before = 6pt`, `After = 0pt`.
4. **`của {{TOA_AN_PT}}`** — **`14pt`, IN ĐẬM**, căn giữa, `Before = 0pt`, `After = 5pt`.
5. **`CHÁNH ÁN`** — **`14pt`, IN ĐẬM, IN HOA**, căn giữa, `Before = 18pt`, `line = 1.0`.
6. **`TÒA ÁN NHÂN DÂN TỐI CAO`** — **`14pt`, IN ĐẬM, IN HOA**, căn giữa, `After = 12pt`, `line = 1.0`.

---

### ⚖️ Thành phần 3: Căn Cứ Pháp Lý & Danh Sách Đương Sự *(Toàn khổ trang, `14pt`, `firstLine = 1.0 cm`, `line = 1.15`)*
- **Các dòng `Căn cứ ...;`:** `14pt`, đứng thường, căn đều, `firstLine = 1.0 cm`, `Before = 0pt`, `After = 0pt`, `line = 1.15`.
  - VD chuẩn BLTTDS hiện hành: `Căn cứ Điều 326, khoản 1 Điều 331, khoản 2 Điều 332 Bộ luật Tố tụng dân sự năm 2015 (sửa đổi, bổ sung theo Luật số 85/2025/QH15);`
  - Các căn cứ Nghị quyết Quốc hội / luật chuyên ngành khác (nếu có).
- **Câu dẫn hồ sơ vụ án:** `Sau khi nghiên cứu hồ sơ vụ án {{LOAI_VU_AN}} “` (`14pt`, thường) + *`{{QUAN_HE_TRANH_CHAP}}`* (`14pt`, **IN NGHIÊNG**) + `” giữa các đương sự:` (`14pt`, thường), `Before = 5pt`, `firstLine = 1.0 cm`.
- **Danh sách đương sự** *(nhãn tư cách tố tụng **IN NGHIÊNG** — khác Tờ trình là in đậm!)*:
  - *`- Nguyên đơn:`* (`14pt`, **IN NGHIÊNG**, `Before = 5pt`, `firstLine = 1.0 cm`). Nếu 1 người viết cùng dòng; nếu nhiều người xuống dòng đánh số `1.`, `2.`, ... (`14pt`, thường; cá nhân dùng `cư trú tại:`, cơ quan/tổ chức dùng `Địa chỉ:`, ghi kèm địa chỉ cũ `(trước đây là ...)` nếu có thay đổi địa giới hành chính).
  - *`Người đại diện hợp pháp của ...`* (`14pt`, **IN NGHIÊNG**) + `: Ông/Bà ...` (`14pt`, thường).
  - *`- Bị đơn:`* (`14pt`, **IN NGHIÊNG**) + thông tin bị đơn (`14pt`, thường).
  - *`- Người có quyền lợi, nghĩa vụ liên quan:`* (`14pt`, **IN NGHIÊNG**) $\rightarrow$ đánh số `1.`, `2.`, ... (`14pt`, thường).

---

### 📌 Thành phần 4: Phần `NHẬN THẤY:`
- **Tiêu đề:** **`NHẬN THẤY:`** — **`14pt`, IN ĐẬM, IN HOA**, căn giữa, `Before = 12pt`, `line = 1.15`.
- **Dòng dẫn bản án bị kháng nghị:** `Tại Bản án {{LOAI_BAN_AN_PT}} số {{SO_BA_PT}} ngày {{NGAY_BA_PT}}, {{TOA_AN_PT}} quyết định:` — `14pt`, đứng thường, `firstLine = 1.0 cm`, `Before = 5pt`, `line = 1.15`.
- **Các mục trích quyết định bản án (`1.`, `1.1.`, `2.`, `3.`...):** **`14pt`, IN NGHIÊNG**, căn đều, `firstLine = 1.0 cm`, `Before = 5pt`, `line = 1.15`.
- **Câu chuẩn tiếp theo:** `Ngoài ra, Tòa án cấp phúc thẩm còn quyết định về án phí, chi phí tố tụng và thi hành án theo quy định của pháp luật.` — `14pt`, đứng thường.
- **Diễn biến gửi đơn & Thông báo giải quyết trước đó (VKSND cấp cao / TAND cấp cao...):** `14pt`, đứng thường, căn đều, `firstLine = 1.0 cm`, `Before = 5pt`, `line = 1.15`.

---

### 🔍 Thành phần 5: Phần `XÉT THẤY:`
- **Tiêu đề:** **`XÉT THẤY:`** — **`14pt`, IN ĐẬM, IN HOA**, căn giữa, `Before = 12pt`, `After = 12pt`, `line = 1.0`.
- **Các đoạn lập luận pháp lý:** Đánh số đoạn theo chuẩn Bản án/Quyết định TANDTC **`[1].`**, **`[2].`**, **`[3].`**, **`[4].`**... — `14pt`, đứng thường, căn đều, `firstLine = 1.0 cm`, `Before = 5pt`, `line = 1.15`.
- **Câu chuyển kết:** `\tVì các lẽ trên;` — `14pt`, đứng thường, `firstLine = 1.0 cm`, `Before = 5pt`, `line = 1.15`.

---

### ⚡ Thành phần 6: Phần `QUYẾT ĐỊNH:`
- **Tiêu đề:** **`QUYẾT ĐỊNH:`** — **`14pt`, IN ĐẬM, IN HOA**, căn giữa, `Before = 12pt`, `After = 12pt`, `line = 1.0`.
- **Điều 1:** `1. Kháng nghị Bản án {{LOAI_BAN_AN_PT}} số {{SO_BA_PT}} ngày {{NGAY_BA_PT}} của {{TOA_AN_PT}} về vụ án ...` — `14pt`, đứng thường, `firstLine = 1.0 cm`, `After = 5pt`, `line = 1.15`.
- **Điều 2:** `2. Đề nghị Hội đồng Thẩm phán Tòa án nhân dân tối cao xét xử {{THU_TUC}}, hủy ...` — `14pt`, đứng thường, `firstLine = 1.0 cm`, `After = 5pt`, `line = 1.15`.
- **Điều 3 *(nếu có tạm đình chỉ thi hành án)*:** `3. Tạm đình chỉ thi hành Bản án {{LOAI_BAN_AN_PT}} số {{SO_BA_PT}} ngày {{NGAY_BA_PT}} của {{TOA_AN_PT}} cho đến khi có Quyết định {{THU_TUC}}.` — `14pt`, đứng thường, `firstLine = 1.0 cm`, `After = 12pt`, `line = 1.15`.

---

### ✍️ Thành phần 7: Bảng Nơi Nhận & Chữ Ký Người Ký Kháng Nghị *(Bảng 2)*
- **Cấu trúc bảng:** 2 cột × 1 hàng, **KHÔNG VIỀN (`none`)**.
  - **Cột trái (`Nơi nhận`):** `9.74 cm` (`5524 dxa`).
  - **Cột phải (`Chữ ký`):** `6.24 cm` (`3538 dxa`).
- **Cột trái (`Nơi nhận`, giãn dòng `1.0`, `firstLine = 0`, `After = 0pt`):**
  - Dòng 1: ***`Nơi nhận:`*** — **`11pt`, IN ĐẬM, IN NGHIÊNG**.
  - Các dòng nơi nhận (`11pt`, đứng thường):
    - `- Chánh án TANDTC (để báo cáo);` *(khi Phó Chánh án ký thay)*
    - `- Vụ 9 - VKSNDTC (kèm hồ sơ vụ án);` *(hoặc Vụ VKSNDTC tương ứng)*
    - `- TAND tỉnh/thành phố ...;`
    - `- TAND khu vực / huyện ...;`
    - `- Cơ quan THADS có thẩm quyền;`
    - `- Các đương sự (theo địa chỉ);`
    - `- {{CHUC_DANH_THAM_PHAN}} {{HO_TEN_THAM_PHAN}};`
    - `- Lưu: TK Phó Chánh án, VT, Vụ II - TANDTC (03 bản).`
- **Cột phải (`Chữ ký`, căn giữa, `left = -0.18 cm`, giãn dòng `1.0`):**
  - Trường hợp Phó Chánh án ký thay: **`KT. CHÁNH ÁN`** (`12pt`, IN ĐẬM) / **`PHÓ CHÁNH ÁN`** (`12pt`, IN ĐẬM).
  - Trường hợp Chánh án ký trực tiếp: **`CHÁNH ÁN`** (`12pt`, IN ĐẬM).
  - 6 dòng trống ký tên và đóng dấu (`12pt`, `line = 1.0`).
  - **`{{HO_TEN_NGUOI_KY}}`** — **`14pt`, IN ĐẬM**, `Before = 5pt`, căn giữa.

---

## 3. Chuẩn Cấu Trúc Dữ Liệu JSON Đầu Vào (Step 5.5 $\rightarrow$ Step 6 Schema)

Sau khi User xác nhận **`"OK"`** và chọn **Option 2** hoặc **Option 3** ở Bước 5.5/6, AI xuất file JSON theo đúng cấu trúc dưới đây để truyền vào `export_khang_nghi.py`:

```json
{
  "metadata": {
    "co_quan_ban_hanh": "TÒA ÁN NHÂN DÂN TỐI CAO",
    "so_quyet_dinh": ".../2026/KN-DS",
    "dia_danh": "Hà Nội",
    "ngay_ky": "...",
    "thang_ky": "...",
    "nam_ky": "2026",
    "tieu_de_1": "QUYẾT ĐỊNH",
    "tieu_de_2": "KHÁNG NGHỊ GIÁM ĐỐC THẨM",
    "loai_vu_an": "dân sự",
    "quan_he_tranh_chap": "Tranh chấp ...",
    "loai_ban_an_pt": "dân sự phúc thẩm",
    "so_ba_pt": ".../2024/DS-PT",
    "ngay_ba_pt": ".../.../2024",
    "toa_an_pt": "Tòa án nhân dân tỉnh ...",
    "tham_quyen_dong_1": "CHÁNH ÁN",
    "tham_quyen_dong_2": "TÒA ÁN NHÂN DÂN TỐI CAO",
    "can_cu_phap_ly": [
      "Căn cứ Điều 326, khoản 1 Điều 331, khoản 2 Điều 332 Bộ luật Tố tụng dân sự năm 2015 (sửa đổi, bổ sung theo Luật số 85/2025/QH15);"
    ]
  },
  "duong_su": {
    "nguyen_don": [{ "mo_ta": "Ông/Bà ..., sinh năm ...; cư trú tại: ..." }],
    "dai_dien_nguyen_don": [],
    "bi_don": [{ "mo_ta": "Ông/Bà ..., sinh năm ...; cư trú tại: ..." }],
    "nguoi_lien_quan": []
  },
  "nhan_thay": {
    "mo_dau_ban_an": "Tại Bản án dân sự phúc thẩm số ... ngày ..., Tòa án nhân dân ... quyết định:",
    "quyet_dinh_ban_an": ["1. ..."],
    "dien_bien_sau_phuc_tham": [
      "Ngoài ra, Tòa án cấp phúc thẩm còn quyết định về án phí, chi phí tố tụng và thi hành án theo quy định của pháp luật."
    ]
  },
  "xet_thay": [
    "[1]. ...",
    "[2]. ..."
  ],
  "quyet_dinh": [
    "1. Kháng nghị Bản án dân sự phúc thẩm số ... ngày ... của Tòa án nhân dân ...",
    "2. Đề nghị Hội đồng Thẩm phán Tòa án nhân dân tối cao xét xử giám đốc thẩm, hủy ...",
    "3. Tạm đình chỉ thi hành Bản án dân sự phúc thẩm số ... ngày ... của Tòa án nhân dân ... cho đến khi có Quyết định giám đốc thẩm."
  ],
  "noi_nhan": [
    "- Chánh án TANDTC (để báo cáo);",
    "- Vụ 9 - VKSNDTC (kèm hồ sơ vụ án);",
    "- Các đương sự (theo địa chỉ);",
    "- Lưu: TK Phó Chánh án, VT, Vụ II - TANDTC (03 bản)."
  ],
  "ky_duyet": {
    "chuc_danh_lines": ["KT. CHÁNH ÁN", "PHÓ CHÁNH ÁN"],
    "ho_ten_nguoi_ky": "..."
  }
}
```

---

## 4. Quy Trình Chạy Script Xuất File & Kiểm Soát Chất Lượng (Step 6 Quality Gate)

Luôn chạy 2 bước tuần tự (**Dry-run kiểm tra trước $\rightarrow$ Xuất file chính thức sau**) bằng Python trong `.venv`:

```powershell
# Bước 1: Dry-run kiểm tra placeholder & tính nhất quán dữ liệu (chưa ghi file)
python 01_Flow_Xu_Ly_Don\scripts\export_khang_nghi.py --input "<path_to_payload.json>" --output "03_Ket_qua_Xu_ly\<Số>_Vu_<ten_vu>\<ten_file>.docx" --dry-run

# Bước 2: Xuất file .docx chính thức (có File-lock check chống ghi đè file đang mở)
python 01_Flow_Xu_Ly_Don\scripts\export_khang_nghi.py --input "<path_to_payload.json>" --output "03_Ket_qua_Xu_ly\<Số>_Vu_<ten_vu>\<ten_file>.docx"
```

**4 chốt kiểm soát tự động bên trong [`export_khang_nghi.py`](scripts/export_khang_nghi.py):**
1. **Bảo toàn 2 đường kẻ VML (`replace_para_runs_preserve_vml`):** Thay chữ ở Bảng 1 nhưng giữ nguyên 100% VML shape dưới tên cơ quan và `Line 5` dưới Tiêu ngữ.
2. **File-lock Check:** Phát hiện file đích đang mở trong Microsoft Word để dừng an toàn.
3. **Placeholder Scan:** Đảm bảo 0 placeholder `{{...}}` bị sót trong toàn bộ văn bản và bảng.
4. **Cross-consistency Check:** Xác thực số Bản án phúc thẩm, ngày tuyên án phúc thẩm, tên Tòa án phúc thẩm và tên người ký thống nhất 100% giữa Trích yếu – `NHẬN THẤY` – `QUYẾT ĐỊNH` – `Nơi nhận / Chữ ký`.
