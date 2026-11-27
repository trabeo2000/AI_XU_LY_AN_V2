# 06. TEMPLATE VĂN BẢN TRUNG GIAN & EARLY EXIT (OPTION 7 — SINGLE SOURCE OF TRUTH)

Tài liệu quy định chuẩn thể thức và cấu trúc dữ liệu cho **4 loại Văn bản trung gian / `Early Exit`** xuất trực tiếp qua [`export_dossier.py`](scripts/export_dossier.py) với cờ `--option 7`.

## 1. 4 Loại Văn Bản Được Hỗ Trợ (`van_ban_trung_gian.loai_van_ban`)

| Mã `loai_van_ban` | Tiêu đề mặc định | Trường hợp áp dụng trong quy trình |
| :--- | :--- | :--- |
| `phieu_chuyen_don` | **PHIẾU CHUYỂN ĐƠN** | Bản án cấp huyện/khu vực thuộc thẩm quyền TAND cấp tỉnh (Điều 4 NQ 01/2025/NQ-HĐTP) — `Early Exit` Bước 2. |
| `de_xuat_rut_ho_so` | **PHIẾU ĐỀ XUẤT RÚT HỒ SƠ VỤ ÁN** | Chế độ hồ sơ *Chưa rút HS* (chỉ có Đơn + Bản án) nhưng phát hiện dấu hiệu vi phạm cần rút hồ sơ gốc (QC 75). |
| `yeu_cau_bo_sung` | **THÔNG BÁO** *(yêu cầu bổ sung tài liệu)* | Đơn thiếu tài liệu tố tụng bắt buộc (bản án ST/PT, văn bản ủy quyền, chứng cứ tình tiết mới) — `Early Exit` Bước 2. |
| `bao_cao_co_quan_chuyen_don` | **CÔNG VĂN BÁO CÁO** | Trả lời cơ quan chuyển đơn (ĐBQH, Ủy ban của Quốc hội, Ban Nội chính, MTTQ, Báo chí...) theo đúng hạn trả lời. |

## 2. Thể Thức Chuẩn Hành Chính Tư Pháp (A4 — Times New Roman)
- **Khổ giấy & lề:** A4 dọc (`21.0 × 29.7 cm`), lề Trên `2.0 cm` · Dưới `2.0 cm` · Trái `3.0 cm` · Phải `2.0 cm`.
- **Phông chữ:** Times New Roman đồng bộ 4 thuộc tính (`ascii`, `hAnsi`, `cs`, `eastAsia`), thân văn bản **`14pt`**, căn đều (`justify`), thụt đầu dòng `1.06 cm`, giãn dòng `1.15`.
- **Cấu trúc:** Bảng Quốc hiệu – Tiêu ngữ – Số văn bản – Ngày tháng *(không viền)* → Tiêu đề & Trích yếu → `Kính gửi:` → Thân văn bản → Bảng `Nơi nhận` & `Chữ ký` *(không viền)*.

## 3. Cấu Trúc Dữ Liệu JSON (`van_ban_trung_gian`)

```json
{
  "van_ban_trung_gian": {
    "loai_van_ban": "phieu_chuyen_don",
    "so_van_ban": ".../PC-TANDTC-Vụ II",
    "tieu_de": "PHIẾU CHUYỂN ĐƠN",
    "trich_yeu": "Về việc chuyển đơn đề nghị giám đốc thẩm, tái thẩm",
    "kinh_gui": ["Tòa án nhân dân tỉnh ..."],
    "noi_dung": [
      "Tòa án nhân dân tối cao nhận được đơn của ... đề nghị xem xét theo thủ tục giám đốc thẩm đối với Bản án số ... ngày ... của Tòa án nhân dân ...",
      "Căn cứ khoản 6, khoản 7 Điều 4 Nghị quyết số 01/2025/NQ-HĐTP ngày 14/02/2025 của Hội đồng Thẩm phán TANDTC, Tòa án nhân dân tối cao chuyển đơn cùng tài liệu kèm theo đến Tòa án nhân dân tỉnh ... để xem xét, giải quyết theo thẩm quyền."
    ],
    "noi_nhan": [
      "- Như trên (kèm đơn và tài liệu);",
      "- Ông/Bà ... (để biết);",
      "- Lưu: VT, Vụ GĐKT II."
    ],
    "chuc_danh_ky": "TL. CHÁNH ÁN\nVỤ TRƯỞNG",
    "nguoi_ky": "..."
  }
}
```
