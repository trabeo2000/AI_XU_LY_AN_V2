# AGENTS.md — QUY ĐỊNH HỆ THỐNG XỬ LÝ HỒ SƠ ĐƠN KHIẾU NẠI TANDTC

Mọi AI Agent trong repository này **BẮT BUỘC TUÂN THỦ 100% QUY TRÌNH 6 BƯỚC (+ CỔNG KHÓA 5.5) CÓ 8 HUMAN CHECKPOINTS** khi nghiên cứu hồ sơ, phân tích đơn GĐT/TT, soạn Tờ trình, dự thảo `.docx` hoặc cập nhật `.xlsx`. CẤM tự ý bỏ bước, gộp bước hay vượt checkpoint *(ngoại lệ duy nhất: `Early Exit` tại Bước 2 sau khi User xác nhận)*.

## 🏛️ Kiến Trúc Modules (100% Relative Path — Portable 3 Máy)

| Phân Vùng | Đường Dẫn | Chức Năng |
| :--- | :--- | :--- |
| **Khung Lõi** | [`AGENTS.md`](AGENTS.md) | Hiến pháp hệ thống: Quy tắc cấm, 6 bước + 5.5, 8 Checkpoints. |
| **Văn Bản Pháp Lý** | [`04_Van_ban_phap_ly/`](../04_Van_ban_phap_ly/) | Luật nội dung/tố tụng, Án lệ (`04_04`), Giải đáp (`04_02`), QĐ GĐT (`04_03`), [NQ 01/2025](../04_Van_ban_phap_ly/NQ_01_2025_NQ_HDTP_Tiep_nhan_nhiem_vu_Tham_quyen_TAND.md). |
| **Nghiệp Vụ & Quy Chế** | [`05_Quy_trinh_va_Pham_vi_Chuc_danh/`](../05_Quy_trinh_va_Pham_vi_Chuc_danh/) | [Cẩm nang NQ 01/2025](../05_Quy_trinh_va_Pham_vi_Chuc_danh/Cam_nang_Nghiep_vu_GDKT_NQ_01_2025_NQ_HDTP.md); [QC 75 / TB 383 / TB 679](../05_Quy_trinh_va_Pham_vi_Chuc_danh/Quy_trinh_Giai_quyet_Don_TANDTC_QC75_TB383_TB679.md). |
| **Biểu Mẫu Chuẩn** | [`01_Flow_Xu_Ly_Don/`](../01_Flow_Xu_Ly_Don/) | [Tờ trình](../01_Flow_Xu_Ly_Don/03_template_to_trinh.md), [Kháng nghị](../01_Flow_Xu_Ly_Don/04_template_khang_nghi.md), [Thông báo](../01_Flow_Xu_Ly_Don/05_template_thong_bao.md), [VB Trung gian](../01_Flow_Xu_Ly_Don/06_template_van_ban_trung_gian.md) & [`templates/`](../01_Flow_Xu_Ly_Don/templates/). |
| **Công Cụ Tech** | [`01_Flow_Xu_Ly_Don/scripts/`](../01_Flow_Xu_Ly_Don/scripts/) | `paths_config.py`, `verify_system.py`, `fast_reader.py`, `export_dossier.py` (Master), `export_to_trinh.py`, `export_khang_nghi.py`, `export_6step_report.py`, `universal_visual_filler.py`. |
| **Hồ Sơ & Đầu Ra** | [`02_Ho_so/`](../02_Ho_so/) → [`03_Ket_qua_Xu_ly/`](../03_Ket_qua_Xu_ly/) | Hồ sơ vào → Đầu ra `03_Ket_qua_Xu_ly/<Số>_Vu_<ten_vu>/` *(tự xếp `01_` mới nhất lên trên)*. |

---

## I. 10 Quy Định Cốt Lõi Bắt Buộc (Mandatory System Rules)

1. **Zero Hallucination & Đủ 3 Nhóm Yêu Cầu Tố Tụng:** Chính xác 100%, không suy diễn ngoài [`02_Ho_so/`](../02_Ho_so/), không bịa số bản án, ngày tuyên, bút lục, điều luật hay án lệ. Điểm chưa rõ gắn `[CẦN KIỂM CHỨNG]`. Mục I Tờ trình/Báo cáo **BẮT BUỘC đủ 3 nhóm**: *(1)* YC khởi kiện nguyên đơn (có đổi/bổ sung/rút không); *(2)* YC phản tố bị đơn (có theo Đ200 BLTTDS hay chỉ là ý kiến/kháng cáo); *(3)* YC độc lập người liên quan (có theo Đ201 BLTTDS hay không, ý kiến và cách Tòa xử lý).
2. **Xác Nhận Luật Trước Khi Đọc Hồ Sơ & Strict Scope:**
   - Tại 1A, liệt kê bằng **văn bản thuần** danh mục [`04_Van_ban_phap_ly/`](../04_Van_ban_phap_ly/) và [`05_Quy_trinh_va_Pham_vi_Chuc_danh/`](../05_Quy_trinh_va_Pham_vi_Chuc_danh/) kèm đề xuất *(CẤM popup `ask_question`, không hỏi lại nguồn hồ sơ)*. DỪNG chờ User xác nhận TRƯỚC KHI mở [`02_Ho_so/`](../02_Ho_so/).
   - Nếu User tag `@file`/`@folder`, CHỈ đọc đúng phạm vi đó; chỉ quét [`02_Ho_so/`](../02_Ho_so/) khi chưa chỉ định (nhiều vụ thì hỏi chọn 01 vụ).
3. **State Persistence (BẮT BUỘC TỪ BƯỚC 2 ĐẾN BƯỚC 6):** Đầu mỗi phản hồi in cố định:
   ```markdown
   > **[TRẠNG THÁI HỒ SƠ ĐÃ KHÓA]**
   > - **Vụ việc / Người gửi đơn:** [Họ tên] | **Mã đơn / Công văn chuyển & Hạn:** [Mã / Cơ quan / Hạn hoặc Không]
   > - **Chế độ HS & Tình trạng THA:** [Đã rút HS gốc (đủ BL) / Chưa rút HS (Đơn+BA)] | [Chưa / Đang / Đã THA xong / Chưa rõ]
   > - **Bản án / QĐ bị khiếu nại:** ST số ... ngày ... (TAND ...); PT số ... ngày ... (TAND ...)
   > - **Thủ tục & Tiền sử đơn:** [GĐT / Tái thẩm / Cả hai] | [Mới / Bổ sung / Trùng lặp / Khiếu nại lại TB cũ]
   > - **Luật đã khóa:** Nội dung [...] | Tố tụng [...]
   > - **Tiến độ:** [Full 6 bước / Early Exit] – Đang thực hiện: Bước [X]
   ```
4. **Early Exit (Ngoại Lệ Duy Nhất Tại Bước 2):** Chỉ đề xuất khi rõ: *(1)* hết thời hiệu; *(2)* không thuộc thẩm quyền TANDTC; *(3)* đã có TB Chánh án TANDTC và không có tình tiết mới hợp pháp; *(4)* thiếu tài liệu bắt buộc chưa thể thụ lý. Sau khi User xác nhận mới đi tắt **Bước 2 → 5.5 → 6**.
5. **Động Cơ Master Python UTF-8 & Zero Junk Files:** Xuất `.docx` vào `03_Ket_qua_Xu_ly/<Số>_Vu_<ten_vu>/` qua duy nhất [`export_dossier.py`](../01_Flow_Xu_Ly_Don/scripts/export_dossier.py) *(hoặc [`universal_visual_filler.py`](../01_Flow_Xu_Ly_Don/scripts/universal_visual_filler.py))*. CẤM viết script `.py` mới, CẤM PowerShell COM. JSON tạm lưu `01_Flow_Xu_Ly_Don/scripts/.temp/` và tự xóa sau khi xuất.
6. **Single Source of Truth:** Thể thức tuân thủ 100% các file mẫu trong [`01_Flow_Xu_Ly_Don/`](../01_Flow_Xu_Ly_Don/) và phôi tại [`templates/`](../01_Flow_Xu_Ly_Don/templates/).
7. **Đọc Hồ Sơ Siêu Tốc (`fast_reader.py`):** `.docx`/`.xlsx`/PDF text gọi `python 01_Flow_Xu_Ly_Don/scripts/fast_reader.py --input "<file>"`. PDF Scan 50–200 trang: *(1)* `--pages "1-3"` định vị mục lục; *(2)* render vùng trọng tâm `--stitch 2` (DPI 135); *(3)* **Parallel Batch Read** 1 turn, CẤM đọc từng trang lẻ.
8. **100% Relative Path & Verify:** CẤM hardcode `C:\`, `D:\`, `E:\`; dùng [`paths_config.py`](../01_Flow_Xu_Ly_Don/scripts/paths_config.py). Sau tái cấu trúc, bắt buộc chạy `python 01_Flow_Xu_Ly_Don/scripts/verify_system.py`.
9. **Cấm Quét Thư Mục Ẩn IDE:** KHÔNG tự ý quét/đọc `.vscode/`, `.git/`, `.env*` trừ khi User yêu cầu đích danh.
10. **Zero Ad-Hoc Scripting:** KHÔNG tự tạo `.py` mới hay chạy `python -c` để đọc/xuất file. Chỉ sửa code khi User yêu cầu nâng cấp công cụ.

---

## II. Quy Trình Bắt Buộc 6 Bước (+ Cổng Khóa 5.5)

### 🔹 Bước 1: Khởi Tạo, Xác Nhận Văn Bản Pháp Lý & Khóa Hiệu Lực Luật
1. **Checkpoint 1A (~300 tokens):** Liệt kê gọn `[Tên file] — [Từ khóa < 10 từ]` tại `04_Van_ban_phap_ly/` và `05_Quy_trinh_va_Pham_vi_Chuc_danh/` gắn nhãn 🏛️ `[KHO 04]`. CẤM trích nội dung. **🛑 DỪNG chờ User xác nhận TRƯỚC KHI đọc hồ sơ.**
2. Dùng `fast_reader.py` đọc đúng hồ sơ chỉ định: trích xuất định danh vụ việc, quan hệ tranh chấp, mốc thời gian (giao dịch, khởi kiện, ST/PT, THA).
3. **Checkpoint 1B (~350 tokens):** Quét tìm quy định/án lệ đặc thù kho có thể thiếu, lập danh mục Top 3–5 gắn 🔍 `[AI-TÌM]` (Tên + 1 câu < 20 từ) theo 3 tầng: *Luật chuyên ngành*, *Nghị quyết/Thông tư*, *Án lệ TANDTC*. **🛑 DỪNG chờ User duyệt.**
4. **Checkpoint 1C (~450 tokens):** Lập **Ma trận 4 cột**: *Văn bản/Điều luật* (🏛️/🔍) | *Từ khóa tình tiết* | *Giải trình phù hợp (1–2 câu)* | *Áp dụng/Dự phòng*. Kèm **Khóa luật nội dung** (thời điểm giao dịch + chuyển tiếp) & **tố tụng** (thụ lý, ST/PT, GĐT/TT).
   - **Auto-Trigger 01/7/2025:** Có mốc tố tụng từ **01/7/2025** trở đi → tự kích hoạt [NQ 01/2025](../04_Van_ban_phap_ly/NQ_01_2025_NQ_HDTP_Tiep_nhan_nhiem_vu_Tham_quyen_TAND.md) và [Cẩm nang NQ 01/2025](../05_Quy_trinh_va_Pham_vi_Chuc_danh/Cam_nang_Nghiep_vu_GDKT_NQ_01_2025_NQ_HDTP.md) xuyên suốt Bước 1–5.
- **🛑 DỪNG (Checkpoint 1C):** Chờ User `"xác nhận"` sang Bước 2.

### 🔹 Bước 2: Tiền Sử Đơn, Chế Độ HS, Thẩm Quyền, Tình Trạng THA & Q&A
*(In khối `[TRẠNG THÁI HỒ SƠ ĐÃ KHÓA]`).*
1. **Tiền sử & Chế độ hồ sơ:** Xác định đơn mới/bổ sung/trùng lặp/khiếu nại lại TB cũ; gắn cờ **công văn chuyển đơn** (hạn trả lời, nơi báo cáo). Phân rõ: *(a) Đã rút HS gốc (đủ `BL ...`)* vs. *(b) Chưa rút HS (chỉ Đơn + BA)* — nếu *(b)*, KHÔNG đòi hỏi `(BL ...)`, đánh giá ngay khả năng *Trả lời đơn không cần rút HS* hay *Đề xuất rút HS* (QC 75).
2. **Chủ thể, Thời hiệu & Chốt chặn Thẩm quyền (Đ2, Đ4 NQ 01/2025):**
   - *Thuộc TANDTC:* BA/QĐ có hiệu lực của TAND cấp tỉnh, TANDCC (cũ), Tòa Phúc thẩm TANDTC.
   - *Thuộc TAND cấp tỉnh (UBTP):* BA/QĐ của **TAND cấp huyện / TAND khu vực** *(k6, k7 Đ4)* → **BẮT BUỘC cảnh báo** chọn: *(1)* Chuyển đơn về TAND cấp tỉnh (`Early Exit` → Option `7`), hoặc *(2)* Báo cáo rút lên TANDTC khi *"xét thấy cần thiết"* *(k3 Đ2)*.
3. **Phân tách GĐT vs. Tái thẩm & Tình trạng THA:** Đánh giá GĐT / TT / Cả hai (nếu có tài liệu mới phải kiểm đủ **3 tiêu chí**: *Tính mới*, *Tính hợp pháp*, *Giá trị chứng minh*). Kiểm tra **tình trạng THA & người thứ ba ngay tình**: đã thi hành/bán đấu giá/sang tên chưa, có cần **Tạm đình chỉ THA** khẩn cấp không?
4. **Đánh giá `Early Exit` & Q&A:** Nêu căn cứ `Early Exit` (nếu thuộc Mục I.4); liệt kê tài liệu thiếu, mâu thuẫn Đơn–ST–PT và đặt câu hỏi `Q1, Q2...` kèm giả định.
- **🛑 DỪNG (Checkpoint 2):** Chờ User trả lời Q&A và chốt **Tiếp Bước 3** hay **`Early Exit` sang Bước 5.5**.

### 🔹 Bước 3: Timeline, Vấn Đề Pháp Lý Cốt Lõi & Phân Rã Yêu Cầu Trong Đơn
*(In khối `[TRẠNG THÁI HỒ SƠ ĐÃ KHÓA]`).*
1. Lập Timeline đầy đủ từng mốc ngày tháng (phát sinh → khởi kiện → ST → PT → THA → gửi đơn/giải quyết trước).
2. **Cây vấn đề pháp lý cốt lõi của vụ án (ĐỘC LẬP VỚI ĐƠN):** Liệt kê toàn bộ vấn đề Tòa án phải giải quyết cho đủ 3 nhóm yêu cầu (*Khởi kiện / Phản tố / Độc lập*) trên 4 phương diện (*Tố tụng*, *Chứng cứ*, *Nội dung*, *Quyết định/THA*).
3. Phân rã **100%** yêu cầu trong đơn theo nhóm (*Tố tụng*, *Chứng cứ*, *Nội dung*, *Nghĩa vụ/án phí*, *Tình tiết mới*) và chỉ rõ chứng cứ cốt lõi.
- **🛑 DỪNG (Checkpoint 3):** Chờ User `"xác nhận"` sang Bước 4.

### 🔹 Bước 4: Đánh Giá Toàn Diện Bản Án (Độc Lập Với Đơn) & Ma Trận Đối Chiếu Đơn
*(In khối `[TRẠNG THÁI HỒ SƠ ĐÃ KHÓA]`).*
1. **Phần 4A — Đánh giá toàn diện việc giải quyết vụ án của Tòa án hai cấp (KHÔNG phụ thuộc lý do trong đơn — BẮT BUỘC đủ 4 trụ cột):**
   - **Trụ cột 1 (Tố tụng & Thẩm quyền):** Quan hệ tranh chấp, thẩm quyền, thời hiệu, đủ người tham gia tố tụng (liên quan/đại diện/kế thừa), trình tự ST/PT.
     - *Rà soát NQ 01/2025 (nếu vắt 01/7/2025):* **3 lỗi hủy án:** (1) TAND tỉnh nhận đơn trước 01/7 nhưng sau 01/7 mới thụ lý ST thay vì chuyển TAND khu vực *(k1 Đ5)*; (2) Đất đổi địa giới, huyện chưa mở phiên giao nộp chứng cứ/hòa giải trước 01/7 mà Tòa khu vực vẫn xử *(k2 Đ6)*; (3) Bỏ sót cơ quan kế thừa khi giải thể huyện/sáp nhập xã *(Đ8)*. **Bộ lọc hợp pháp (Đ7 — không bắt lỗi nhầm):** Tên Tòa mở phiên khác tên Tòa tuyên án *(k7)*; gạch sửa biểu mẫu cũ *(k8)*; dùng tiếp kết quả tố tụng cũ *(k2)*; đổi người tiến hành tố tụng do sắp xếp bộ máy *(k6)*.
   - **Trụ cột 2 (Thu thập & Đánh giá Chứng cứ):** Đủ và khách quan chưa (thẩm định tại chỗ, đo vẽ, định giá, giám định, đối chất, xác minh nguồn gốc/hiện trạng)? Có mâu thuẫn chứng cứ chưa làm rõ không?
   - **Trụ cột 3 (Áp dụng Pháp luật Nội dung):** Đánh giá độc lập bản chất giao dịch, hiệu lực hợp đồng, quyền sở hữu/QSDĐ, di sản, lỗi bồi thường, bảo vệ người thứ ba ngay tình đối với **toàn bộ yêu cầu khởi kiện, phản tố, độc lập**.
   - **Trụ cột 4 (Quyết định, Bảng kiểm Số học & Khả năng THA):** Khớp giữa Nhận định và Quyết định; không vượt/thiếu yêu cầu; có thi hành án được không? **Bắt buộc lập Bảng kiểm số học:** án phí *(NQ 326/2016)*, lãi suất *(NQ 01/2019)*, chi phí tố tụng, diện tích đo vẽ vs. bản án tuyên, tỷ lệ/số tiền chia.
2. **Phần 4B — Ma trận đối chiếu 100% yêu cầu trong đơn:** Mã & nội dung đơn | Phán quyết ST/PT | Chứng cứ / `(BL ...)` | Điều luật đã khóa + Án lệ/NQ HĐTP (🏛️/🔍) | Nhận xét sơ bộ.
- **🛑 DỪNG (Checkpoint 4):** Chờ User `"xác nhận"` sang Bước 5.

### 🔹 Bước 5: Đánh Giá Căn Cứ, Phản Biện Ngược & Chốt Phương Án Xử Lý
*(In khối `[TRẠNG THÁI HỒ SƠ ĐÃ KHÓA]`).*
1. **Bảng đánh giá căn cứ & mức độ vi phạm** cho từng vấn đề *(cả 4 trụ cột ngoài đơn & các lý do trong đơn)*: *Vi phạm nghiêm trọng – Kháng nghị GĐT / Tình tiết mới – Kháng nghị TT / Sai sót nhỏ – Không kháng nghị / Không có căn cứ / Cần rút HS – xác minh*.
2. **Phản biện ngược (Kiểm tra độ bền lập luận — BẮT BUỘC):** Tự nêu lập luận phản biện mạnh nhất chống lại phương án dự kiến để thử độ vững trước **Tổ 05 Thẩm phán** *(đề xuất Kháng nghị → lý lẽ bảo vệ bản án PT & kiểm ngưỡng "vi phạm nghiêm trọng" Đ326 BLTTDS; đề xuất Trả lời đơn → rủi ro sai sót tiềm ẩn dễ bị chất vấn)*.
3. **Đề xuất Chính thức (điền Mục IV Tờ trình):**
   - **Mục IV.1 (Ý kiến Thư ký viên):** Gồm rõ 2 lớp *(a) Đánh giá toàn diện bản án 2 cấp (Tố tụng – Chứng cứ – Nội dung)* và *(b) Đối chiếu từng lý do trong đơn*; đủ `(BL ...)` (nếu đã rút HS), điều luật đúng phiên bản, câu chốt rõ:
     - *TB 383 & 679:* **Trả lời đơn** → **Tờ trình rút gọn** (`che_do_rut_gon: true`, bỏ mục 3–5 Lãnh đạo) + dự thảo Thông báo để Thẩm phán (`TANDTC`/`TAND bậc 3`) duyệt ký trực tiếp *(TTV ký nháy, phát hành trong 03 ngày làm việc, gửi 01 bản Vụ trưởng)*. **Kháng nghị** / vụ phức tạp → **Tờ trình toàn diện** (`che_do_rut_gon: false`) + QĐ Kháng nghị trình **Tổ 05 Thẩm phán**.
     - *NQ 01/2025:* Hủy án ST giao **TAND khu vực [tên mới]** (CẤM giao TAND huyện cũ); Chánh án TANDTC kháng nghị bản án huyện/khu vực phải đề xuất giao **UBTP TAND cấp tỉnh** *(k3 Đ2)*.
   - **Mục IV.2 (Dự thảo Ý kiến Thẩm phán):** Soạn sẵn ngắn gọn *(hoặc để trống theo yêu cầu ở Bước 5.5)*.
4. **Phương án trung gian/khẩn cấp:** Tạm đình chỉ THA, Đề xuất rút hồ sơ, Yêu cầu bổ sung tài liệu, Phiếu chuyển đơn, Báo cáo cơ quan chuyển đơn *(xuất qua Option `7`)*.
- **🛑 DỪNG (Checkpoint 5):** Chờ User chốt phương án. **CẤM tự xuất file** — chuyển sang Bước 5.5.

### 🔹 Bước 5.5: Bảng Rà Soát Tiếng Việt & Menu Chọn Văn Bản Cần Xuất
*(In khối `[TRẠNG THÁI HỒ SƠ ĐÃ KHÓA]`).*
- **CẤM in `JSON`, dấu `{}`, tên biến hay thuật ngữ kỹ thuật (`Payload`, `Metadata`...) ra chat.** Trình bày 100% bằng **Bảng Markdown tiếng Việt**.
1. **Bảng Rà Soát 6 nhóm:** *(1)* Thông tin đơn, chế độ HS & nguồn chuyển đơn; *(2)* Danh sách đương sự; *(3)* Bản án ST/PT & tình trạng THA; *(4)* Hướng chốt, chế độ Tờ trình (Rút gọn/Toàn diện) & biện pháp kèm theo; *(5)* Trình duyệt, chức danh Thẩm phán (`TANDTC`/`TAND bậc 3`), người ký, Mục IV.2 (`Có`/`Để trống`), nơi nhận; *(6)* Tên file `.docx`/`.xlsx` dự kiến.
2. Báo cáo ngắn kết quả tự soát nhất quán (ngày ST/PT, họ tên, số bản án, số học).
3. **BẮT BUỘC IN NGUYÊN VĂN KHUNG CHỌN SAU:**
   ```text
   🛑 CHECKPOINT 5.5 — XIN THẨM TRA VIÊN CHỌN LOẠI VĂN BẢN CẦN XUẤT (lưu chung theo thư mục con vụ án trong 03_Ket_qua_Xu_ly/):
   1. Tờ trình Chánh án TANDTC (Tờ trình toàn diện hoặc Tờ trình rút gọn theo TB 383/679)
   2. Quyết định Kháng nghị Giám đốc thẩm / Tái thẩm
   3. ĐỒNG THỜI CẢ HAI (Tờ trình và Quyết định Kháng nghị)
   4. Thông báo giải quyết đơn (Mẫu Thẩm phán TANDTC / Thẩm phán TAND bậc 3 duyệt ký theo TB 383/679)
   5. Báo cáo Toàn diện Nghiên cứu 6 Bước
   6. Xuất TẤT CẢ các văn bản trên (Bộ hồ sơ đầy đủ bao gồm cả 2 phương án + Báo cáo 6 bước + Thông báo)
   7. Văn bản trung gian / Early Exit (Phiếu chuyển đơn / Đề xuất rút hồ sơ / Yêu cầu bổ sung tài liệu / Báo cáo cơ quan chuyển đơn)

   👉 Kính mời Thẩm tra viên phản hồi số lựa chọn (Ví dụ: "1", "2", "3", "4", "5", "6" hoặc "7") để hệ thống xuất đúng văn bản theo yêu cầu!
   ```
- **🛑 DỪNG (Checkpoint 5.5):** Chờ User chọn rõ số `1–7`. Nếu chỉ gõ `"OK"` chưa nêu số, CẤM xuất — hỏi lại ngay.

### 🔹 Bước 6: Xuất File Word (.docx) Qua Động Cơ Master (< 3 Giây, 01 Turn)
*(In khối `[TRẠNG THÁI HỒ SƠ ĐÃ KHÓA]`).*
1. Không đọc lại file demo; CẤM tự viết script Python mới.
2. Thực thi trong **01 turn** đúng 2 thao tác:
   - `write_file` lưu dữ liệu ngầm vào `01_Flow_Xu_Ly_Don/scripts/.temp/payload_<ten_vu>.json`.
   - `run_command` gọi duy nhất:
     ```bash
     python 01_Flow_Xu_Ly_Don/scripts/export_dossier.py --payload "01_Flow_Xu_Ly_Don/scripts/.temp/payload_<ten_vu>.json" --option <1|2|3|4|5|6|7> --case-slug "<ten_slug>"
     ```
3. `export_dossier.py` tự kiểm tra file-lock, xếp `01_Vu_<ten_slug>/`, cập nhật mục lục và xóa file tạm. Báo cáo hoàn tất kèm link trực tiếp tới từng file.

---

## III. Anti-Patterns (TUYỆT ĐỐI NGHIÊM CẤM)

- ❌ Tự viết script `.py` mới, chạy `python -c` để đọc/xuất file, hoặc dùng PowerShell COM ở Bước 1–6.
- ❌ Đọc lại `demo_*payload.json` ở Bước 6 hoặc xả rác file tạm trong `01_Flow_Xu_Ly_Don/scripts/`.
- ❌ In JSON / `{}` / tên biến lập trình ra chat ở Bước 5.5.
- ❌ Đọc mù PDF scan từng trang lẻ qua nhiều turn (phải định vị + `--stitch 2` + Batch Read 1 turn).
- ❌ Gộp bước, vượt Checkpoint, tự ý `Early Exit`, hoặc xuất file khi User chưa chọn rõ số `1–7` ở Checkpoint 5.5.
- ❌ Đọc `02_Ho_so/` trước 1A hoặc đọc lan sang hồ sơ khác ngoài phạm vi chỉ định.
- ❌ Quên khối `[TRẠNG THÁI HỒ SƠ ĐÃ KHÓA]` từ Bước 2 đến Bước 6.
- ❌ Áp sai mốc hiệu lực luật, nhầm GĐT/Tái thẩm, bỏ sót đánh giá toàn diện 4 trụ cột ngoài đơn ở Bước 4 hoặc phản biện ngược ở Bước 5.
- ❌ Ôm việc của TAND cấp tỉnh (bỏ qua cảnh báo thẩm quyền cấp huyện/khu vực) hoặc bắt lỗi nhầm thủ tục chuyển tiếp hợp pháp theo Đ7 NQ 01/2025.
- ❌ Ghi cứng đường dẫn ổ đĩa (`C:\`, `D:\`, `E:\`) hoặc tự ý quét/đọc `.vscode/`, `.git/`, `.env*`.
