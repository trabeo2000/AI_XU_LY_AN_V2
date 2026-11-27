# AGENTS.md — QUY ĐỊNH HỆ THỐNG XỬ LÝ HỒ SƠ ĐƠN KHIẾU NẠI TANDTC

Mọi AI Agent làm việc trong repository xử lý án này **BẮT BUỘC TUÂN THỦ 100% QUY TRÌNH 6 BƯỚC (+ CỔNG KHÓA 5.5) CÓ 8 HUMAN CHECKPOINTS** khi nghiên cứu hồ sơ, phân tích đơn GĐT/Tái thẩm, soạn Tờ trình, dự thảo văn bản (`.docx`) hoặc cập nhật sổ (`.xlsx`). Không được tự ý bỏ bước, gộp bước hay vượt checkpoint *(ngoại lệ duy nhất: `Early Exit` tại Bước 2 sau khi Người dùng xác nhận rõ ràng)*.

## 🏛️ Kiến Trúc Modules (100% Relative Path — Portable 3 Máy Tính)

| Phân Vùng | Đường Dẫn | Chức Năng |
| :--- | :--- | :--- |
| **Khung Lõi** | [`AGENTS.md`](AGENTS.md) | Hiến pháp hệ thống: Quy tắc cấm, 6 bước + 5.5, 8 Checkpoints. |
| **Văn Bản Pháp Lý** | [`04_Van_ban_phap_ly/`](../04_Van_ban_phap_ly/) | Luật nội dung/tố tụng, Án lệ, [NQ 01/2025/NQ-HĐTP](../04_Van_ban_phap_ly/NQ_01_2025_NQ_HDTP_Tiep_nhan_nhiem_vu_Tham_quyen_TAND.md). |
| **Nghiệp Vụ & Quy Chế** | [`05_Quy_trinh_va_Pham_vi_Chuc_danh/`](../05_Quy_trinh_va_Pham_vi_Chuc_danh/) | [Cẩm nang GĐKT NQ 01/2025](../05_Quy_trinh_va_Pham_vi_Chuc_danh/Cam_nang_Nghiep_vu_GDKT_NQ_01_2025_NQ_HDTP.md); [Quy trình QC 75 / TB 383 / TB 679](../05_Quy_trinh_va_Pham_vi_Chuc_danh/Quy_trinh_Giai_quyet_Don_TANDTC_QC75_TB383_TB679.md). |
| **Biểu Mẫu Chuẩn** | [`01_Flow_Xu_Ly_Don/`](../01_Flow_Xu_Ly_Don/) | Single Source of Truth: [Tờ trình](../01_Flow_Xu_Ly_Don/03_template_to_trinh.md), [Kháng nghị](../01_Flow_Xu_Ly_Don/04_template_khang_nghi.md), [Thông báo](../01_Flow_Xu_Ly_Don/05_template_thong_bao.md) & phôi tại [`templates/`](../01_Flow_Xu_Ly_Don/templates/). |
| **Công Cụ Tech** | [`01_Flow_Xu_Ly_Don/scripts/`](../01_Flow_Xu_Ly_Don/scripts/) | `paths_config.py`, `verify_system.py`, `fast_reader.py`, `export_dossier.py` (Master), `export_to_trinh.py`, `export_khang_nghi.py`, `export_6step_report.py`, `universal_visual_filler.py`. |
| **Hồ Sơ & Đầu Ra** | [`02_Ho_so/`](../02_Ho_so/) → [`03_Ket_qua_Xu_ly/`](../03_Ket_qua_Xu_ly/) | Hồ sơ đầu vào → Đầu ra gom theo vụ án `03_Ket_qua_Xu_ly/<Số>_Vu_<ten_vu>/` *(tự động xếp `01_` ngày tạo mới nhất lên trên cùng)*. |

---

## I. 10 Quy Định Cốt Lõi Bắt Buộc (Mandatory System Rules)

1. **Zero Hallucination, Zero Truncation & Chuẩn Hóa 3 Nhóm Yêu Cầu Tố Tụng:** Chính xác 100%, không suy diễn ngoài hồ sơ ([`02_Ho_so/`](../02_Ho_so/)), không bịa số bản án, ngày tuyên, bút lục, điều luật hay án lệ. Điểm chưa rõ gắn `[CẦN KIỂM CHỨNG]`. Tóm tắt vụ án (Mục I Tờ trình/Báo cáo) **BẮT BUỘC đủ 3 nhóm yêu cầu**: *(1)* YC khởi kiện nguyên đơn (nội dung, có đổi/bổ sung/rút không); *(2)* YC phản tố bị đơn (có theo Đ200 BLTTDS hay không, chỉ là ý kiến/kháng cáo); *(3)* YC độc lập người liên quan (có theo Đ201 BLTTDS hay không, ý kiến và cách Tòa xử lý). CẤM bịa đặt.
2. **Xác Nhận Văn Bản Pháp Lý Trước Khi Đọc Hồ Sơ & Strict Dossier Scope:**
   - Tại Bước 1A, liệt kê trực tiếp bằng **văn bản thuần** các file trong [`04_Van_ban_phap_ly/`](../04_Van_ban_phap_ly/) và [`05_Quy_trinh_va_Pham_vi_Chuc_danh/`](../05_Quy_trinh_va_Pham_vi_Chuc_danh/) kèm đề xuất *(CẤM dùng popup `ask_question`; không quét `.agents/rules/`; không hỏi lại nguồn hồ sơ)*. DỪNG chờ User xác nhận/tag file TRƯỚC KHI mở [`02_Ho_so/`](../02_Ho_so/).
   - **Strict Dossier Scope:** Nếu User đã tag `@file`/`@folder`, CHỈ đọc đúng phạm vi đó, CẤM quét lan sang vụ việc khác. Chỉ quét [`02_Ho_so/`](../02_Ho_so/) khi User chưa chỉ định; nếu có nhiều vụ thì hỏi chọn đúng 01 vụ.
3. **State Persistence (BẮT BUỘC TỪ BƯỚC 2 ĐẾN BƯỚC 6):** Đầu mỗi phản hồi phải in cố định khối:
   ```markdown
   > **[TRẠNG THÁI HỒ SƠ ĐÃ KHÓA]**
   > - **Vụ việc / Đương sự gửi đơn:** [Họ tên] | **Mã đơn / Công văn chuyển đơn & Hạn trả lời:** [Mã / Cơ quan / Hạn hoặc Không có]
   > - **Bản án / QĐ bị khiếu nại:** Sơ thẩm số ... ngày ... (TAND ...); Phúc thẩm số ... ngày ... (TAND ...)
   > - **Thủ tục & Tiền sử đơn:** [GĐT / Tái thẩm / Cả hai / Khác] | [Mới lần đầu / Bổ sung / Trùng lặp / Khiếu nại lại TB cũ]
   > - **Mốc & Phiên bản Luật đã khóa:** Luật nội dung [...] | Luật tố tụng [...]
   > - **Tiến độ luồng:** [Full 6 bước / Early Exit] – Đang thực hiện: Bước [X]
   ```
4. **Early Exit (Ngoại Lệ Duy Nhất Tại Bước 2):** Chỉ đề xuất khi rõ căn cứ: *(1)* hết thời hiệu kháng nghị; *(2)* không thuộc thẩm quyền TANDTC; *(3)* đã có Thông báo của Chánh án TANDTC và không có tình tiết mới hợp pháp; *(4)* thiếu tài liệu tố tụng bắt buộc chưa thể thụ lý. Chỉ sau khi User xác nhận mới đi tắt: **Bước 2 → Bước 5.5 → Bước 6**.
5. **Động Cơ Master Python UTF-8 Duy Nhất & Zero Junk Files:**
   - Xuất toàn bộ `.docx` vào `03_Ket_qua_Xu_ly/<Số>_Vu_<ten_vu>/` *(tự xếp `01_` mới nhất)* qua duy nhất [`export_dossier.py`](../01_Flow_Xu_Ly_Don/scripts/export_dossier.py) (và `openpyxl` cho `.xlsx`).
   - **CẤM tự viết script Python mới** (`run_export_all.py`...) và **CẤM PowerShell COM**. Bắt buộc Dry-run, File-lock check, Placeholder scan, đối chiếu nhất quán 100%. Mọi JSON payload tạm lưu vào `01_Flow_Xu_Ly_Don/scripts/.temp/` và **tự xóa ngay sau khi xuất**.
6. **Single Source of Truth & Phôi Trực Quan:** Định dạng văn bản tuân thủ 100% 3 file mẫu trong [`01_Flow_Xu_Ly_Don/`](../01_Flow_Xu_Ly_Don/). Hỗ trợ Thẩm tra viên tự sửa phôi Word trong [`templates/`](../01_Flow_Xu_Ly_Don/templates/) rồi rót dữ liệu qua [`universal_visual_filler.py`](../01_Flow_Xu_Ly_Don/scripts/universal_visual_filler.py).
7. **Giao Thức Đọc Hồ Sơ Siêu Tốc (`fast_reader.py`):**
   - `.docx`, `.xlsx`, PDF có lớp chữ: chạy `python 01_Flow_Xu_Ly_Don/scripts/fast_reader.py --input "<file>"`.
   - PDF Scan 50–200 trang: *(1)* Định vị Mục lục/Bìa trước (`--pages "1-3"`) tìm đúng trang Đơn, BA sơ thẩm, BA phúc thẩm; *(2)* Render dải trang trọng tâm với `--stitch 2` (DPI 135); *(3)* **Parallel Batch Read** toàn bộ ảnh ghép trong 1 turn, CẤM đọc tuần tự từng trang lẻ.
8. **100% Đường Dẫn Tương Đối & Tự Kiểm Định Khi Tái Cấu Trúc:** Tuyệt đối không ghi cứng `C:\`, `D:\`, `E:\`; mọi script đọc đường dẫn tập trung qua [`paths_config.py`](../01_Flow_Xu_Ly_Don/scripts/paths_config.py). Sau mỗi lần tái cấu trúc, bắt buộc chạy `python 01_Flow_Xu_Ly_Don/scripts/verify_system.py` để kiểm định khớp 100%.
9. **Cấm Tự Ý Quét/Đọc Thư Mục Ẩn & Cấu Hình IDE:** Tuyệt đối **KHÔNG** tự ý quét, đọc hay sửa `.vscode/`, `.git/`, `.env*` hoặc cấu hình ẩn của IDE khi kiểm tra workspace (trừ khi Người dùng yêu cầu đích danh), tránh bật bảng hỏi cấp quyền liên tục.
10. **Zero Ad-Hoc Scripting:** Tuyệt đối **KHÔNG** tự ý tạo file `.py` mới (`read_*.py`, `extract_*.py`...) hay chạy code inline (`python -c`) để đọc hồ sơ hoặc xuất văn bản. Mọi thao tác **đọc file** chỉ gọi [`fast_reader.py`](../01_Flow_Xu_Ly_Don/scripts/fast_reader.py); **xuất Word** chỉ gọi [`export_dossier.py`](../01_Flow_Xu_Ly_Don/scripts/export_dossier.py) *(hoặc [`universal_visual_filler.py`](../01_Flow_Xu_Ly_Don/scripts/universal_visual_filler.py))*. Chỉ sửa code khi Người dùng yêu cầu nâng cấp công cụ.

---

## II. Quy Trình Bắt Buộc 6 Bước (+ Cổng Khóa 5.5)

### 🔹 Bước 1: Khởi Tạo, Xác Nhận Văn Bản Pháp Lý & Khóa Hiệu Lực Luật
1. **Liệt kê danh mục rút gọn (Checkpoint 1A — Tối ưu ~300 tokens):** Liệt kê dạng bullet ngắn `[Tên file] — [Từ khóa < 10 từ]` tại [`04_Van_ban_phap_ly/`](../04_Van_ban_phap_ly/) và [`05_Quy_trinh_va_Pham_vi_Chuc_danh/`](../05_Quy_trinh_va_Pham_vi_Chuc_danh/) gắn nhãn 🏛️ `[KHO 04]`. CẤM trích dẫn nội dung file.
   - **🛑 DỪNG (Checkpoint 1A):** Chờ User xác nhận/tag văn bản pháp lý áp dụng TRƯỚC KHI đọc hồ sơ.
2. Dùng `fast_reader.py` đọc đúng hồ sơ được chỉ định, trích xuất định danh vụ việc, quan hệ pháp luật tranh chấp, và các mốc thời gian cốt lõi (giao dịch, khởi kiện, xét xử ST/PT).
3. **BƯỚC AI TÌM TÀI LIỆU (Checkpoint 1B — Thẻ Tóm Tắt Top 3–5, Tối ưu ~350 tokens):**
   - Dựa trên tình tiết vừa đọc, AI chủ động quét tìm các quy định/án lệ đặc thù hồ sơ cần nhưng kho có thể thiếu, lập danh mục gắn nhãn 🔍 `[AI-TÌM]` (Tên điều luật/Án lệ + 01 câu định vị < 20 từ) gồm 3 tầng: *(1) Luật chuyên ngành*, *(2) Nghị quyết/Thông tư*, *(3) Án lệ TANDTC tương tự*. CẤM chép nguyên văn điều luật hay toàn văn án lệ.
   - **🛑 DỪNG (Checkpoint 1B):** Chờ User duyệt/bổ sung danh mục tài liệu AI đã tìm được.
4. **THẨM ĐỊNH NGỮ CẢNH & KHÓA LUẬT (Checkpoint 1C — Ma Trận 4 Cột, Tối ưu ~450 tokens):**
   - Lập **Bảng Ma Trận Giải Trình Tính Phù Hợp 4 Cột**: *Văn bản/Điều luật* (gắn rõ 🏛️ `[KHO 04]` hoặc 🔍 `[AI-TÌM]`) | *Từ khóa tình tiết thực tế* (ngắn gọn, cấm chép lại bản án) | *Giải trình vì sao phù hợp (1-2 câu)* | *Đề xuất (Áp dụng / Dự phòng)*.
   - Kèm Tóm tắt Phạm vi Nghiên cứu & **Dòng Khóa Rút Gọn** cho State Persistence:
     + *Luật nội dung:* Theo thời điểm xác lập giao dịch / hành vi + điều khoản chuyển tiếp/hồi tố.
     + *Luật tố tụng:* Theo thời điểm thụ lý, xét xử ST/PT và xem xét GĐT/TT.
     + **Auto-Trigger mốc 01/7/2025:** Nếu bất kỳ mốc tố tụng nào từ **01/7/2025 trở đi**, tự kích hoạt [NQ 01/2025/NQ-HĐTP](../04_Van_ban_phap_ly/NQ_01_2025_NQ_HDTP_Tiep_nhan_nhiem_vu_Tham_quyen_TAND.md) và [Cẩm nang GĐKT NQ 01/2025](../05_Quy_trinh_va_Pham_vi_Chuc_danh/Cam_nang_Nghiep_vu_GDKT_NQ_01_2025_NQ_HDTP.md) xuyên suốt Bước 1–5.
- **🛑 DỪNG (Checkpoint 1C):** Chờ User `"xác nhận"` mới sang Bước 2.

### 🔹 Bước 2: Tiền Sử Đơn, Điều Kiện Thụ Lý, Phân Loại GĐT/TT & Q&A
*(In khối `[TRẠNG THÁI HỒ SƠ ĐÃ KHÓA]` ở đầu).*
1. **Tiền sử & Trùng lặp:** Xác định đơn mới / bổ sung / trùng lặp / khiếu nại lại TB cũ (TANDTC, TANDCC, VKSNDTC). Gắn cờ nếu có **công văn chuyển đơn** (Đảng, Quốc hội, ĐBQH, Chính phủ, MTTQ, Báo chí) để theo dõi hạn trả lời & nơi báo cáo.
2. **Hình thức, Chủ thể, Thẩm quyền & Thời hiệu:**
   - **Chốt chặn thẩm quyền TANDTC (Điều 2 & 4 NQ 01/2025):**
     - *Thuộc TANDTC:* BA/QĐ có hiệu lực của TAND cấp tỉnh, TANDCC (cũ), Tòa Phúc thẩm TANDTC.
     - *Thuộc TAND cấp tỉnh (UBTP):* BA/QĐ có hiệu lực của **TAND cấp huyện / TAND khu vực** *(k6, k7 Đ4)*. Nếu nhận đơn thuộc cấp này, **BẮT BUỘC cảnh báo** để Thẩm tra viên chọn: *(1)* Chuyển đơn về TAND cấp tỉnh (`Early Exit`), hoặc *(2)* Báo cáo Lãnh đạo Vụ rút lên TANDTC khi *"xét thấy cần thiết"* *(k3 Đ2)*.
3. **Phân tách GĐT vs. Tái thẩm:** Xác định rõ hướng GĐT / TT / Cả hai. Nếu có tài liệu mới, bắt buộc đánh giá đủ **3 tiêu chí tình tiết mới**: *(1) Tính mới*, *(2) Tính hợp pháp*, *(3) Giá trị chứng minh*.
4. **Đánh giá `Early Exit`:** Nếu thuộc 1 trong 4 trường hợp Mục I.4, nêu rõ căn cứ và đề xuất đi tắt sang Bước 5.5.
5. **Điểm mờ, mâu thuẫn & Q&A:** Liệt kê tài liệu thiếu, mâu thuẫn giữa Đơn – ST – PT – Chứng cứ và đặt câu hỏi `Q1, Q2...` kèm giả định.
- **🛑 DỪNG (Checkpoint 2):** Chờ User trả lời Q&A và chốt hướng (**Tiếp Bước 3** hoặc **`Early Exit` sang Bước 5.5**).

### 🔹 Bước 3: Dòng Thời Gian (Timeline) & Phân Rã Yêu Cầu Khiếu Nại
*(In khối `[TRẠNG THÁI HỒ SƠ ĐÃ KHÓA]` ở đầu).*
1. Lập Timeline tố tụng đầy đủ từng mốc ngày tháng (phát sinh → khởi kiện/khởi tố → ST → PT → THA → gửi đơn/giải quyết trước).
2. Phân rã **100%** yêu cầu khiếu nại theo nhóm (*Tố tụng*, *Chứng cứ/tình tiết*, *Pháp luật nội dung*, *Nghĩa vụ/hình phạt/án phí*, *Tình tiết mới*).
3. Chỉ rõ mối liên hệ phụ thuộc giữa các yêu cầu và chứng cứ cốt lõi.
- **🛑 DỪNG (Checkpoint 3):** Chờ User `"xác nhận"` mới sang Bước 4.

### 🔹 Bước 4: Ma Trận Đối Chiếu Chứng Cứ – Điều Luật & Rà Soát Vi Phạm Ngoài Đơn
*(In khối `[TRẠNG THÁI HỒ SƠ ĐÃ KHÓA]` ở đầu).*
1. **Ma trận đối chiếu 100% yêu cầu Bước 3:** Mã & nội dung khiếu nại | Phán quyết ST/PT | Chứng cứ / Bút lục `(BL ...)` | Điều luật đúng phiên bản đã khóa + Án lệ/NQ HĐTP (kèm nhãn 🏛️ `[KHO 04]` hoặc 🔍 `[AI-TÌM]`) | Nhận xét sơ bộ.
2. **Rà soát độc lập — "Vi Phạm Ngoài Nội Dung Đơn" (BẮT BUỘC):**
   - Bỏ sót người tham gia tố tụng (đặc biệt người liên quan, người đại diện)? Sai quan hệ tranh chấp / thẩm quyền / thời hiệu? Tính sai án phí / chi phí tố tụng / lãi suất? Tuyên án mâu thuẫn nhận định / vượt phạm vi / không thể thi hành án? Vi phạm nghiêm trọng thủ tục tố tụng khác?
   - **Rà soát chuyên sâu NQ 01/2025 (nếu vắt mốc 01/7/2025):**
     - *3 lỗi hủy án nghiêm trọng:* (1) TAND cấp tỉnh nhận đơn trước 01/7 nhưng sau 01/7 mới thụ lý ST thay vì chuyển TAND khu vực *(k1 Đ5)*; (2) Đất đổi địa giới, cấp huyện chưa mở phiên họp giao nộp chứng cứ/hòa giải trước 01/7 mà Tòa khu vực tiếp nhận vẫn xử *(k2 Đ6)*; (3) Bỏ sót cơ quan kế thừa tố tụng khi UBND huyện giải thể / xã sáp nhập *(Đ8)*.
     - *Bộ lọc chống bắt lỗi nhầm (HỢP PHÁP theo Đ7):* (1) Tên Tòa mở phiên trước 01/7 khác tên Tòa ban hành bản án sau 01/7 *(k7)*; (2) Dùng biểu mẫu cũ gạch sửa tên Tòa mới *(k8)*; (3) Tòa tiếp nhận dùng tiếp kết quả tố tụng cũ *(k2)*; (4) Chánh án Tòa tiếp nhận ra QĐ thay đổi người tiến hành tố tụng do sắp xếp bộ máy *(k6)*.
- **🛑 DỪNG (Checkpoint 4):** Chờ User `"xác nhận"` mới sang Bước 5.

### 🔹 Bước 5: Đánh Giá Căn Cứ & Chốt Phương Án Xử Lý
*(In khối `[TRẠNG THÁI HỒ SƠ ĐÃ KHÓA]` ở đầu).*
1. **Bảng đánh giá căn cứ & mức độ vi phạm** cho từng vấn đề trong đơn & ngoài đơn *(Vi phạm nghiêm trọng – Kháng nghị GĐT / Tình tiết mới – Kháng nghị TT / Sai sót nhỏ – Không kháng nghị / Không có căn cứ / Cần xác minh)*.
2. **Tự nêu Đề xuất Chính thức (điền thẳng Mục IV Tờ trình):**
   - **Mục IV.1 (Ý kiến Thư ký viên):** Lập luận chặt chẽ, đủ `(BL ...)`, điều luật đúng phiên bản, câu chốt rõ ràng:
     - *Quy trình rút gọn (TB 383 & 679):* Nếu **Trả lời đơn** → lập Tờ trình rút gọn + dự thảo Thông báo cho Thẩm phán duyệt ký trực tiếp *(không bắt buộc qua Lãnh đạo Vụ; TTV ký nháy, phát hành trong 03 ngày làm việc, gửi 01 bản Vụ trưởng theo dõi)*. Nếu **Kháng nghị** (hoặc vụ phức tạp) → Tờ trình đầy đủ + Dự thảo QĐ Kháng nghị báo cáo họp **Tổ 05 Thẩm phán TANDTC**.
     - *Tên cơ quan mới (NQ 01/2025):* Hủy án ST phải giao **TAND khu vực [tên tiếp nhận mới]**, CẤM giao TAND cấp huyện cũ đã giải thể.
     - *Thẩm quyền xét xử GĐT (k3 Đ2 NQ 01):* Chánh án TANDTC kháng nghị bản án cấp huyện/khu vực phải đề xuất giao **Ủy ban Thẩm phán TAND cấp tỉnh**, KHÔNG giao HĐTP TANDTC.
   - **Mục IV.2 (Dự thảo Ý kiến Thẩm phán):** Tự soạn sẵn bản dự thảo sắc bén, cô đọng để Thẩm phán xem xét/chỉnh sửa *(hoặc để trống nếu User yêu cầu ở Bước 5.5)*. Dự phòng mục duyệt Lãnh đạo Vụ / Phó Chánh án / Chánh án.
3. **Phương án trung gian/khẩn cấp (nếu cần):** Tạm đình chỉ THA, Rút hồ sơ, Xác minh bổ sung, Báo cáo cơ quan chuyển đơn.
- **🛑 DỪNG (Checkpoint 5):** Chờ User chốt Phương án 1, Phương án 2, hoặc làm cả 2 phương án. **CẤM tự ý xuất file tại Bước 5** — bắt buộc chuyển sang Bước 5.5.

### 🔹 Bước 5.5: Bảng Rà Soát Tiếng Việt & Menu Chọn Văn Bản Cần Xuất
*(In khối `[TRẠNG THÁI HỒ SƠ ĐÃ KHÓA]` ở đầu).*
- **CẤM TUYỆT ĐỐI in khối code `JSON`, dấu `{}`, tên biến lập trình hay thuật ngữ kỹ thuật (`Payload`, `Metadata`...) ra màn hình chat.** Trình bày 100% bằng **Bảng Markdown tiếng Việt dễ hiểu** (giải thích ngắn mục đích rà soát lần cuối trước khi xuất Word).
1. **Bảng Rà Soát Thông Tin Hồ Sơ** gồm 6 nhóm: *(1)* Thông tin đơn & nguồn chuyển đơn; *(2)* Danh sách đương sự (thống nhất 100% chính tả, năm sinh, địa chỉ); *(3)* Thông tin Bản án ST/PT; *(4)* Hướng giải quyết chốt & biện pháp kèm theo; *(5)* Thông tin trình duyệt, người ký, tùy chọn Mục IV.2 (`Có`/`Để trống`), nơi nhận; *(6)* Tên các file Word dự kiến lưu tại `03_Ket_qua_Xu_ly/<Số>_Vu_<ten_vu>/` và file `.xlsx` (nếu có).
2. Báo cáo ngắn kết quả tự soát nhất quán (ngày ST/PT, họ tên, số bản án).
3. **BẮT BUỘC IN NGUYÊN VĂN KHUNG CHỌN SAU** *(kể cả khi User đã dặn làm cả 2 phương án ở Bước 5)*:
   ```text
   🛑 CHECKPOINT 5.5 — XIN THẨM TRA VIÊN CHỌN LOẠI VĂN BẢN CẦN XUẤT (lưu chung theo thư mục con vụ án trong 03_Ket_qua_Xu_ly/):
   1. Tờ trình Chánh án TANDTC
   2. Quyết định Kháng nghị Giám đốc thẩm / Tái thẩm
   3. ĐỒNG THỜI CẢ HAI (Tờ trình và Quyết định Kháng nghị)
   4. Thông báo giải quyết đơn (Mẫu Thẩm phán TANDTC duyệt ký trực tiếp theo TB 383/679)
   5. Báo cáo Toàn diện Nghiên cứu 6 Bước
   6. Xuất TẤT CẢ các văn bản trên (Bộ hồ sơ đầy đủ bao gồm cả 2 phương án + Báo cáo 6 bước + Thông báo)

   👉 Kính mời Thẩm tra viên phản hồi số lựa chọn (Ví dụ: "1", "2", "3", "4", "5" hoặc "6") để hệ thống xuất đúng văn bản theo yêu cầu!
   ```
- **🛑 DỪNG (Checkpoint 5.5):** Chờ User chọn rõ số `1–6`. **KHÓA BẢO VỆ:** Nếu User chỉ gõ `"OK"`/`"xác nhận"` mà chưa nêu số/loại file, CẤM xuất file — phải hỏi lại ngay.

### 🔹 Bước 6: Xuất File Word (.docx) Qua Động Cơ Master (< 3 Giây, 01 Turn Duy Nhất)
*(In khối `[TRẠNG THÁI HỒ SƠ ĐÃ KHÓA]` ở đầu).*
1. **Không đọc lại file demo** (`demo_payload.json`...) — dữ liệu đã đầy đủ trong hội thoại. **CẤM tự viết script Python mới.**
2. Thực thi trong **01 turn duy nhất** đúng 2 thao tác:
   - `write_file` lưu dữ liệu ngầm vào `01_Flow_Xu_Ly_Don/scripts/.temp/payload_<ten_vu>.json` *(không in JSON ra chat)*.
   - `run_command` gọi duy nhất:
     ```bash
     python 01_Flow_Xu_Ly_Don/scripts/export_dossier.py --payload "01_Flow_Xu_Ly_Don/scripts/.temp/payload_<ten_vu>.json" --option <1|2|3|4|5|6> --case-slug "<ten_slug>"
     ```
3. `export_dossier.py` tự kiểm tra file-lock, chống `KeyError`, tự xếp số `01_Vu_<ten_slug>/` (mới nhất lên trên cùng) trong `03_Ket_qua_Xu_ly/`, cập nhật `README.md` mục lục và tự xóa file tạm. Báo cáo hoàn tất kèm link trực tiếp tới từng file đã xuất.

---

## III. Anti-Patterns (TUYỆT ĐỐI NGHIÊM CẤM)

- ❌ Tự viết script Python mới (`read_*.py`, `extract_*.py`, `run_export_all.py`, `run_generation.py`...), chạy `python -c` để đọc/xuất file, hoặc dùng PowerShell COM ở Bước 1–6 (bắt buộc dùng duy nhất `fast_reader.py` để đọc và `export_dossier.py` / `universal_visual_filler.py` để xuất).
- ❌ Đọc lại `demo_*payload.json` ở Bước 6 hoặc xả rác file tạm trong `01_Flow_Xu_Ly_Don/scripts/`.
- ❌ In JSON / dấu `{}` / tên biến lập trình ra màn hình chat ở Bước 5.5.
- ❌ Đọc mù PDF scan từng trang lẻ qua nhiều turn (phải định vị mục lục + `--stitch 2` + Batch Read 1 turn).
- ❌ Chạy gộp bước, vượt Checkpoint, tự ý Early Exit, hoặc tự xuất file khi User chưa chọn rõ số `1–6` ở Checkpoint 5.5.
- ❌ Đọc `02_Ho_so/` trước Checkpoint 1A hoặc đọc lan sang hồ sơ khác ngoài phạm vi User chỉ định.
- ❌ Quên khối `[TRẠNG THÁI HỒ SƠ ĐÃ KHÓA]` từ Bước 2 đến Bước 6.
- ❌ Áp sai mốc hiệu lực luật, nhầm lẫn GĐT/Tái thẩm, bỏ sót rà soát vi phạm ngoài đơn ở Bước 4.
- ❌ Ôm việc của TAND cấp tỉnh (bỏ qua cảnh báo thẩm quyền bản án cấp huyện/khu vực) hoặc bắt lỗi nhầm thủ tục chuyển tiếp hợp pháp theo Điều 7 NQ 01/2025/NQ-HĐTP.
- ❌ Ghi cứng đường dẫn ổ đĩa cá nhân (`C:\`, `D:\`, `E:\`) vào script hoặc tài liệu hệ thống.
- ❌ Tự ý quét/đọc `.vscode/`, `.git/`, `.env*` hoặc thư mục cấu hình ẩn của IDE khi Người dùng không yêu cầu đích danh.
