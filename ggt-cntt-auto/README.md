# GGT-CNTT Auto — Hệ thống Tự động tạo Giấy giới thiệu thực tập (Khoa CNTT)

Hệ thống tự động hóa toàn diện quy trình cấp Giấy giới thiệu thực tập (GGT) cho sinh viên học phần Kiến tập công nghiệp (KTCN) & Tập sự nghề nghiệp (TSNN).

**Luồng hoạt động chính:**
1. Sinh viên đăng ký qua **Google Form** (thông tin tự lưu vào **Google Sheet**).
2. Hệ thống tự động **bóc tách, chuẩn hóa dữ liệu** (xử lý định dạng số điện thoại, MSSV và lỗi số khoa học `E+`).
3. **Kiểm chứng Mã số thuế (MST)** doanh nghiệp qua API Tổng cục Thuế: `https://api.xinvoice.vn/gdt-api/tax-payer/{mst}`.
4. **Đối chứng so sánh tên công ty** do sinh viên nhập với tên đăng ký của cơ quan Thuế:
   - Nhận diện chi nhánh: ưu tiên phần tên công ty tổng ở phía trước khớp chính xác.
   - Đối chiếu tên viết tắt, tên tiếng nước ngoài và chuẩn hóa các danh xưng viết tắt (`CTY`, `TNHH`, `MTV`, `CP`, `JSC`...).
   - Nếu sai MST hoặc tên công ty khác biệt rõ rệt -> ghi nhận không hợp lệ, không xuất giấy.
5. **Lọc thời gian nộp đơn**: Tự động lọc chỉ xử lý các yêu cầu bắt đầu từ đầu tháng **08/2026** tới mới nhất hiện nay (`FILTER_START_DATE: '2026-08-01'`).
6. Tự động điền dữ liệu vào mẫu Google Doc -> Xuất file **PDF (vừa khít đúng 1 trang)** -> Lưu vào Google Drive (tự động phân quyền xem qua liên kết).
7. **Ghi link PDF và kết quả trực tiếp vào Google Sheet**: Không gửi email cho giảng viên (`SEND_EMAIL: false`), giảng viên chỉ cần bấm vào link trong cột `Link PDF` trên Sheet để mở xem ngay file PDF.

---

## Cấu trúc thư mục

```
d:/TSNN_KTCN/tool/
├── yêu cầu tool.docx                              # Tài liệu bài toán & yêu cầu nghiệp vụ
├── Cấp Giấy Giới Thiệu Thực Tập (Responses).xlsx # Sheet thực tế của Khoa (16 cột)
└── ggt-cntt-auto/
    ├── README.md                                  # Hướng dẫn chi tiết hệ thống
    ├── templates/
    │   ├── GGT_mau_goc.docx                       # Mẫu Word hành chính chuẩn của Khoa
    │   └── GGT_mau_placeholder.docx               # Mẫu Word đã gắn sẵn các placeholder
    └── sources/
        └── Code.gs                                # Mã nguồn Google Apps Script hoàn chỉnh
```

---

## 1. Các trường thông tin & Placeholder tương ứng

### Danh sách Placeholder trong Google Doc mẫu (`templates/GGT_mau_placeholder.docx`)

| # | Trường thông tin | Placeholder | Ghi chú |
|---|---|---|---|
| 1 | Họ và tên | `{{HO_TEN}}` | Viết hoa chữ cái đầu (Title Case) |
| 2 | Mã số sinh viên | `{{MSSV}}` | Đã khử định dạng số khoa học `E+` |
| 3 | Email cá nhân | `{{EMAIL}}` | Viết thường, chuẩn RFC |
| 4 | Số điện thoại | `{{SDT}}` | Tự động chuẩn hóa đầu 0 và +84 |
| 5 | Vị trí thực tập | `{{VI_TRI}}` | Mặc định "Thực tập sinh" nếu để trống |
| 6 | Tên Công ty / Doanh nghiệp | `{{CONG_TY}}` | Tên doanh nghiệp thực tập |
| 7 | Mã số thuế (Tùy chọn) | `{{MST}}` | Tự động thay thế nếu có trong mẫu |
| 8 | Học phần (Tùy chọn) | `{{HOC_PHAN}}` | KTCN hoặc TSNN |
| 9 | Địa chỉ công ty (Tùy chọn) | `{{DIA_CHI_CTY}}` | Địa chỉ theo giấy phép Thuế |

> *Lưu ý*: Phần `Số: ....-GGT-CNTT` và ngày ký trong mẫu được giữ nguyên để giảng viên/văn phòng Khoa cấp số và ký duyệt.

---

## 2. Quy tắc đối soát Mã số thuế và Tên Doanh nghiệp

Theo yêu cầu nghiệp vụ bắt buộc:
1. **Kiểm tra cú pháp MST**: MST phải là 10 chữ số (doanh nghiệp) hoặc 13-14 ký tự (`XXXXXXXXXX-XXX` đối với chi nhánh). Tự bù số 0 nếu bảng tính làm mất định dạng.
2. **Gọi API kiểm tra**: Gọi `https://api.xinvoice.vn/gdt-api/tax-payer/{mst}` để lấy dữ liệu thực tế.
   - Nếu mã phản hồi `404` hoặc không có dữ liệu -> Báo lỗi: `mã số thuế không hợp lệ hoặc không tồn tại`.
   - Nếu trạng thái là `ngừng hoạt động` -> Báo lỗi: `doanh nghiệp đã ngừng hoạt động`.
3. **Thuật toán so sánh tên (Matching Algorithm)**:
   - **Xử lý Chi nhánh**: Nếu doanh nghiệp là chi nhánh, trích xuất phần tên công ty tổng ở phía trước (trước dấu `-` hoặc từ "chi nhánh"). Ưu tiên so sánh phần công ty tổng này với tên sinh viên nhập.
   - **Chuẩn hóa chữ viết tắt**: `cty`, `ct` -> `công ty`; `tnhh` -> `trách nhiệm hữu hạn`; `mtv` -> `một thành viên`; `cp`, `jsc` -> `cổ phần`; `ltd` -> `trách nhiệm hữu hạn`.
   - **Trọng số từ khóa**: Tên thương hiệu riêng biệt (như *Tường Minh*, *FPT*, *SiGlaz*, *Delements*) có trọng số 1.0; từ chỉ lĩnh vực chung (*phần mềm*, *công nghệ*, *dịch vụ*, *thương mại*) có trọng số 0.35.
   - **Đối chiếu mở rộng**: Đối chiếu thêm tên tiếng Anh và tên viết tắt (viết tắt như *ISD*, *ITD*, *TMA*, *FPT*...).
   - **Nguyên tắc từ chối**: Nếu đối chiếu so sánh có sự khác biệt rõ rệt không tương đồng, hệ thống ghi nhận lỗi: `mã số thuế hoặc tên công ty không hợp lệ`, không tạo bản PDF và không gửi email.

---

## 3. Hướng dẫn cài đặt từng bước

### Bước 1: Chuẩn bị Google Doc mẫu
1. Mở [Google Drive](https://drive.google.com).
2. Tải tệp `templates/GGT_mau_placeholder.docx` lên Drive.
3. Mở tệp vừa tải lên -> Chọn **Tệp -> Lưu dưới dạng Google Tài liệu**.
4. Sao chép **Doc ID** trên thanh địa chỉ URL:
   ```
   https://docs.google.com/document/d/ [ĐÂY_LÀ_DOC_ID] /edit
   ```

### Bước 2: Chuẩn bị Thư mục lưu trữ PDF
1. Trên Google Drive, tạo một thư mục mới (ví dụ đặt tên `GGT_Xuat_PDF`).
2. Mở thư mục và sao chép **Folder ID** trên URL:
   ```
   https://drive.google.com/drive/folders/ [ĐÂY_LÀ_FOLDER_ID]
   ```

### Bước 3: Cài đặt Google Apps Script
1. Mở Google Sheet nhận câu trả lời Form (hoặc Sheet `Cấp Giấy Giới Thiệu Thực Tập (Responses)`).
2. Vào menu **Tiện ích mở rộng** (Extensions) -> **Apps Script**.
3. Xóa toàn bộ nội dung hiện tại, sao chép toàn bộ mã từ file `sources/Code.gs` và dán vào.
4. Cập nhật các thông số trong khối `CONFIG` ở đầu file:
   ```javascript
   const CONFIG = {
     TEMPLATE_DOC_ID: 'DÁN_DOC_ID_VÀO_ĐÂY',
     OUTPUT_FOLDER_ID: 'DÁN_FOLDER_ID_VÀO_ĐÂY',
     SEND_EMAIL: false,             // Không gửi email cho giảng viên
     FILTER_START_DATE: '2026-08-01', // Chỉ xử lý yêu cầu từ đầu tháng 8/2026 tới nay
     LECTURER_EMAIL: 'email_giangvien@tdtu.edu.vn',
     ...
   };
   ```
5. Nhấn **Lưu** (Ctrl + S).

### Bước 4: Khởi tạo và cấp quyền
1. Trong màn hình Apps Script, chọn hàm **`setup`** từ danh sách hàm -> Nhấn **Chạy (Run)**.
2. Cấp quyền truy cập cho Script khi Google hiển thị hộp thoại xác thực:
   - Chọn tài khoản Google của bạn -> Bấm **Nâng cao (Advanced)** -> Bấm **Đi tới dự án (không an toàn)** -> Bấm **Cho phép (Allow)**.
3. Hàm `setup` sẽ:
   - Tự động tạo Trigger kích hoạt mỗi khi có Form submit mới.
   - Tự động bổ sung các cột quản trị vào Sheet: `TRẠNG THÁI GỬI GGT`, `LINK PDF GGT`, `THỜI GIAN XỬ LÝ GGT`, `Tên Doanh Nghiệp (Tra cứu MST)`, `Kết quả đối soát MST` (tự động nhận diện cột đã có sẵn, không tạo trùng lặp).
   - Kiểm tra kết nối Doc mẫu và Thư mục Drive.

---

## 4. Các hàm quản trị và kiểm thử trong `Code.gs`

| Tên hàm | Chức năng | Cách dùng |
|---|---|---|
| `onOpen` | Tự động tạo menu `GGT-CNTT` trên thanh công cụ Google Sheet | Tự động chạy khi mở bảng tính; cho phép thực hiện kiểm thử và quản trị trực tiếp từ giao diện Sheet |
| `setup` | Cài đặt hệ thống, tạo trigger và thêm cột | Chạy 1 lần sau khi cấu hình |
| `testTaxValidationManual` | Kiểm tra trực tiếp thuật toán đối soát MST trên console | Chạy thử ngay trên editor không cần Sheet |
| `testLastRow` | Chạy thử nghiệm với dòng cuối cùng trong bảng tính | Dùng khi muốn kiểm tra quy trình xuất PDF và ghi link |
| `testRowPrompt()` / `testRow2()` | Chạy thử nghiệm với hộp thoại UI prompt hỏi số dòng (không cần truyền tham số) | Chạy trực tiếp từ menu `GGT-CNTT > Kiểm thử dòng cụ thể...` trên Sheet hoặc từ dropdown Apps Script Editor |
| `testMultipleRows()` | Chạy thử nghiệm nhiều dòng cùng lúc qua hộp thoại UI (hỗ trợ nhập `2, 5, 7-10`) | Chạy trực tiếp từ menu `GGT-CNTT > Kiểm thử nhiều dòng...` |
| `testRow(rowNumber)` | Chạy thử nghiệm với một dòng cụ thể (tự mở prompt nếu không truyền tham số) | Gọi `testRow(2)` qua script call hoặc chạy từ dropdown editor |
| `processBatchPrompt()` | In PDF hàng loạt nhiều dòng cùng lúc qua hộp thoại UI (hỗ trợ nhập `2, 5, 7-10`) | Chạy trực tiếp từ menu `GGT-CNTT > In PDF hàng loạt...` |
| `processBatch(from, to)` | Xử lý một khoảng dòng cụ thể (có giãn cách) | Ví dụ: `processBatch(2, 20)` |
| `retryPending` | Quét và xử lý lại tất cả các dòng bị lỗi hoặc chưa gửi | Chạy sau khi giảng viên/sinh viên đã chỉnh sửa dữ liệu sai |

---

## 5. Xử lý sự cố thường gặp (Troubleshooting)

| Hiện tượng | Nguyên nhân | Cách khắc phục |
|---|---|---|
| `Lỗi: mã số thuế không hợp lệ hoặc không tồn tại` | Sinh viên gõ nhầm số thuế, hoặc gõ chữ ("SPS", "Không có", CCCD 12 số) | Hướng dẫn sinh viên tra cứu MST chính xác trên trang Thuế, sửa lại ô trong Sheet rồi chạy `retryPending()`. |
| `Lỗi: mã số thuế hoặc tên công ty không hợp lệ` | Sinh viên đăng ký công ty A nhưng nhập MST của công ty B, hoặc gõ sai tên hoàn toàn | Kiểm tra cột `Tên Doanh Nghiệp (Tra cứu MST)` để thấy tên thật, liên hệ sinh viên xác nhận và cập nhật lại Sheet. |
| `API thuế phản hồi mã 429 (Too Many Requests)` | Vượt quá hạn mức 10 yêu cầu / 30 giây của `api.xinvoice.vn` | Script đã tự động đọc header `Retry-After`, áp dụng cơ chế Retry Exponential Backoff (chờ 6-12s) và đặt `CONFIG.BATCH_SLEEP_MS = 3200` để chạy hàng loạt mượt mà. |
| File PDF bị tràn sang trang 2 | Mẫu Doc gốc có khoảng trắng hoặc lề quá lớn | Script đã tích hợp hàm `trimTrailingBlank_`. Nếu vẫn bị, mở Google Doc mẫu và giảm lề trang (Margin) xuống còn 1.5 cm - 2.0 cm. |
| MSSV, MST hoặc Số điện thoại dạng khoa học `E+` | Excel/Sheet tự động chuyển số dài thành `5.19E+07` hoặc `7.75E+08` | Script tích hợp bộ giải mã `normalizeMssv_`, `normalizeMst_` và `normalizePhone_` tự động khôi phục số nguyên và thêm đầu `0` chuẩn xác. |
| Doanh nghiệp chi nhánh có tên trong ngoặc đơn `(...)` | Sinh viên ghi chú chi nhánh trong ngoặc hoặc API thuế thêm `(LOẠI HÌNH DOANH NGHIỆP: ...)` | Hàm `extractParentCompanyName_` tự động bóc tách và làm sạch phần ngoặc đơn, ưu tiên đối khớp tên công ty tổng theo đúng đề bài. |
| Tên viết tắt có kèm loại hình (`ITD Corporation`, `TMA SOLUTIONS CO.`) | Viết tắt nhiều từ chứa các từ pháp lý chung | Thuật toán tự động chuẩn hóa và lọc stopwords để đối soát chính xác từ khóa thương hiệu (`ITD`, `TMA`, `ISD`, `CEP`). |
