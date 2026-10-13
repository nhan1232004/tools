# GGT-CNTT Auto — Tự động tạo Giấy giới thiệu thực tập

Sinh viên điền Google Form → hệ thống tự điền vào mẫu giấy giới thiệu → xuất **PDF** → gửi vào **email giảng viên**.
Viết bằng Google Apps Script, không cần máy chủ.

```
ggt-cntt-auto/
├── README.md                          (hướng dẫn này)
├── templates/
│   ├── GGT_mau_goc.docx               mẫu Word gốc của Khoa
│   └── GGT_mau_placeholder.docx       mẫu Word đã thay chỗ trống bằng placeholder
└── sources/
    └── Code.gs                        mã Apps Script
```

## Link mẫu

| Mẫu | Link (bấm để **Tạo bản sao**) |
|---|---|
| Google Sheet mẫu | `https://docs.google.com/spreadsheets/d/1X5b2PKFYbYgO0GO45CS19dwK-CQmms_CC4jybYxQwa4/edit?resourcekey=&gid=703923767#gid=703923767` |
| Google Form mẫu | `https://docs.google.com/forms/d/19t2QF0-52aXpKkaYTeiwm8Ud-zNw5WE6s5rvISh4QmI/copy` |

> Mẹo: đổi đuôi `/edit` của link thành `/copy` để người mở thấy ngay nút **Tạo bản sao**.
> Nếu chưa có Form mẫu, tự tạo Form theo Bước 2 bên dưới.

## 6 trường của Form và placeholder trong mẫu Word

| # | Câu hỏi trong Form | Placeholder |
|---|---|---|
| 1 | Họ và tên | `{{HO_TEN}}` |
| 2 | Mã số sinh viên - MSSV | `{{MSSV}}` |
| 3 | Email cá nhân | `{{EMAIL}}` |
| 4 | Số điện thoại | `{{SDT}}` |
| 5 | Vị trí thực tập | `{{VI_TRI}}` |
| 6 | Tên Công ty / Doanh nghiệp thực tập | `{{CONG_TY}}` |

"Số: ....-GGT-CNTT" và ngày ký **để trống** cho giảng viên điền sau.

---

## Hướng dẫn cài đặt

### Bước 1. Tạo Google Doc mẫu

1. Vào Google Drive → **Mới → Tải tệp lên** → chọn `templates/GGT_mau_placeholder.docx`.
2. Mở file vừa tải lên → **Tệp → Lưu dưới dạng Google Tài liệu**.
3. Kiểm tra Doc có đủ 6 placeholder ở bảng trên (viết đúng, liền một mạch).
4. Xuất thử PDF từ Docs (**Tệp → Tải xuống → PDF**) để chắc nội dung vừa **một trang**.
5. Copy **ID của Doc** trên thanh địa chỉ, là đoạn nằm giữa `/d/` và `/edit`:

```
https://docs.google.com/document/d/  ĐÂY_LÀ_ID_DOC  /edit
```

### Bước 2. Tạo Google Form

Dùng **Form mẫu** (Tạo bản sao) hoặc tự tạo Form mới với đúng **6 câu hỏi dạng "Câu trả lời ngắn"**, đều đặt **Bắt buộc**, tiêu đề đặt chính xác như sau:

1. Họ và tên
2. Mã số sinh viên - MSSV
3. Email cá nhân
4. Số điện thoại
5. Vị trí thực tập
6. Tên Công ty / Doanh nghiệp thực tập

Không bật "Thu thập địa chỉ email". Nên bật kiểm tra định dạng: MSSV (biểu thức chính quy `^[A-Za-z0-9]{6,12}$`), Email cá nhân (văn bản là email), Số điện thoại (`^(0|\+84)[0-9]{9}$`).

> Script tìm cột theo tiêu đề câu hỏi, nên nếu tự tạo Form thì đừng đổi tên các tiêu đề trên.

### Bước 3. Liên kết Form với Google Sheet

1. Trong Form, mở tab **Câu trả lời** → bấm biểu tượng **Liên kết với Trang tính**.
2. Chọn **Tạo bảng tính mới**, hoặc **Chọn bảng tính hiện có** (dùng bản sao của Sheet mẫu).
3. Mở Sheet vừa liên kết, kiểm tra có tab câu trả lời với 6 cột đúng tiêu đề.

### Bước 4. Tạo thư mục Drive chứa PDF

1. Trong Drive: **Mới → Thư mục mới**, đặt tên ví dụ `GGT_PDF`.
2. Mở thư mục, copy **ID thư mục** trên thanh địa chỉ, là đoạn sau `/folders/`:

```
https://drive.google.com/drive/folders/  ĐÂY_LÀ_ID_THƯ_MỤC
```

### Bước 5. Đưa `Code.gs` vào Apps Script và sửa cấu hình

1. Mở Google Sheet ở Bước 3 → **Tiện ích mở rộng → Apps Script**.
   (Nếu dùng bản sao Sheet mẫu đã có sẵn code thì bỏ qua việc dán, chỉ sửa cấu hình.)
2. Xóa mã mặc định, dán toàn bộ nội dung `sources/Code.gs`.
3. Sửa 3 dòng đầu của khối `CONFIG`:

```javascript
const CONFIG = {
  TEMPLATE_DOC_ID: 'ID_DOC_Ở_BƯỚC_1',
  OUTPUT_FOLDER_ID: 'ID_THƯ_MỤC_Ở_BƯỚC_4',
  LECTURER_EMAIL: 'email_nhận_pdf@gmail.com',
  CC_EMAILS: '',        // nhiều mail cách nhau dấu phẩy, có thể để trống
  ADMIN_EMAIL: '',      // nhận mail báo lỗi, để trống = chính tài khoản chạy script
  ...
};
```

   **Chỉ dán ID, không dán cả đường link.** Giữ nguyên dấu nháy `' '`, dấu phẩy và dấu `};` đóng khối.
4. Bấm **Lưu** (biểu tượng đĩa mềm). Vào **Cài đặt dự án (bánh răng)** và đặt múi giờ `Asia/Ho_Chi_Minh`.

### Bước 6. Cài đặt trigger và cấp quyền

1. Chọn hàm **`setup`** trên thanh công cụ → **Chạy**.
2. Cấp quyền khi Google hỏi: **Xem lại quyền → chọn tài khoản → Nâng cao → Đi tới dự án (không an toàn) → Cho phép**.
3. `setup` sẽ tạo trigger "khi gửi biểu mẫu", thêm 3 cột (Trạng thái, Link PDF, Thời gian xử lý) và báo lỗi nếu Doc mẫu thiếu placeholder.

### Bước 7. Kiểm thử

1. Mở link Form, nộp thử một phiếu bằng dữ liệu giả hợp lệ.
2. Đợi 10 đến 30 giây: hộp thư giảng viên nhận mail kèm PDF, dòng mới trong Sheet có trạng thái `Đã gửi`.
3. Muốn thử lại mà không cần nộp Form: chọn hàm **`testLastRow`** → **Chạy**.
4. Mở PDF kiểm tra: đúng một trang, thông tin đúng, "Số" và ngày để trống.

---

## Các hàm trong `Code.gs`

| Hàm | Dùng khi |
|---|---|
| `setup` | Chạy **một lần** sau khi sửa cấu hình: tạo trigger, kiểm tra Doc mẫu |
| `testLastRow` | Chạy thử với dòng dữ liệu cuối cùng |
| `retryPending` | Xử lý lại các dòng chưa gửi hoặc đang báo lỗi (sau khi sửa dữ liệu) |
| `handleFormSubmit` | Trigger tự gọi khi có phiếu mới, **không bấm Chạy tay** |

## Lỗi thường gặp

| Hiện tượng | Cách xử lý |
|---|---|
| `Doc mẫu thiếu placeholder: ...` | Mở Doc, gõ lại đúng `{{HO_TEN}}`, `{{MSSV}}`, ... liền một mạch |
| `Không tìm thấy cột cho "..."` | Tiêu đề câu hỏi trong Form bị đổi, đặt lại đúng 6 tiêu đề |
| `Không tìm thấy sheet câu trả lời` | Form chưa liên kết với Sheet (Bước 3) |
| Lỗi truy cập Doc hoặc thư mục | Dán nhầm cả link thay vì ID, hoặc tài khoản chạy script không có quyền với Doc/thư mục |
| `Lỗi: Dữ liệu không hợp lệ ...` ở cột Trạng thái | Sửa ô sai trong Sheet rồi chạy `retryPending` |
| PDF thừa trang trắng | Script đã tự dọn; nếu vẫn thừa, giảm lề dưới hoặc giãn cách trong Doc mẫu |
| Không có mail sau khi nộp Form | Chạy lại `setup` để tạo lại trigger |

## Lưu ý

- Mail gửi đi từ **tài khoản chạy script**. Gmail thường gửi tối đa khoảng 100 mail mỗi ngày, Workspace khoảng 1.500.
- Mỗi giảng viên tạo **bản sao riêng** và chạy `setup` một lần: trigger và quyền truy cập không đi theo khi sao chép Sheet.
- Không đưa ID thật, email thật hay dữ liệu sinh viên lên repo công khai.
