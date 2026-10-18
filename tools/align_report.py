#!/usr/bin/env python3
"""Apply deterministic corrections to the generated internship report."""
from pathlib import Path
from docx import Document

ROOT = Path(__file__).resolve().parents[1]
PATH = ROOT / "BaoCao_KTCN_GGT_CNTT_Auto_522H0045.docx"
COMPANY = "CÔNG TY CỔ PHẦN TRÍ TUỆ NHÂN TẠO THẾ HỆ MỚI"
doc = Document(PATH)


def set_paragraph(p, text):
    if p.runs:
        p.runs[0].text = text
        for run in p.runs[1:]:
            run.text = ""
    else:
        p.add_run(text)


def set_cell(cell, text):
    cell.text = text


# Fix complete paragraphs instead of replacing fragments that might split across Word runs.
paragraph_rewrites = {
    "Theo số liệu thực tế lưu trong bảng tính của Khoa Công nghệ Thông tin - Trường Đại học Tôn Đức Thắng, tính đến tháng 10/2026 có 361 hồ sơ đăng ký cấp giấy giới thiệu thực tập của sinh viên các học phần Kiến tập công nghiệp, Tập sự nghề nghiệp và Dự án CNTT 1; trong các học kỳ cao điểm số lượng đăng ký có thể lên tới hàng trăm hồ sơ. Với cách làm thủ công, mỗi hồ sơ cần khoảng 5 phút thao tác sao chép, kiểm tra và định dạng, tương đương khoảng 30 giờ làm việc cho toàn bộ số hồ sơ nêu trên.":
    "Tệp Excel hiện có của bảng tính Khoa CNTT chứa 999 hồ sơ và 20 cột. Code.gs cấu hình FILTER_START_DATE = 2026-08-01, nên chỉ nhận xử lý hồ sơ từ ngày 01/08/2026. Kết quả 361 hồ sơ (330 PDF, 31 bị chặn) là bản tổng hợp của lô vận hành ngày 17–18/10/2026, không phải toàn bộ workbook. Mốc thủ công 5 phút/hồ sơ là ước lượng, tương đương 30 giờ cho 361 hồ sơ." ,
    "Từ kết quả thực nghiệm trên 361 hồ sơ thật, có thể rút ra một số nhận xét:":
    "Từ kết quả thực nghiệm trên lô 361 hồ sơ trong bản tổng hợp vận hành, có thể rút ra một số nhận xét:",
    "Tỷ lệ cấp giấy tự động đạt 91,41% (330/361). Phần còn lại không phải lỗi hệ thống mà là hồ sơ có dữ liệu sai; việc chặn và ghi rõ nguyên nhân giúp sinh viên sửa đúng chỗ và chạy lại.":
    "Trong lô vận hành được báo cáo, tỷ lệ cấp giấy tự động đạt 91,41% (330/361). Phần còn lại không phải lỗi hệ thống mà là hồ sơ có dữ liệu sai; việc chặn và ghi rõ nguyên nhân giúp sinh viên sửa đúng chỗ và chạy lại.",
    "Thời gian xử lý tự động được trích xuất từ chính cột “THỜI GIAN XỬ LÝ GGT” do hệ thống ghi lại, nên phản ánh đúng nhịp chạy thật thay vì số đo trong phòng thí nghiệm. Đợt chạy ngày 17–18/10/2026 gồm bốn đoạn xử lý liên tục: 10 hồ sơ trong 3 phút 26 giây, 42 hồ sơ trong 9 phút 46 giây, 40 hồ sơ trong 8 phút 37 giây và 269 hồ sơ trong 60 phút 51 giây. Tổng thời gian máy chạy thực tế cho cả 361 hồ sơ là 82,7 phút, tương đương 13,7 giây cho một hồ sơ, tức khoảng 4,4 hồ sơ mỗi phút.":
    "Bản tổng hợp đợt chạy ngày 17–18/10/2026 ghi bốn đoạn xử lý: 10 hồ sơ trong 3 phút 26 giây, 42 hồ sơ trong 9 phút 46 giây, 40 hồ sơ trong 8 phút 37 giây và 269 hồ sơ trong 60 phút 51 giây. Tổng thời gian máy chạy báo cáo cho lô 361 hồ sơ là 82,7 phút, tương đương 13,7 giây/hồ sơ hoặc khoảng 4,4 hồ sơ/phút." ,

    "Đơn vị tiếp nhận thực tập là CÔNG TY CỔ PHẦN TRÍ TUỆ NHÂN TẠO THẾ HỆ MỚI, trực thuộc Khoa Công nghệ Thông tin, Trường Đại học Tôn Đức Thắng (Số 19 Nguyễn Hữu Thọ, Phường Tân Phong, Quận 7, TP. Hồ Chí Minh). Thông tin địa chỉ, quy mô nhân sự, lĩnh vực hoạt động và sản phẩm cụ thể của doanh nghiệp chưa được đưa vào do chưa có tài liệu xác nhận.":
    f"Đơn vị tiếp nhận thực tập là {COMPANY}. Sinh viên thực tập tham gia xây dựng GGT-CNTT Auto, hệ thống tự động hóa quy trình cấp Giấy giới thiệu thực tập của Khoa CNTT. Chưa bổ sung địa chỉ, quy mô, lĩnh vực hoạt động hoặc sản phẩm doanh nghiệp do chưa có tài liệu xác thực.",
    "Trong thời gian thực tập tại CÔNG TY CỔ PHẦN TRÍ TUỆ NHÂN TẠO THẾ HỆ MỚI, sinh viên tham gia nhóm dự án theo tiến độ công việc được giao. Hoạt động chuyên môn gồm khảo sát yêu cầu, cập nhật mã nguồn qua Git/GitHub, triển khai từng chức năng, kiểm tra bằng bộ test và đối chiếu kết quả với dữ liệu thực tế. Phạm vi đề tài gắn với quy trình cấp Giấy giới thiệu thực tập của Khoa CNTT; đơn vị tiếp nhận thực tập là doanh nghiệp nêu trên.":
    f"Trong thời gian thực tập tại {COMPANY}, sinh viên tham gia nhóm dự án theo công việc được giao. Nhiệm vụ gồm khảo sát yêu cầu, phát triển và kiểm thử mã nguồn Google Apps Script, xử lý dữ liệu biểu mẫu, đối soát mã số thuế và tên doanh nghiệp, xuất PDF, ghi trạng thái vào Google Sheets và đồng bộ tiến độ với Code.gs hiện tại.",
}
for p in doc.paragraphs:
    if p.text in paragraph_rewrites:
        set_paragraph(p, paragraph_rewrites[p.text])
    elif p.text.startswith("Hình 2.2: Sơ đồ cơ cấu tổ chức"):
        set_paragraph(p, "Hình 2.2: Sơ đồ đơn vị thực tập và phạm vi nhiệm vụ dự án")
    elif p.text.startswith("Hệ thống đã được kiểm nghiệm thực tế trên 92 hồ sơ"):
        set_paragraph(p, "Bản tổng hợp vận hành ngày 17–18/10/2026 ghi nhận một lô 361 hồ sơ: 330 PDF được tạo (91,41%), 31 hồ sơ bị chặn (8,59%). Bộ kiểm thử tự động đạt 93/93 ca. Tệp Excel hiện có chứa 999 hồ sơ và 20 cột; Code.gs cấu hình FILTER_START_DATE = 2026-08-01, nên số liệu 361 là lô vận hành được báo cáo, không phải toàn bộ workbook.")
    elif p.text.startswith("The issuance of internship recommendation letters"):
        set_paragraph(p, "This report presents GGT-CNTT Auto, a Google Apps Script system for automating internship recommendation letters. The attached workbook contains 999 records across 20 columns, and the current Code.gs filters records from 1 August 2026 (FILTER_START_DATE = 2026-08-01). The supplied operation summary reports one batch of 361 records, with 330 PDFs generated and 31 blocked; this batch is not represented as the entire workbook. Automated tests report 93 passing cases. Gemini review is optional and disabled by default, while bulk processing uses a checkpointed queue.")
    elif p.text.startswith("Đợt thực nghiệm được thực hiện trên 361 hồ sơ được tổng hợp"):
        set_paragraph(p, "Đợt thực nghiệm sử dụng số liệu lô 361 hồ sơ trong bản tổng hợp vận hành ngày 17–18/10/2026. Workbook Excel hiện có chứa 999 hồ sơ và 20 cột; cấu hình Code.gs (FILTER_START_DATE = 2026-08-01) chỉ nhận hồ sơ từ ngày 01/08/2026. Các kết quả trong chương là số liệu của lô được báo cáo, không đại diện toàn bộ workbook.")
    elif p.text.startswith("Báo cáo đề tài đã đạt được 5 kết quả định lượng cốt lõi sau:"):
        set_paragraph(p, "Báo cáo đề tài đã đạt được 5 kết quả định lượng cốt lõi sau:\\n1. Xây dựng hệ thống GGT-CNTT Auto trên Google Workspace.\\n2. Lô vận hành được báo cáo gồm 361 hồ sơ: 330 PDF (91,41%), 31 hồ sơ bị chặn (8,59%); không đại diện toàn bộ 999 hồ sơ trong workbook.\\n3. Thời gian máy chạy của lô là 82,7 phút (13,7 giây/hồ sơ), so với mốc thủ công ước tính 5 phút/hồ sơ.\\n4. Bộ kiểm thử tự động đạt 93/93 ca.\\n5. Hệ thống có hàng đợi checkpoint và Dashboard; tính năng AI tùy chọn đang tắt theo cấu hình hiện tại.")
    elif p.text.startswith("Nhờ cơ chế hàng đợi nền trình bày ở Mục 4.7.2, lô 361 hồ sơ"):
        set_paragraph(p, p.text.replace("lô 361 hồ sơ được tổng hợp được xử lý", "lô 361 hồ sơ trong bản tổng hợp vận hành được xử lý"))
    elif p.text.startswith("Hàm refreshDashboard()"):
        set_paragraph(p, p.text.replace("Tại thời điểm kết thúc đợt chạy, Dashboard ghi nhận: tổng 361 hồ sơ", "Theo bản tổng hợp đợt chạy, Dashboard ghi nhận lô 361 hồ sơ"))

# Sweep every paragraph, table cell and text box represented in table cells for stale lab names.
def clean_old_names(text):
    text = text.replace("Phòng Thí nghiệm Trí tuệ Nhân tạo (Artificial Intelligence Laboratory - AI Lab)", COMPANY)
    text = text.replace("Phòng Thí nghiệm Trí tuệ Nhân tạo (AI Lab)", COMPANY)
    text = text.replace("Phòng Thí nghiệm AI Lab - Khoa CNTT", COMPANY)
    text = text.replace("Phòng Thí nghiệm AI Lab", COMPANY)
    text = text.replace("AI Lab - Khoa CNTT", COMPANY)
    text = text.replace("AI Lab", COMPANY)
    text = text.replace("Phòng Thí nghiệm", COMPANY)
    return text

for p in doc.paragraphs:
    current = p.text
    updated = clean_old_names(current)
    if updated != current:
        set_paragraph(p, updated)
for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            current = cell.text
            updated = clean_old_names(current)
            if updated != current:
                set_cell(cell, updated)
            if "361/361 hồ sơ" in cell.text:
                set_cell(cell, cell.text.replace("361/361 hồ sơ", "361 hồ sơ trong lô vận hành được báo cáo"))

# Explicitly fix chapter 2 history, position and task description if they still carry lab claims.
for p in doc.paragraphs:
    if p.text.startswith("• Vị trí thực tập:"):
        text = p.text.replace("Thực tập sinh Kỹ thuật Phần mềm (Software Engineer Intern)",
                              f"Thực tập sinh Kỹ thuật Phần mềm tại {COMPANY} (Software Engineer Intern)")
        text = text.replace("Giảng viên / Nghiên cứu viên AI Lab - Khoa CNTT", f"Cán bộ hướng dẫn tại {COMPANY}")
        if text != p.text:
            set_paragraph(p, text)

doc.save(PATH)
print(f"Updated {PATH.name}")
