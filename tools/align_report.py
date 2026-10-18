#!/usr/bin/env python3
"""Reconcile repeated report claims with the supplied workbook and current code."""
from pathlib import Path
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parents[1]
PATH = ROOT / "BaoCao_KTCN_GGT_CNTT_Auto_522H0045.docx"
doc = Document(PATH)


def set_text(p, text):
    if p.runs:
        p.runs[0].text = text
        for run in p.runs[1:]:
            run.text = ""
        p.runs[0].font.name = "Times New Roman"
        p.runs[0]._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "Times New Roman")
    else:
        run = p.add_run(text)
        run.font.name = "Times New Roman"
        run.font.size = Pt(13)


def replace_start(prefix, replacement):
    found = False
    for p in doc.paragraphs:
        if p.text.strip().startswith(prefix):
            set_text(p, replacement)
            found = True
    return found

# Existing summary paragraph was edited by the report generator; rewrite in-place.
replace_start("Theo bản tổng hợp vận hành ngày 17–18/10/2026, hệ thống đã xử lý 361 hồ sơ:",
    "Bản tổng hợp vận hành ngày 17–18/10/2026 ghi nhận một lô 361 hồ sơ: 330 PDF được tạo "
    "(91,41%), 31 hồ sơ bị chặn (8,59%), gồm 5 trường hợp MST hợp lệ nhưng tên doanh nghiệp khác "
    "đăng ký thuế. Bộ kiểm thử tự động đạt 93/93. Tổng thời gian máy chạy của lô được báo cáo là "
    "82,7 phút (13,7 giây/hồ sơ), tiết kiệm khoảng 95,4% so với thời gian thủ công ước tính. Đây là "
    "số liệu của lô vận hành, không phải khẳng định toàn bộ dữ liệu workbook đã được xử lý.")
replace_start("The issuance of internship recommendation letters for students",
    "The issuance of internship recommendation letters for students of the Faculty of Information "
    "Technology still involves manual data entry and document preparation. This report describes "
    "GGT-CNTT Auto, a Google Apps Script system that normalizes form data, validates tax IDs, matches "
    "company names, generates single-page PDFs, and records results in Google Sheets. The supplied "
    "operation summary for 17–18 October 2026 reports one run of 361 submissions: 330 PDFs generated "
    "and 31 records blocked; the attached workbook contains 999 records in 20 columns, and the current "
    "Code.gs configuration filters submissions from 1 August 2026. The 361-record run is therefore "
    "reported as an observed batch, not as the entire workbook. The automated suite reports 93 passing "
    "cases. Gemini review is optional and disabled by default; bulk processing uses a checkpointed "
    "queue to fit Apps Script execution limits.")
replace_start("Theo số liệu thực tế lưu trong bảng tính của Khoa Công nghệ Thông tin",
    "Tệp Excel hiện có của bảng tính Khoa CNTT chứa 999 hồ sơ và 20 cột. Code.gs cấu hình "
    "FILTER_START_DATE = 2026-08-01 nên chỉ nhận xử lý hồ sơ từ ngày 01/08/2026. Kết quả 361 hồ sơ "
    "(330 PDF, 31 hồ sơ bị chặn) được trình bày là bản tổng hợp của lô vận hành ngày 17–18/10/2026, "
    "không phải toàn bộ dữ liệu trong workbook.")
replace_start("Từ kết quả thực nghiệm trên 361 hồ sơ thật,",
    "Từ kết quả thực nghiệm trên lô 361 hồ sơ trong bản tổng hợp vận hành,")
replace_start("Nhờ cơ chế hàng đợi nền trình bày ở Mục 4.7.2, toàn bộ 361 hồ sơ",
    "Nhờ cơ chế hàng đợi nền trình bày ở Mục 4.7.2, lô 361 hồ sơ trong bản tổng hợp vận hành")
replace_start("Báo cáo đề tài đã đạt được 5 kết quả định lượng cốt lõi sau",
    "Báo cáo đề tài đã đạt được 5 kết quả định lượng cốt lõi sau:\\n"
    "1. Xây dựng hệ thống GGT-CNTT Auto vận hành trên Google Workspace.\\n"
    "2. Bản tổng hợp vận hành ghi nhận lô 361 hồ sơ: 330 PDF được tạo (91,41%), 31 hồ sơ bị chặn (8,59%); không đại diện toàn bộ 999 hồ sơ trong workbook.\\n"
    "3. Thời gian máy chạy của lô là 82,7 phút (13,7 giây/hồ sơ), so với mốc thủ công ước tính 5 phút/hồ sơ.\\n"
    "4. Bộ kiểm thử tự động đạt 93/93 ca.\\n"
    "5. Hệ thống có menu quản trị, hàng đợi checkpoint và Dashboard; các tính năng có cấu hình bật/tắt.")
replace_start("Vận hành thực tế trên toàn bộ 361 hồ sơ trong Sheet của Khoa",
    "Xử lý lô 361 hồ sơ trong bản tổng hợp vận hành (330 PDF, 31 hồ sơ bị chặn); đây không phải "
    "toàn bộ 999 hồ sơ trong workbook.")
replace_start("Hệ thống hoàn thành được mục tiêu cốt lõi:",
    "Trên lô dữ liệu được báo cáo, hệ thống hoàn thành mục tiêu cốt lõi:")
replace_start("Tỷ lệ cấp giấy tự động đạt 91,41% (330/361).",
    "Tỷ lệ cấp giấy tự động trong lô được báo cáo đạt 91,41% (330/361).")
replace_start("Hạn chế của số liệu:",
    "Hạn chế của số liệu: workbook đính kèm có 999 hồ sơ, trong khi số liệu 361 hồ sơ là bản tổng "
    "hợp một lô vận hành; mốc thời gian thủ công 5 phút/hồ sơ là ước lượng từ khảo sát, không phải đo trực tiếp.")

# Qualify table claims that previously implied all rows were processed.
for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            if "361/361 hồ sơ" in cell.text:
                cell.text = cell.text.replace("361/361 hồ sơ", "361 hồ sơ trong lô vận hành được báo cáo")
            if "tổng 361 hồ sơ" in cell.text:
                cell.text = cell.text.replace("tổng 361 hồ sơ", "lô vận hành báo cáo: 361 hồ sơ")

doc.save(PATH)
print(PATH.name)
