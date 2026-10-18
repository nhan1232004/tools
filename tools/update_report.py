#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Cập nhật / viết lại báo cáo KTCN cho khớp với Code.gs mới và số liệu thực tế.

- Giữ nguyên khung template của Khoa (khổ A4, lề, font Times New Roman, heading styles).
- Viết lại toàn bộ Chương 5 theo số liệu thật (361 hồ sơ: 330 PDF, 31 lỗi).
- Bổ sung mục hàng đợi xử lý nền (4.7.2), cập nhật 4.5.3/4.8/4.9 cho đúng cấu hình mới.
- Cập nhật tóm tắt, mục tiêu, kết luận, bảng đối chiếu mục tiêu, phụ lục, danh mục hình/bảng.
- Xóa trang "PHỤ LỤC HƯỚNG DẪN CHO TÁC GIẢ".

Đầu ra: BaoCao_KTCN_GGT_CNTT_Auto_522H0045.docx
"""
import copy
import datetime
import os

import docx
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from docx.table import Table
from docx.text.paragraph import Paragraph

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "BaoCao_KTCN_KhungSuon.docx")
OUT = os.path.join(ROOT, "BaoCao_KTCN_GGT_CNTT_Auto_522H0045.docx")
FIG = os.path.join(ROOT, "figures")

W = qn("w:")[2:]  # namespace prefix helper not used directly


def w(tag):
    return "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}" + tag


# --------------------------------------------------------------------------- #
#  Tiện ích thao tác trên cây XML của document
# --------------------------------------------------------------------------- #
def body_items(doc):
    out = []
    for child in doc.element.body.iterchildren():
        if child.tag == w("p"):
            out.append(("p", Paragraph(child, doc)))
        elif child.tag == w("tbl"):
            out.append(("t", Table(child, doc)))
    return out


def strip_runs(p_el):
    for child in list(p_el):
        if child.tag in (w("r"), w("hyperlink"), w("bookmarkStart"), w("bookmarkEnd"),
                         w("proofErr")):
            p_el.remove(child)


def make_par(template_p, text, bold=False, italic=False, no_indent=False):
    """Tạo đoạn văn mới từ pPr của template (giữ nguyên style/lề/giãn dòng)."""
    el = copy.deepcopy(template_p._p)
    strip_runs(el)
    p = Paragraph(el, template_p._parent)
    if text:
        run = p.add_run(text)
        set_run_font(run, bold=bold, italic=italic)
    if no_indent and el.find(w("pPr")) is not None:
        ind = el.find(w("pPr")).find(w("ind"))
        if ind is not None:
            ind.set(w("firstLine"), "0")
    return el


def set_run_font(run, size=13, bold=False, italic=False, name="Times New Roman", color=None):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = color
    rpr = run._r.get_or_add_rPr()
    rfonts = rpr.find(w("rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts"); rpr.insert(0, rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), name)
    return run


def set_par_text(p, text, bold=False, italic=False):
    """Thay văn bản của một đoạn sẵn có, giữ định dạng của run đầu tiên."""
    if p.runs:
        first = p.runs[0]
        first.text = text
        for r in p.runs[1:]:
            r.text = ""
        if bold:
            first.font.bold = True
        if italic:
            first.font.italic = True
    else:
        set_run_font(p.add_run(text), bold=bold, italic=italic)


def find_par(doc, startswith, nth=0):
    hits = [p for p in doc.paragraphs if p.text.strip().startswith(startswith)]
    if not hits or nth >= len(hits):
        return None
    return hits[nth]


def find_body_index(items, pred):
    for i, (k, o) in enumerate(items):
        if k == "p" and pred(o.text.strip()):
            return i
    return -1


# --------------------------------------------------------------------------- #
#  Bảng và hình
# --------------------------------------------------------------------------- #
def style_table(tbl, widths):
    """Đặt khổ bảng theo chuẩn dxa và bảo đảm thứ tự phần tử hợp lệ trong tblPr:
    tblW → tblBorders → tblLayout → tblLook."""
    tblPr = tbl._tbl.tblPr
    # gỡ các phần tử sẽ tạo lại để tránh trùng lặp/sai thứ tự
    look = tblPr.find(w("tblLook"))
    if look is not None:
        tblPr.remove(look)
    for tag in ("tblBorders", "tblLayout"):
        el = tblPr.find(w(tag))
        if el is not None:
            tblPr.remove(el)

    tblW = tblPr.find(w("tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW"); tblPr.insert(0, tblW)
    tblW.set(w("w"), str(sum(widths))); tblW.set(w("type"), "dxa")

    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement("w:" + edge)
        e.set(w("val"), "single"); e.set(w("color"), "auto"); e.set(w("sz"), "4")
        e.set(w("space"), "0")
        borders.append(e)
    tblPr.append(borders)
    layout = OxmlElement("w:tblLayout")
    layout.set(w("type"), "fixed")
    tblPr.append(layout)
    if look is not None:
        tblPr.append(look)
    grid = tbl._tbl.find(w("tblGrid"))
    for gc, width in zip(grid.findall(w("gridCol")), widths):
        gc.set(w("w"), str(width))


def fill_cell(cell, text, bold=False, size=10.5, center=False, fill=None):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
    pf = p.paragraph_format
    pf.space_before = Pt(1); pf.space_after = Pt(1)
    pf.line_spacing = 1.15
    if text:
        set_run_font(p.add_run(text), size=size, bold=bold)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    if fill:
        shd = OxmlElement("w:shd")
        shd.set(w("val"), "clear"); shd.set(w("color"), "auto"); shd.set(w("fill"), fill)
        tcPr = cell._tc.get_or_add_tcPr()
        va = tcPr.find(w("vAlign"))
        if va is not None:
            va.addprevious(shd)
        else:
            tcPr.append(shd)


def make_table(doc, rows_data, widths, header_fill="E7E6E6", size=10.5):
    """Tạo bảng mới (chưa gắn vào tài liệu)."""
    n_rows, n_cols = len(rows_data), len(rows_data[0])
    tbl = doc.add_table(rows=n_rows, cols=n_cols)
    for r, row in enumerate(rows_data):
        for c, val in enumerate(row):
            fill_cell(tbl.cell(r, c), val, bold=(r == 0), size=size,
                      center=(r == 0), fill=header_fill if r == 0 else None)
    style_table(tbl, widths)
    el = tbl._tbl
    el.getparent().remove(el)
    return el


def make_figure(doc, image_path, width_cm=15.2):
    """Tạo khung 1x1 chứa hình (giống các figure container có sẵn)."""
    tbl = doc.add_table(rows=1, cols=1)
    style_table(tbl, [8787])
    cell = tbl.cell(0, 0)
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(2); p.paragraph_format.space_after = Pt(2)
    run = p.add_run()
    run.add_picture(image_path, width=Cm(width_cm))
    el = tbl._tbl
    el.getparent().remove(el)
    return el


# --------------------------------------------------------------------------- #
#  Bắt đầu
# --------------------------------------------------------------------------- #
doc = docx.Document(SRC)

# ---- 1. Lấy template định dạng trước khi xóa/sửa ---------------------------- #
TPL = {}
paras = doc.paragraphs
TPL["body"] = paras[88]                       # đoạn thân, justified, thụt đầu dòng
TPL["bullet"] = paras[96]                     # bullet List Paragraph
TPL["caption"] = paras[371]                   # caption bảng/hình, căn giữa
TPL["code"] = paras[289]                      # khối mã
TPL["h2"] = find_par(doc, "5.4 Thực nghiệm")
TPL["h3"] = find_par(doc, "5.4.1 Bộ ca kiểm thử")
TPL["h2"] = TPL["h2"] or paras[374]
TPL["h3"] = TPL["h3"] or paras[392]
LIST_FIG_TPL = None  # lấy sau

FIG_TPL = copy.deepcopy(doc.tables[11]._tbl)  # khung hình 1x1 rỗng (chỉ giữ cấu trúc)


def fig_block(doc, ref_el, image, caption, width_cm=15.2):
    """Chèn [khung chứa hình] + [caption] ngay trước ref_el."""
    tbl = copy.deepcopy(FIG_TPL)
    # xóa nội dung cũ trong ô, giữ tcPr
    tc = tbl.findall(w("tr"))[0].findall(w("tc"))[0]
    for child in list(tc):
        if child.tag == w("p"):
            tc.remove(child)
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    jc = OxmlElement("w:jc"); jc.set(w("val"), "center")
    pPr.append(jc); p.append(pPr)
    tc.append(p)

    tmp = doc.add_paragraph()
    run = tmp.add_run()
    run.add_picture(os.path.join(FIG, image), width=Cm(width_cm))
    drawing_r = copy.deepcopy(run._r)
    doc.element.body.remove(tmp._p)
    p.append(drawing_r)

    ref_el.addprevious(tbl)
    cap_el = make_par(TPL["caption"], caption)
    ref_el.addprevious(cap_el)


def insert_pars(doc, ref_el, specs):
    """specs: list các tuple, kind in {'body','bullet','h2','h3','caption','code','gap',
    'caption_table','table'}. Với 'table': (kind, rows_data, widths)."""
    for spec in specs:
        kind = spec[0]
        if kind == "table":
            rows_data, widths = spec[1], spec[2]
            ref_el.addprevious(make_table(doc, rows_data, widths))
            continue
        text = spec[1]
        if kind == "gap":
            ref_el.addprevious(make_par(TPL["body"], "", no_indent=True))
            continue
        if kind == "caption_table":
            ref_el.addprevious(make_par(TPL["caption"], text))
            continue
        tpl = TPL[kind]
        el = make_par(tpl, text)
        if kind == "bullet":
            # gắn numbering bullet giống các mục sẵn có trong tài liệu
            pPr = el.find(w("pPr"))
            if pPr.find(w("numPr")) is None:
                numPr = OxmlElement("w:numPr")
                ilvl = OxmlElement("w:ilvl"); ilvl.set(w("val"), "0")
                numId = OxmlElement("w:numId"); numId.set(w("val"), "2")
                numPr.append(ilvl); numPr.append(numId)
                pPr.insert(1, numPr)
        ref_el.addprevious(el)


# --------------------------------------------------------------------------- #
#  2. Cập nhật phần đầu báo cáo
# --------------------------------------------------------------------------- #
def replace_par(doc, startswith, new_text, bold=False, nth=0):
    p = find_par(doc, startswith, nth)
    if p is None:
        print("  [cảnh báo] không tìm thấy đoạn:", startswith[:60])
        return False
    set_par_text(p, new_text, bold=bold)
    return True


# --- Tóm tắt tiếng Việt ---
replace_par(doc, "Hệ thống đã được kiểm nghiệm thực tế trên 92 hồ sơ",
            "Hệ thống đã được vận hành thực tế trên toàn bộ 361 hồ sơ đăng ký cấp Giấy giới thiệu "
            "đang lưu trong Google Sheet của Khoa CNTT. Kết quả: 330 hồ sơ được cấp Giấy giới thiệu "
            "PDF tự động (91,41%) và 31 hồ sơ bị hệ thống chặn kèm nguyên nhân cụ thể (8,59%) — "
            "trong đó có 5 trường hợp mã số thuế đúng nhưng tên doanh nghiệp do sinh viên khai khác "
            "với tên đăng ký của cơ quan Thuế, nếu cấp giấy sẽ sai đơn vị tiếp nhận. Bộ kiểm thử tự "
            "động gồm 93 ca đều đạt. Tổng thời gian máy chạy cho cả 361 hồ sơ là 82,7 phút "
            "(tương đương 13,7 giây/hồ sơ), tiết kiệm khoảng 95,4% khối lượng thời gian so với cách "
            "làm thủ công; toàn bộ 330 tệp PDF được lưu tập trung trên Google Drive và liên kết mở "
            "tức thì đã được ghi trực tiếp vào cột “LINK PDF GGT” của Sheet.")
replace_par(doc, "Hệ thống nhận dữ liệu từ Google Form, chuẩn hóa thông tin",
            "Hệ thống nhận dữ liệu từ Google Form, tự động chuẩn hóa các giá trị bị bảng tính biến "
            "dạng (MSSV, số điện thoại, mã số thuế dạng số khoa học hoặc mất số 0 đầu), kiểm chứng "
            "mã số thuế qua API tra cứu của cơ quan Thuế, đối chiếu tên doanh nghiệp bằng thuật toán "
            "so khớp từ khóa có trọng số (weighted Dice) với các nhánh xử lý riêng cho chi nhánh, "
            "tên viết tắt và tên tiếng nước ngoài; sau đó tự động điền văn bản mẫu, xuất tệp PDF "
            "đúng một trang (bảo toàn khoảng trống ký tên) và ghi kết quả trở lại Google Sheet. "
            "Mô-đun AI (Gemini) được thiết kế sẵn cho vùng xám nhưng mặc định đang tắt bằng cấu "
            "hình AI_ENABLED = false để bảo đảm mọi quyết định cấp giấy đều xuất phát từ luật "
            "kiểm chứng được.")
replace_par(doc, "Từ khóa: số hóa, tự động hóa quy trình",
            "Từ khóa: số hóa, tự động hóa quy trình, văn bản hành chính, Google Apps Script, "
            "mã số thuế, xử lý hàng đợi.")

# --- Abstract tiếng Anh ---
replace_par(doc, "The issuance of internship recommendation letters",
            "The issuance of internship recommendation letters for students of the Faculty of "
            "Information Technology still involves numerous manual steps: students submit information "
            "through Google Forms, and academic officers must copy data into document templates, verify "
            "corporate tax identification numbers (TIN), format margins, export single-page PDFs and "
            "upload them to shared storage. This report presents the design, deployment and evaluation "
            "of GGT-CNTT Auto — a serverless automation system running entirely on Google Workspace with "
            "Google Apps Script. The system ingests form responses stored in Google Sheets, repairs "
            "spreadsheet-induced data corruption (scientific notation and lost leading zeros), validates "
            "tax codes against the national tax-payer API, and matches declared company names with "
            "official registry names using a weighted token-similarity algorithm that handles branches, "
            "abbreviations and foreign-language trade names. It then populates a Google Docs template, "
            "guarantees a clean single-page PDF while preserving the signature area, stores the output in "
            "Google Drive and writes the status, tax-verification result and direct URL back to the sheet. "
            "An optional Gemini-based review layer for the grey similarity zone is implemented but "
            "disabled by default (AI_ENABLED = false). Because Apps Script executions are limited to six "
            "minutes, bulk runs are handled by a checkpointed background queue "
            "(25 rows / 240 s per run, resumed by a one-minute trigger). The system was executed on the "
            "faculty's real dataset of 361 submissions: it issued 330 letters automatically (91.41%) and "
            "blocked 31 invalid submissions with explicit reasons (8.59%), including five cases in which a "
            "valid tax code was paired with a different company name. An automated test suite of 93 cases "
            "passes completely. Total machine time for all 361 records was 82.7 minutes (13.7 seconds per "
            "record), a 95.4% reduction compared with the estimated manual workflow.")

# --- Mục tiêu cụ thể: nói rõ AI là tuỳ chọn ---
replace_par(doc, "Tích hợp mô hình ngôn ngữ lớn để hỗ trợ đối soát tên doanh nghiệp",
            "Thiết kế sẵn lớp hỗ trợ bằng mô hình ngôn ngữ lớn cho các trường hợp đối soát tên không "
            "chắc chắn (vùng xám), có công tắc bật/tắt và ngưỡng tin cậy rõ ràng; bảo đảm khi tắt lớp "
            "này hệ thống vẫn hoạt động đầy đủ bằng luật.")
replace_par(doc, "Phạm vi: hệ thống chạy trong Google Apps Script",
            "Phạm vi: hệ thống chạy trong Google Apps Script gắn với Google Sheet; chỉ xử lý loại văn "
            "bản giấy giới thiệu thực tập; đối soát do luật có trọng số quyết định, mô-đun AI chỉ là "
            "lớp mở rộng tùy chọn (mặc định tắt) và chỉ nhận tên doanh nghiệp, không nhận dữ liệu cá "
            "nhân của sinh viên.")
replace_par(doc, "Theo số liệu khảo sát thực tế tại Văn phòng Khoa Công nghệ Thông tin",
            "Theo số liệu thực tế lưu trong bảng tính của Khoa Công nghệ Thông tin - Trường Đại học "
            "Tôn Đức Thắng, tính đến tháng 10/2026 có 361 hồ sơ đăng ký cấp giấy giới thiệu thực tập "
            "của sinh viên các học phần Kiến tập công nghiệp, Tập sự nghề nghiệp và Dự án CNTT 1; "
            "trong các học kỳ cao điểm số lượng đăng ký có thể lên tới hàng trăm hồ sơ. Với cách làm "
            "thủ công, mỗi hồ sơ cần khoảng 5 phút thao tác sao chép, kiểm tra và định dạng, tương "
            "đương khoảng 30 giờ làm việc cho toàn bộ số hồ sơ nêu trên.")

# --- Bảng 6.1: đối chiếu mục tiêu ---
t61 = None
for t in doc.tables:
    if len(t.rows) == 7 and t.rows[0].cells[0].text.strip().startswith("Mục tiêu"):
        t61 = t
        break
if t61 is not None:
    mapping = {
        "Tích hợp AI hỗ trợ ca không chắc chắn":
            "Hoàn thành thiết kế, mặc định tắt",
        "Luồng tự động từ Form đến PDF và Sheet":
            "Hoàn thành (361/361 hồ sơ)",
        "Kiểm chứng MST và đối soát tên":
            "Hoàn thành (330 giấy cấp, 31 ca chặn)",
        "Kiểm thử và đo hiệu quả":
            "Hoàn thành (93 ca đạt; 13,7 giây/hồ sơ)",
        "Quy trình duyệt, cấp số, nhật ký, thống kê":
            "Hoàn thành (có công tắc cấu hình)",
    }
    for row in t61.rows[1:]:
        key = row.cells[0].text.strip()
        for k, v in mapping.items():
            if k.lower() in key.lower():
                fill_cell(row.cells[1], v, size=10.5)
        ev = row.cells[2].text
        if "5.4" in ev or "5.7" in ev:
            fill_cell(row.cells[2], ev.replace("5.4, 5.5", "5.4–5.6").replace("4.5.3, 5.4.2", "4.5.3, 5.4"),
                      size=10.5)

print("Đã cập nhật phần đầu báo cáo.")

# --------------------------------------------------------------------------- #
#  3. Cập nhật Chương 4: AI tuỳ chọn, hàng đợi nền, mô-đun, bảo mật
# --------------------------------------------------------------------------- #
# 4.5.3 lớp AI
replace_par(doc, "Hệ thống hỗ trợ kiến trúc mở rộng sẵn sàng tích hợp với các mô hình ngôn ngữ lớn",
            "Hệ thống có sẵn một lớp mở rộng gọi mô hình Gemini qua UrlFetchApp để xử lý các trường hợp "
            "đối soát tên không chắc chắn. Lớp này chỉ được kích hoạt khi cấu hình AI_ENABLED = true; "
            "trong bản triển khai chính thức tại Khoa, tham số này được đặt bằng false nên toàn bộ kết "
            "quả trong Chương 5 đều do luật đối soát quyết định. Cách bố trí này giữ cho hệ thống có "
            "thể giải thích được (explainable) và không phụ thuộc vào dịch vụ bên ngoài.")
replace_par(doc, "Thuật toán luật có điểm yếu với các biến thể chưa lường trước",
            "Vùng can thiệp của AI được giới hạn chặt: chỉ những hồ sơ có điểm tương đồng Dice nằm "
            "trong khoảng [AI_DICE_MIN, AI_DICE_MAX] = [0,30; 0,50] mới được gửi sang mô hình; kết quả "
            "chỉ được chấp nhận khi mô hình trả về match = true với độ tin cậy tối thiểu "
            "AI_CONFIDENCE_THRESHOLD = 0,90. Mọi trường hợp mô hình trả lời không đủ ngưỡng đều được "
            "ghi nhận là cần người kiểm tra, không tự động cấp giấy. Dữ liệu gửi đi chỉ gồm tên doanh "
            "nghiệp sinh viên khai, tên đăng ký thuế, tên nước ngoài và tên viết tắt; họ tên, MSSV, "
            "email và số điện thoại của sinh viên không được gửi. Nếu lời gọi AI thất bại, hệ thống "
            "ghi log và quay về kết quả của luật, không làm gián đoạn luồng xử lý.")

# mã giả 4.2 – cập nhật theo cấu hình mới (xóa cả khối pseudocode cũ)
items4 = body_items(doc)
i_pc = find_body_index(items4, lambda t: t.startswith("// Pseudocode điểm tích hợp"))
if i_pc != -1:
    lines = [
        "// Đoạn trích compareCompanyNames_() — nhánh vùng xám (đang tắt trong bản triển khai)",
        "if (CONFIG.AI_ENABLED &&",
        "    directSim.dice >= CONFIG.AI_DICE_MIN &&",
        "    directSim.dice <= CONFIG.AI_DICE_MAX) {",
        "  const r = judgeCompanyNames_(studentName, apiName, foreignName, abbrev);",
        "  if (r.match === true && r.confidence >= CONFIG.AI_CONFIDENCE_THRESHOLD)",
        "    return { isMatch: true,",
        "             reason: 'AI hỗ trợ đối soát vùng xám (' + Math.round(r.confidence*100) + '%)' };",
        "  return { isMatch: false, reason: 'AI không xác nhận đủ ngưỡng tin cậy' };",
        "}",
        "return { isMatch: false, reason: 'Không tương đồng (điểm số: ' +",
        "         Math.round(directSim.dice * 100) + '%)' };",
    ]
    anchor = items4[i_pc][1]._p
    to_remove, started = [], False
    for k, o in items4[i_pc:]:
        el = o._p if k == "p" else o._tbl
        text = o.text.strip() if k == "p" else ""
        if text.startswith("Đoạn mã 4.2"):
            break
        to_remove.append(el)
    parent = anchor.getparent()
    idx = list(parent).index(anchor)
    for el in to_remove:
        parent.remove(el)
    for i, line in enumerate(lines):
        parent.insert(idx + i, make_par(TPL["code"], line))
    print(f"Đã thay khối mã giả 4.2 ({len(to_remove)} đoạn cũ).")

def replace_par_last(doc, startswith, new_text):
    """Đổi nội dung lần xuất hiện CUỐI (bỏ qua danh mục hình vẽ ở đầu tài liệu)."""
    hits = [p for p in doc.paragraphs if p.text.strip().startswith(startswith)]
    if not hits:
        print("  [cảnh báo] không tìm thấy đoạn:", startswith[:60])
        return False
    set_par_text(hits[-1], new_text)
    return True


# Đánh số lại 2 hình Chương 2 theo thứ tự xuất hiện: sơ đồ tổ chức (mục 2.2) in trước
# quy trình thủ công (mục 2.4) nên phải là Hình 2.1, và ngược lại.
_h2 = [p for p in doc.paragraphs if p.text.strip().startswith("Hình 2.")]
_org = [p for p in _h2 if "Sơ đồ cơ cấu" in p.text]
_proc = [p for p in _h2 if "Quy trình cấp giấy giới thiệu" in p.text]
if _org and _proc:
    set_par_text(_org[-1], "Hình 2.1: Sơ đồ cơ cấu tổ chức và vị trí thực tập tại "
                         "Phòng Thí nghiệm AI Lab")
    set_par_text(_proc[-1], "Hình 2.2: Quy trình cấp giấy giới thiệu thực tập trước khi "
                            "tự động hóa")
    print("Đã đánh số lại Hình 2.1 / Hình 2.2 theo thứ tự xuất hiện.")

# Đánh số lại hình Chương 4 theo đúng thứ tự xuất hiện:
# đối soát tên (mục 4.5) = Hình 4.3, máy trạng thái (mục 4.7) = Hình 4.4
replace_par_last(doc, "Hình 4.4: Lưu đồ đối soát tên doanh nghiệp",
                 "Hình 4.3: Lưu đồ đối soát tên doanh nghiệp")
replace_par_last(doc, "Hình 4.3: Máy trạng thái của hồ sơ",
                 "Hình 4.4: Máy trạng thái của hồ sơ")
replace_par_last(doc, "Hình 4.1: Kiến trúc tổng thể của hệ thống",
                 "Hình 4.1: Kiến trúc tổng thể của hệ thống (bản cập nhật theo Code.gs mới)")

# 4.7 bổ sung tiểu mục 4.7.1 + mục hàng đợi trước caption Hình 4.4
p47 = find_par(doc, "4.7 Quản lý trạng thái")
if p47 is not None:
    insert_pars(doc, p47._p.getnext(),
                [("h3", "4.7.1 Trạng thái hồ sơ và menu quản trị")])
    print("Đã bổ sung tiểu mục 4.7.1.")

items = body_items(doc)
# lấy lần xuất hiện CUỐI (bỏ qua danh mục hình vẽ ở đầu tài liệu)
_fig43 = [i for i, (k, o) in enumerate(items)
          if k == "p" and o.text.strip().startswith("Hình 4.4: Máy trạng thái")]
if not _fig43:
    _fig43 = [i for i, (k, o) in enumerate(items)
              if k == "p" and o.text.strip().startswith("Hình 4.1: Kiến trúc tổng thể")]
    _fig43 = [max(_fig43) + 2] if _fig43 else []
i_fig43 = _fig43[-1] if _fig43 else -1
print("  [debug] vị trí chèn 4.7.2:", i_fig43)
# chèn SAU caption của Hình 4.3 (tức trước phần tử kế tiếp) để không chen vào giữa hình và caption
if i_fig43 != -1 and i_fig43 + 1 < len(items):
    ref = items[i_fig43 + 1][1]._p
else:
    ref = doc.paragraphs[-1]._p
insert_pars(doc, ref, [
    ("h3", "4.7.2 Xử lý hàng loạt bằng hàng đợi nền (batch queue)"),
    ("body", "Google Apps Script giới hạn mỗi lần thực thi (execution) tối đa 6 phút. Với hàng trăm "
             "hồ sơ cần in trong cao điểm, việc xử lý tất cả trong một lần chạy sẽ bị dừng giữa "
             "chừng và để lại dữ liệu nửa vời. Hệ thống giải quyết bằng cơ chế hàng đợi nền có "
             "checkpoint: hàm enqueueBatch_() lưu danh sách dòng cần xử lý vào thuộc tính "
             "GGT_BATCH_QUEUE của Script Properties, tạo một trigger thời gian chạy mỗi phút và "
             "khởi động ngay đợt đầu tiên."),
    ("body", "Hàm continueBatchQueue_() lấy khóa script để bảo đảm chỉ một đợt chạy tại một thời "
             "điểm, sau đó xử lý lần lượt các dòng: tối đa MAX_BATCH_ROWS = 25 dòng trong một đợt và "
             "dừng trước khi chạm mốc BATCH_RUNTIME_BUDGET_MS = 240.000 ms (4 phút), tạo khoảng an "
             "toàn so với giới hạn 6 phút của nền tảng. Sau mỗi dòng, tiến độ (index, success, "
             "failed) được ghi lại ngay vào Script Properties, nhờ đó nếu execution bị dừng bất ngờ "
             "thì lần chạy kế tiếp tiếp tục đúng dòng còn dang dở thay vì làm lại từ đầu. Giữa hai "
             "hồ sơ liên tiếp, hệ thống nghỉ BATCH_SLEEP_MS = 3.200 ms để không vượt hạn mức "
             "10 yêu cầu/30 giây của API tra cứu mã số thuế."),
    ("body", "Khi danh sách đã xử lý hết, hệ thống xóa hàng đợi, hủy trigger tạm và hiển thị thông "
             "báo tổng kết (số thành công, số lỗi, tổng số dòng) trên giao diện Sheet. Người dùng có "
             "thể dừng phần còn lại bất cứ lúc nào bằng hàm cancelBatchQueue(); tiến độ đã checkpoint "
             "vẫn được giữ nguyên. Hình 4.5 minh họa toàn bộ vòng đời của một lô xử lý."),
    ("gap", ""),
])
fig_block(doc, ref, "hinh_4_5_hang_doi_xu_ly_nen.png",
          "Hình 4.5: Vòng đời của một lô xử lý nền trong hệ thống")
print("Đã bổ sung mục 4.7.2 và Hình 4.5.")

# 4.8 mô-đun mở rộng — cập nhật trạng thái
replace_par(doc, "Đặc quyền tối thiểu: script chỉ xin những quyền OAuth cần thiết",
            "Đặc quyền tối thiểu: script chỉ xin những quyền OAuth cần thiết; Sheet và thư mục "
            "Drive chỉ chia sẻ cho người liên quan; tệp PDF giữ ở chế độ RESTRICTED và chỉ cấp "
            "quyền xem cho tài khoản có quyền với thư mục, không mở công khai theo liên kết.")

menu_bullet = find_par(doc, "Menu quản trị")
if menu_bullet is not None:
    set_par_text(menu_bullet,
                 "Menu quản trị `GGT-CNTT`: kiểm thử một dòng, nhiều dòng (2, 5, 7-10), dòng cuối; "
                 "in PDF hàng loạt qua hàng đợi nền; xử lý lại các dòng lỗi; cập nhật Dashboard; "
                 "cài đặt trigger.")

mods = [
    ("Cấp số văn bản tự động: bộ đếm theo năm",
     "Cấp số văn bản tự động: bộ đếm theo năm lưu trong Script Properties, bảo vệ bằng khóa, định "
     "dạng 0001-GGT-CNTT/2026; đã cài đặt xong nhưng mặc định tắt (AUTO_NUMBERING = false) vì Khoa "
     "muốn giữ số và ngày ký cho văn phòng cấp."),
    ("Thông báo cho sinh viên: email xác nhận",
     "Thông báo qua email: hệ thống có sẵn mẫu thư gửi giảng viên kèm tệp PDF và bảng thông tin "
     "đối soát, nhưng trong đợt triển khai này SEND_EMAIL = false nên giảng viên mở trực tiếp liên "
     "kết trong cột “LINK PDF GGT”."),
    ("Bảng thống kê: sheet Dashboard",
     "Bảng thống kê: hàm refreshDashboard() tự tạo/cập nhật sheet “Dashboard GGT” bằng công thức "
     "COUNTA/COUNTIF theo trạng thái, kèm thời điểm cập nhật; cán bộ Khoa chỉ cần bấm một mục trong "
     "menu GGT-CNTT."),
]
dashboard_par = None
for start, new in mods:
    p = find_par(doc, start[:40])
    if p is not None:
        set_par_text(p, new)
        if start.startswith("Bảng thống kê"):
            dashboard_par = p
    else:
        print("  [cảnh báo] không tìm thấy mô-đun:", start[:40])

# bổ sung bullet “Quy trình duyệt” vào 4.8 (chưa có sẵn trong bản cũ)
if dashboard_par is not None:
    new_bullet = make_par(TPL["bullet"],
        "Quy trình duyệt: khi bật REQUIRE_APPROVAL = true, hồ sơ hợp lệ được ghi trạng thái “Chờ "
        "duyệt” và chỉ cấp giấy khi cán bộ Khoa chọn dòng rồi chạy approveSelectedRow(); mặc định "
        "trong đợt triển khai này được đặt bằng false.")
    pPr = new_bullet.find(w("pPr"))
    if pPr.find(w("numPr")) is None:
        numPr = OxmlElement("w:numPr")
        ilvl = OxmlElement("w:ilvl"); ilvl.set(w("val"), "0")
        numId = OxmlElement("w:numId"); numId.set(w("val"), "2")
        numPr.append(ilvl); numPr.append(numId)
        pPr.insert(1, numPr)
    dashboard_par._p.addnext(new_bullet)

# 4.9 bảo mật
sec = [
    ("PDF chia sẻ theo đường dẫn ở chế độ chỉ xem",
     "PDF mặc định ở chế độ RESTRICTED: chỉ tài khoản có quyền với thư mục Drive của Khoa mới xem "
     "được, không mở công khai theo liên kết. Chỉ khi đơn vị cho phép mới đổi "
     "PDF_SHARING_MODE sang ANYONE_WITH_LINK; cấu hình được kiểm tra hợp lệ trong checkConfig_() "
     "trước khi hệ thống chạy."),
    ("Với AI chỉ gửi tên doanh nghiệp",
     "Với lớp AI (khi bật) chỉ gửi tên doanh nghiệp, tên viết tắt và tên nước ngoài; không gửi họ "
     "tên, MSSV, email, số điện thoại của sinh viên."),
]
for start, new in sec:
    p = find_par(doc, start[:38])
    if p is not None:
        set_par_text(p, new)

print("Đã cập nhật Chương 4.")

# --------------------------------------------------------------------------- #
#  4. Viết lại toàn bộ Chương 5
# --------------------------------------------------------------------------- #
items = body_items(doc)
i5 = find_body_index(items, lambda t: t.startswith("CHƯƠNG 5."))
i6 = find_body_index(items, lambda t: t.startswith("CHƯƠNG 6."))
assert i5 != -1 and i6 != -1 and i6 > i5, "Không xác định được Chương 5/6"
ref6 = items[i6][1]._p
# xóa nội dung cũ của chương 5 (giữ lại tiêu đề chương)
for k, o in items[i5 + 1:i6]:
    el = o._p if k == "p" else o._tbl
    el.getparent().remove(el)
print(f"Đã xóa {i6 - i5 - 1} phần tử cũ của Chương 5.")

CH5 = [
    ("h2", "5.1 Môi trường triển khai"),
    ("body", "Hệ thống được triển khai trực tiếp trên không gian làm việc Google của Khoa Công nghệ "
             "thông tin, không cần máy chủ riêng và không phát sinh chi phí hạ tầng. Bảng 5.1 tổng "
             "hợp các thành phần của môi trường thực nghiệm."),
    ("caption_table", "Bảng 5.1: Môi trường triển khai và thực nghiệm"),
    ("table", [
        ["Thành phần", "Thông tin"],
        ["Nền tảng nghiệp vụ", "Google Form (thu thập), Google Sheet (dữ liệu và kết quả), "
                               "Google Docs (văn bản mẫu), Google Drive (lưu trữ PDF)"],
        ["Môi trường thực thi", "Google Apps Script (V8), script gắn với bảng tính; trigger "
                                "onFormSubmit cho từng lượt nộp và trigger thời gian 1 phút cho hàng đợi nền"],
        ["Bảng tính dữ liệu", "Sheet “Form Responses 1”: 20 cột, 361 dòng dữ liệu thực tế; hệ thống "
                              "tự nhận diện cột theo từ khóa"],
        ["Văn bản mẫu", "Google Doc chuyển từ GGT_mau_placeholder.docx, 6 placeholder bắt buộc và "
                        "8 placeholder mở rộng ({{SO_GGT}}, {{MST}}, {{DIA_CHI_CTY}}, {{NGAY}}…)"],
        ["Thư mục lưu PDF", "Google Drive, mã thư mục 1RgRcem14hvJdTa_Cq7As02RnYS8idFQ6, cấu hình "
                            "PDF_SHARING_MODE = RESTRICTED"],
        ["Dịch vụ ngoài", "API tra cứu người nộp thuế: https://api.xinvoice.vn/gdt-api/tax-payer/{mst}; "
                          "Gemini API cho lớp AI tùy chọn (đang tắt)"],
        ["Bộ kiểm thử", "Node.js v24.21.0 chạy tệp ggt-cntt-auto/sources/test_suite.js — 93 ca kiểm thử"],
        ["Mã nguồn", "GitHub nhan1232004/tools — Code.gs (1.977 dòng) và test_suite.js (485 dòng)"],
        ["Thời gian thực nghiệm", "Ngày 17–18/10/2026, gồm 4 đợt chạy liên tục"],
    ], [2300, 6487]),
    ("gap", ""),
    ("h2", "5.2 Quy trình cài đặt và khởi tạo"),
    ("body", "Toàn bộ hệ thống được cài đặt trong 4 bước, không cần thao tác nào ngoài giao diện "
             "Google Workspace:"),
    ("bullet", "Tải mẫu Word đã gắn placeholder lên Drive, chuyển thành Google Doc và ghi lại Doc ID;"),
    ("bullet", "Tạo thư mục nhận tệp PDF, ghi lại Folder ID và mã thư mục vào cấu hình;"),
    ("bullet", "Mở Sheet câu trả lời → Tiện ích mở rộng → Apps Script, dán mã nguồn Code.gs và chỉnh "
               "các tham số trong khối CONFIG (Hình 5.1);"),
    ("bullet", "Chạy hàm setup(): hệ thống kiểm tra Doc mẫu và thư mục Drive bằng checkConfig_(), "
               "kiểm tra/thêm các cột quản trị, rồi mới tạo trigger onFormSubmit."),
    ("body", "Thiết kế “kiểm tra trước – thay đổi sau” của hàm setup() giúp nếu cấu hình sai thì "
             "trigger cũ vẫn được giữ nguyên và người dùng nhận thông báo lỗi rõ ràng ngay trong "
             "Apps Script, tránh tình trạng hệ thống ngừng chạy mà không rõ nguyên nhân. Sau khi cài "
             "đặt, menu GGT-CNTT xuất hiện trên thanh công cụ của Sheet mỗi lần mở tệp (Hình 5.5)."),
    ("gap", ""),
]
insert_pars(doc, ref6, CH5)
# hình 5.1 và 5.2 (menu) đặt sau các đoạn tương ứng
fig_block(doc, ref6, "hinh_5_1_apps_script_config.png",
          "Hình 5.1: Cấu hình tập trung CONFIG của hệ thống (bản cập nhật 10/2026)")

CH5B = [
    ("h2", "5.3 Kết quả triển khai các chức năng"),
    ("body", "Các chức năng chính của hệ thống đều đã được triển khai và chạy thật trên dữ liệu của "
             "Khoa: biểu mẫu thu thập thông tin (Hình 5.2), bảng tính ghi nhận kết quả kèm các cột "
             "quản trị được tạo tự động (Hình 5.3), tệp Giấy giới thiệu PDF đúng một trang và bảo "
             "toàn khoảng trống ký tên (Hình 5.4), cùng menu quản trị GGT-CNTT dành cho cán bộ Khoa "
             "(Hình 5.5)."),
]
insert_pars(doc, ref6, CH5B)
fig_block(doc, ref6, "hinh_5_2_google_form.png",
          "Hình 5.2: Biểu mẫu Google Form thu thập thông tin cấp Giấy giới thiệu",
          width_cm=11.8)
fig_block(doc, ref6, "hinh_5_3_google_sheet_result.png",
          "Hình 5.3: Bảng tính Google Sheet sau khi hệ thống xử lý (trích 8 dòng thực tế, đã ẩn danh họ tên)")
fig_block(doc, ref6, "hinh_5_4_mau_pdf_ggt.png",
          "Hình 5.4: Giấy giới thiệu thực tập PDF do hệ thống tạo, đúng 1 trang và còn nguyên khoảng ký tên",
          width_cm=13.0)
fig_block(doc, ref6, "hinh_5_5_menu_sheet.png",
          "Hình 5.5: Menu quản trị GGT-CNTT trên Google Sheet")

CH5C = [
    ("h2", "5.4 Kiểm thử hệ thống"),
    ("h3", "5.4.1 Bộ kiểm thử tự động"),
    ("body", "Bộ kiểm thử được viết thành tệp test_suite.js chạy độc lập bằng Node.js: mã nguồn "
             "Code.gs được nạp vào một sandbox và các dịch vụ của Google (SpreadsheetApp, DriveApp, "
             "PropertiesService…) được thay bằng đối tượng giả lập. Nhờ vậy có thể kiểm thử logic "
             "chuẩn hóa, đối soát và ánh xạ cột ngay trên máy tính mà không cần chạy thật trên "
             "Google, đồng thời không tiêu tốn hạn mức của tài khoản Khoa. Kết quả lần chạy gần nhất: "
             "93/93 ca đạt, 0 ca lỗi (Bảng 5.2)."),
    ("caption_table", "Bảng 5.2: Các nhóm kiểm thử tự động và kết quả"),
    ("table", [
        ["Nhóm kiểm thử", "Nội dung kiểm tra", "Số ca", "Kết quả"],
        ["Chuẩn hóa số điện thoại", "Số 10 chữ số, tiền tố +84/84, thiếu số 0, dạng E+", "6", "Đạt"],
        ["Chuẩn hóa MSSV", "MSSV chữ–số, dạng E+, dạng số nguyên", "3", "Đạt"],
        ["Chuẩn hóa mã số thuế", "10 chữ số, chi nhánh có gạch nối, E+, khoảng trắng", "5", "Đạt"],
        ["Đối soát tên doanh nghiệp", "Chi nhánh, tên viết tắt, tên ngoặc đơn, tỷ lệ bao hàm và "
                                      "3 ca phủ định", "8", "Đạt"],
        ["Trạng thái đã xử lý", "Done, Đã gửi, Đã tạo PDF, Đã cấp, chuỗi rỗng và chuỗi lỗi", "16", "Đạt"],
        ["Lọc theo mốc thời gian", "31/7/2026, 1/8/2026, 6/8/2026 và số serial của bảng tính", "4", "Đạt"],
        ["Kiểm tra dữ liệu đầu vào", "Hồ sơ có nhiều trường bị biến dạng số cùng lúc", "1", "Đạt"],
        ["Nhận diện và ánh xạ cột", "Sheet 16 cột của Khoa, sheet cũ 12 cột, từ khóa rút gọn, "
                                    "chống nhận nhầm cột “Link CV”/“Trạng thái học vụ”", "23", "Đạt"],
        ["Ghi kết quả vào Sheet", "writeSuccess_ và writeError_ ghi đúng cột trạng thái, link, "
                                  "thời gian, tên thuế", "8", "Đạt"],
        ["Hộp thoại chọn dòng", "Nhập dòng đơn, khoảng dòng, dữ liệu sai, ngữ cảnh không có UI", "17", "Đạt"],
        ["Menu onOpen", "Tạo menu GGT-CNTT không gây lỗi khi thiếu quyền giao diện", "2", "Đạt"],
        ["Tổng cộng", "", "93", "93/93 đạt"],
    ], [2500, 4087, 900, 1300]),
    ("gap", ""),
    ("h3", "5.4.2 Kiểm thử đối soát trên các tình huống thực tế"),
    ("body", "Ngoài bộ kiểm thử tự động, thuật toán đối soát còn được chạy trực tiếp trên các mã số "
             "thuế thật xuất hiện trong dữ liệu của Khoa. Bảng 5.3 trình bày 8 tình huống tiêu biểu: "
             "sáu ca hợp lệ (khớp chính xác, viết tắt loại hình, chi nhánh, ngoặc đơn, tên ngắn, "
             "viết tắt thương hiệu) và hai ca bị chặn."),
    ("caption_table", "Bảng 5.3: Kết quả đối soát trên các tình huống thực tế"),
    ("table", [
        ["STT", "MST", "Tên sinh viên khai", "Tình huống", "Kết quả thực tế"],
        ["1", "0301425912", "Công ty TNHH Giải Pháp Phần Mềm Tường Minh", "Khớp chính xác",
         "Hợp lệ — nhiều hồ sơ nhất (53 giấy)"],
        ["2", "0313577878", "CTY TNHH MTV INNOVATIVE SOFTWARE DEVELOPMENT", "Viết tắt loại hình",
         "Hợp lệ — mở rộng CTY/TNHH/MTV"],
        ["3", "0300942001-008", "Tổng công ty Điện lực miền Nam", "Chi nhánh",
         "Hợp lệ — khớp công ty tổng của chi nhánh"],
        ["4", "0300954529", "Tập đoàn Bưu chính Viễn thông Việt Nam (Trung tâm Viễn thông Nam Sài Gòn)",
         "Có phần ngoặc đơn", "Hợp lệ — bóc tách phần trong ngoặc"],
        ["5", "0305140259", "Tổ chức tài chính vi mô CEP", "Tên ngắn, tỷ lệ bao hàm",
         "Hợp lệ — inclusion ratio đạt ngưỡng"],
        ["6", "0301596604", "CÔNG TY CỔ PHẦN CÔNG NGHỆ TIÊN PHONG (viết tắt ITD)",
         "Khớp viết tắt thương hiệu", "Hợp lệ — khớp từ khóa ITD"],
        ["7", "0301425912", "Công ty Cổ phần Công nghệ FPT", "Dùng MST của doanh nghiệp khác",
         "Bị chặn — không tương đồng, ghi rõ tên đăng ký thật"],
        ["8", "9999999999", "Công ty Không Tồn Tại", "MST không tồn tại",
         "Bị chặn — API trả về không có dữ liệu"],
    ], [500, 1500, 2700, 1700, 2387]),
    ("gap", ""),
]
insert_pars(doc, ref6, CH5C)

CH5D = [
    ("h2", "5.5 Kết quả thực nghiệm trên dữ liệu thật của Khoa"),
    ("body", "Đợt thực nghiệm được thực hiện trên toàn bộ 361 hồ sơ đang có trong Sheet của Khoa "
             "(đăng ký trải từ năm 2023 đến tháng 10/2026), tương ứng 361 dòng được hệ thống cập "
             "nhật đầy đủ trạng thái, liên kết PDF, thời gian xử lý và kết quả đối soát. Kết quả "
             "tổng hợp được trình bày trong Bảng 5.4 và Hình 5.6. Các biểu đồ trong chương này "
             "được sinh trực tiếp từ tệp Sheet bằng script figures/make_figures.py, không nhập số "
             "liệu bằng tay, nên có thể chạy lại bất cứ lúc nào khi Sheet thay đổi."),
    ("caption_table", "Bảng 5.4: Kết quả xử lý 361 hồ sơ thực tế"),
    ("table", [
        ["Kết quả", "Số hồ sơ", "Tỷ lệ", "Ghi chú"],
        ["Đã tạo PDF (cấp giấy giới thiệu)", "330", "91,41%",
         "Mỗi hồ sơ có một liên kết Drive riêng trong cột “LINK PDF GGT”"],
        ["Lỗi — bị chặn và ghi rõ nguyên nhân", "31", "8,59%",
         "Cột “KẾT QUẢ ĐỐI SOÁT MST” ghi chi tiết để sinh viên sửa và chạy lại"],
        ["Đang xử lý", "0", "0%", "Không còn hồ sơ nào ở trạng thái Processing"],
        ["Chờ duyệt", "0", "0%", "REQUIRE_APPROVAL đang tắt trong đợt này"],
        ["Tổng cộng", "361", "100%", "Dữ liệu lấy trực tiếp từ Sheet ngày 18/10/2026"],
    ], [3300, 1200, 1100, 3187]),
    ("body", "Tỷ lệ cấp giấy tự động thành công đạt 91,41%. Đáng chú ý, phần lớn trong 31 hồ sơ bị "
             "chặn không phải lỗi kỹ thuật của hệ thống mà là dữ liệu đầu vào không hợp lệ do sinh "
             "viên nhập, và hệ thống đã ngăn được việc cấp giấy sai. Bảng 5.5 phân loại chi tiết 31 "
             "hồ sơ này."),
    ("caption_table", "Bảng 5.5: Phân loại 31 hồ sơ bị hệ thống chặn"),
    ("table", [
        ["Nhóm nguyên nhân", "Số ca", "Ví dụ thực tế trong Sheet"],
        ["Mã số thuế sai cú pháp hoặc để trống", "11",
         "SPS, DN, “Không có”, “chưa ạ”, “(để trống)”, 00001 (2 ca), 079209300023, 1170063, "
         "010177816, hoặc ghi cả tên công ty vào ô mã số thuế"],
        ["Mã số thuế thuộc doanh nghiệp đã ngừng hoạt động", "8",
         "0317537019 (Plink), 0104753865-006 (Viettel Trà Vinh), 0101778163-006 (FPT Telecom "
         "Tân Thuận), 1200132693 (Viễn thông Tiền Giang), 1500189108-039 (VNPT Vĩnh Long)…"],
        ["Mã số thuế không tồn tại trên hệ thống Thuế", "5",
         "0100684379 (2 ca), 0301477926, 0301777291, 0100150619044"],
        ["Mã số thuế đúng nhưng tên doanh nghiệp không khớp", "5",
         "0301596604, 0300555450, 0312692469, 0107779290, 2601099276 (chi tiết ở Bảng 5.6)"],
        ["Thiếu tên doanh nghiệp", "1", "Một hồ sơ chỉ ghi mã số thuế 0318431978, bỏ trống tên công ty"],
        ["Số điện thoại sai định dạng", "1", "03378010472 (11 chữ số, vượt quá độ dài hợp lệ)"],
        ["Tổng cộng", "31", ""],
    ], [3100, 800, 4887]),
]
insert_pars(doc, ref6, CH5D)
fig_block(doc, ref6, "hinh_5_6_bieu_do_ket_qua.png",
          "Hình 5.6: Tỷ lệ xử lý và phân loại 31 hồ sơ bị chặn (số liệu thật từ Sheet)")

CH5E = [
    ("body", "Năm hồ sơ ở nhóm “mã số thuế đúng nhưng tên doanh nghiệp không khớp” là bằng chứng rõ "
             "nhất cho giá trị của bước đối soát: nếu chỉ kiểm tra mã số thuế tồn tại, hệ thống — "
             "hoặc cán bộ Khoa — đã có thể cấp giấy ghi sai đơn vị tiếp nhận. Bảng 5.6 đối chiếu "
             "từng trường hợp."),
    ("caption_table", "Bảng 5.6: Năm hồ sơ có mã số thuế đúng nhưng tên doanh nghiệp khác biệt"),
    ("table", [
        ["Mã số thuế", "Tên sinh viên khai trên Form", "Tên doanh nghiệp theo cơ quan Thuế"],
        ["0301596604", "CÔNG TY CỔ PHẦN CÔNG NGHỆ TIÊN PHONG", "CÔNG TY CỔ PHẦN CÔNG NGHỆ ITD"],
        ["0300555450", "TỔNG KHO XĂNG DẦU NHÀ BÈ", "CÔNG TY TNHH MTV PETROLIMEX SÀI GÒN"],
        ["0312692469", "ESTEC", "CÔNG TY TNHH KỸ THUẬT CÔNG NGHỆ ĐIỆN TỰ ĐỘNG BIỂN ĐÔNG"],
        ["0107779290", "CÔNG TY TÀI CHÍNH TRÁCH NHIỆM HỮU HẠN NGÂN HÀNG…",
         "CÔNG TY TÀI CHÍNH TÍN DỤNG TIÊU DÙNG TNHH MTV SFINANCE"],
        ["2601099276", "Công ty TNHH NHANFUHOSTAR", "CÔNG TY TNHH PHÁT TRIỂN NÔNG SẢN TÂY BẮC"],
    ], [1700, 3487, 3600]),
    ("body", "Hình 5.7 thống kê phương thức đối soát đã xác nhận 330 hồ sơ: 296 hồ sơ khớp chính "
             "xác ngay từ bước so chuỗi đầu tiên, 20 hồ sơ cần đến độ tương đồng có trọng số, 7 hồ "
             "sơ khớp nhờ tên viết tắt thương hiệu, 6 hồ sơ khớp theo tên công ty tổng của chi nhánh "
             "và 1 hồ sơ khớp qua tên tiếng nước ngoài. Nói cách khác, chỉ 34/330 hồ sơ (10,3%) cần "
             "đến các nhánh xử lý nâng cao; phần còn lại khớp chính xác nên kết quả vẫn dễ giải thích "
             "và dễ đối chiếu với cơ quan Thuế."),
    ("body", "Về phía các doanh nghiệp tiếp nhận, 330 giấy giới thiệu được cấp cho 183 doanh nghiệp "
             "khác nhau, trong đó có những đơn vị nhận rất nhiều sinh viên cùng lúc: Công ty TNHH "
             "Giải pháp Phần mềm Tường Minh (53 hồ sơ, chiếm 16,1%), Bosch Global Software "
             "Technologies và Công ty Cổ phần IDTEK (cùng 10 hồ sơ), Công ty TNHH Kỹ thuật Công nghệ "
             "Điện tự động Biển Đông (9), Công ty TNHH Truyền hình FPT (7). Hình 5.8 trình bày "
             "nhóm doanh nghiệp nhận nhiều sinh viên nhất và phân bố theo học phần."),
]
insert_pars(doc, ref6, CH5E)
fig_block(doc, ref6, "hinh_5_7_phuong_thuc_doi_soat.png",
          "Hình 5.7: Các phương thức đối soát đã xác nhận 330 hồ sơ hợp lệ")
fig_block(doc, ref6, "hinh_5_8_doanh_nghiep_hoc_phan.png",
          "Hình 5.8: Top 10 doanh nghiệp tiếp nhận và phân bố hồ sơ theo học phần")

CH5F = [
    ("h2", "5.6 Hiệu năng và thời gian xử lý"),
    ("body", "Thời gian xử lý tự động được trích xuất từ chính cột “THỜI GIAN XỬ LÝ GGT” do hệ thống "
             "ghi lại, nên phản ánh đúng nhịp chạy thật thay vì số đo trong phòng thí nghiệm. Đợt "
             "chạy ngày 17–18/10/2026 gồm bốn đoạn xử lý liên tục: 10 hồ sơ trong 3 phút 26 giây, "
             "42 hồ sơ trong 9 phút 46 giây, 40 hồ sơ trong 8 phút 37 giây và 269 hồ sơ trong 60 "
             "phút 51 giây. Tổng thời gian máy chạy thực tế cho cả 361 hồ sơ là 82,7 phút, tương "
             "đương 13,7 giây cho một hồ sơ, tức khoảng 4,4 hồ sơ mỗi phút."),
    ("body", "Khoảng cách trung vị giữa hai hồ sơ liên tiếp là 11 giây — phù hợp với thiết kế nghỉ "
             "BATCH_SLEEP_MS = 3.200 ms cộng thời gian gọi API tra cứu thuế, tạo bản sao Google Doc "
             "và xuất PDF. Bảng 5.7 so sánh với cách làm thủ công; mốc thủ công 5 phút/hồ sơ được lấy "
             "từ khảo sát thực tế tại Văn phòng Khoa (thao tác sao chép dữ liệu, kiểm tra doanh "
             "nghiệp, định dạng và xuất tệp)."),
    ("caption_table", "Bảng 5.7: So sánh hiệu năng giữa cách làm thủ công và hệ thống tự động"),
    ("table", [
        ["Chỉ số", "Thủ công (ước lượng)", "GGT-CNTT Auto (đo thật)", "Chênh lệch"],
        ["Thời gian cho 1 hồ sơ", "5 phút (300 giây)", "13,7 giây", "Giảm 95,4%"],
        ["Tổng thời gian cho 361 hồ sơ", "1.805 phút (≈30,1 giờ)", "82,7 phút (4 đợt chạy)",
         "Tiết kiệm ≈28,7 giờ"],
        ["Thông lượng", "0,2 hồ sơ/phút", "4,4 hồ sơ/phút (≈262 hồ sơ/giờ)", "Nhanh hơn khoảng 22 lần"],
        ["Số tệp PDF tạo ra", "330 tệp nếu làm thủ công", "330 tệp tự động, ≈140 KB/tệp (≈45 MB)",
         "Không đổi về số lượng"],
        ["Ca sai lệch tên bị bỏ lọt", "Dễ bỏ sót vì kiểm tra bằng mắt", "Chưa phát hiện ca nào",
         "5 ca MST đúng nhưng tên lệch đều bị chặn trước khi cấp giấy"],
    ], [2900, 2200, 2300, 1387]),
    ("body", "Không nên hiểu 82,7 phút là thời gian phải ngồi chờ: đây là tổng thời gian máy thực "
             "hiện công việc. Trong đợt chạy dài nhất, hệ thống xử lý 269 hồ sơ trong gần 61 phút, "
             "tức trung bình khoảng 4,4 hồ sơ mỗi phút. Nếu quy đổi, một cán bộ Khoa cần khoảng 30 "
             "giờ làm việc liên tục mới hoàn tất khối lượng tương đương, trong khi hệ thống tự chạy "
             "nền và chỉ cần người dùng quyết định đợt chạy."),
]
insert_pars(doc, ref6, CH5F)
fig_block(doc, ref6, "hinh_5_9_bieu_do_thoi_gian.png",
          "Hình 5.9: So sánh thời gian và các chỉ số hiệu năng đo trên dữ liệu thật")

CH5G = [
    ("h2", "5.7 Vận hành hàng loạt và Dashboard"),
    ("body", "Nhờ cơ chế hàng đợi nền trình bày ở Mục 4.7.2, toàn bộ 361 hồ sơ được xử lý trong bốn "
             "đợt mà không đợt nào chạm giới hạn thời gian của Google Apps Script. Cán bộ Khoa chỉ "
             "cần nhập danh sách dòng vào hộp thoại (ví dụ “2-361”) hoặc dùng các mục kiểm thử riêng "
             "lẻ; phần còn lại của lô tự tiếp tục sau mỗi phút cho tới khi hoàn tất."),
    ("body", "Hàm refreshDashboard() tổng hợp kết quả thành sheet “Dashboard GGT” (Hình 5.10) với các "
             "chỉ tiêu theo đúng công thức COUNTIF trên cột trạng thái. Tại thời điểm kết thúc đợt "
             "chạy, Dashboard ghi nhận: tổng 361 hồ sơ; 330 hồ sơ đã tạo PDF; 31 hồ sơ lỗi; 0 hồ sơ "
             "đang xử lý và 0 hồ sơ chờ duyệt. Dashboard giúp cán bộ Khoa nắm tình trạng xử lý mà "
             "không phải tự đếm dòng, đồng thời là căn cứ để chạy lại các hồ sơ lỗi bằng hàm "
             "retryPending() sau khi sinh viên sửa dữ liệu."),
    ("body", "Một điểm cần lưu ý trong vận hành: các hồ sơ đã cấp giấy được nhận diện qua hàm "
             "isAlreadyProcessed_() nên khi chạy lại toàn bộ danh sách, hệ thống bỏ qua chúng và chỉ "
             "xử lý phần còn thiếu. Cơ chế này tránh việc vô tình tạo lại hàng loạt tệp PDF trùng, "
             "đồng thời tiết kiệm đáng kể số lần gọi API tra cứu thuế."),
]
insert_pars(doc, ref6, CH5G)
fig_block(doc, ref6, "hinh_5_10_dashboard.png",
          "Hình 5.10: Dashboard tổng hợp kết quả vận hành với số liệu thực tế")

CH5H = [
    ("h2", "5.8 Vấn đề gặp phải và cách khắc phục"),
    ("body", "Quá trình xây dựng và vận hành hệ thống đã bộc lộ một số vấn đề thực tế. Bảng 5.8 "
             "tổng hợp nguyên nhân và giải pháp đã áp dụng."),
    ("caption_table", "Bảng 5.8: Các vấn đề thực tế và giải pháp đã áp dụng"),
    ("table", [
        ["Vấn đề", "Nguyên nhân", "Giải pháp"],
        ["MSSV, MST, số điện thoại bị biến dạng", "Bảng tính tự chuyển số dài sang dạng khoa học "
                                                 "(E+) hoặc cắt số 0 đầu; thực tế 206/361 giá trị "
                                                 "MSSV và 42/361 giá trị MST bị ảnh hưởng",
         "Ba hàm normalizeMssv_, normalizeMst_, normalizePhone_ khôi phục số nguyên, bù số 0 và "
         "chuẩn hóa tiền tố +84/84"],
        ["API tra cứu trả mã 429", "Vượt hạn mức 10 yêu cầu/30 giây khi chạy hàng loạt",
         "Đọc header Retry-After, chờ tăng dần (6–12 giây), cache kết quả 6 giờ và nghỉ 3.200 ms "
         "giữa hai hồ sơ"],
        ["PDF tràn sang trang 2", "Google Doc còn ngắt trang hoặc quá nhiều đoạn trắng ở cuối",
         "Hàm trimTrailingBlank_ xóa mọi page break và chỉ cắt các dòng trắng dư phía sau nội dung "
         "cuối cùng"],
        ["Mất khoảng trống ký tên", "Phiên bản đầu của trimTrailingBlank_ xóa toàn bộ đoạn trắng ở "
                                    "cuối, xóa nhầm vùng dành cho chữ ký",
         "Viết lại: giữ tối thiểu 2 dòng trắng sau nội dung cuối và thu nhỏ dòng trắng cuối xuống "
         "1 pt để không chiếm diện tích"],
        ["Chạy hàng trăm hồ sơ bị dừng giữa chừng", "Giới hạn 6 phút cho mỗi lần thực thi của Apps "
                                                    "Script",
         "Hàng đợi nền có checkpoint: 25 dòng và 240 giây mỗi đợt, trigger mỗi phút tự chạy tiếp, "
         "ghi tiến độ sau từng dòng"],
        ["Trạng thái bị xử lý trùng hoặc bỏ sót", "Chuỗi trạng thái cũ trong Sheet rất đa dạng "
                                                 "(“Done”, “[10/7/2025] Done”, “ĐỒNG Ý CẤP GIẤY GTTT”…)",
         "Hàm isAlreadyProcessed_ chuẩn hóa bỏ dấu và nhận diện 7 dạng trạng thái hoàn tất, luôn "
         "coi trạng thái bắt đầu bằng “Lỗi” là chưa hoàn tất"],
        ["Cột quản trị bị tạo trùng", "Sheet của Khoa đã có sẵn các cột trạng thái, link, thời gian",
         "getColumns_ quét và nhận diện cột theo tiêu đề chính xác trước, chỉ tạo cột mới khi thật "
         "sự chưa có; 23 ca kiểm thử bảo vệ quy tắc này"],
        ["Nguy cơ lộ dữ liệu qua liên kết PDF", "Mặc định cũ chia sẻ công khai theo liên kết",
         "Chuyển mặc định sang PDF_SHARING_MODE = RESTRICTED, chỉ đổi khi được đơn vị cho phép; "
         "checkConfig_ từ chối giá trị cấu hình không hợp lệ"],
    ], [2400, 3300, 3087]),
    ("gap", ""),
    ("h2", "5.9 Nhận xét kết quả"),
    ("body", "Từ kết quả thực nghiệm trên 361 hồ sơ thật, có thể rút ra một số nhận xét:"),
    ("bullet", "Hệ thống hoàn thành được mục tiêu cốt lõi: bóc tách dữ liệu, kiểm chứng mã số thuế, "
               "đối soát tên doanh nghiệp, sinh văn bản PDF và ghi kết quả về Sheet mà không cần "
               "máy chủ hay thao tác thủ công nào."),
    ("bullet", "Tỷ lệ cấp giấy tự động đạt 91,41% (330/361). Phần còn lại không phải lỗi hệ thống mà "
               "là hồ sơ có dữ liệu sai; việc chặn và ghi rõ nguyên nhân giúp sinh viên sửa đúng "
               "chỗ và chạy lại."),
    ("bullet", "Bước đối soát tên thể hiện rõ giá trị ở 5 hồ sơ có mã số thuế đúng nhưng tên doanh "
               "nghiệp khác với đăng ký thuế: hệ thống từ chối cấp giấy thay vì phát hành văn bản "
               "ghi sai đơn vị tiếp nhận."),
    ("bullet", "Tốc độ xử lý tăng khoảng 22 lần so với ước lượng thủ công (13,7 giây so với 5 phút "
               "mỗi hồ sơ); đợt chạy dài nhất hoàn tất 269 hồ sơ trong 61 phút mà không cần người "
               "theo dõi."),
    ("bullet", "Bộ kiểm thử 93 ca giúp việc chỉnh sửa mã nguồn về sau an toàn hơn: mọi thay đổi "
               "trong chuẩn hóa dữ liệu, ánh xạ cột hay luật đối soát đều được kiểm tra lại trước "
               "khi dán lên Apps Script."),
    ("bullet", "Hạn chế của số liệu: mốc thời gian thủ công 5 phút/hồ sơ là ước lượng từ khảo sát, "
               "không phải đo trực tiếp; dữ liệu trải qua nhiều năm nên các giá trị mã số thuế có "
               "thể đã thay đổi theo thời gian. Lớp AI chưa được bật trong đợt này nên chưa có số "
               "liệu đánh giá riêng cho nó."),
]
insert_pars(doc, ref6, CH5H)
print("Đã viết lại Chương 5.")

# --------------------------------------------------------------------------- #
#  5. Cập nhật Chương 6 (kết luận, hạn chế, hướng phát triển)
# --------------------------------------------------------------------------- #
replace_par(doc, "Báo cáo đề tài đã đạt được 4 kết quả định lượng cốt lõi sau",
            "Báo cáo đề tài đã đạt được 5 kết quả định lượng cốt lõi sau:\n"
            "1. Xây dựng hoàn chỉnh hệ thống GGT-CNTT Auto vận hành 100% serverless trên Google "
            "Workspace (Form, Sheet, Doc, Drive, Apps Script), không tốn chi phí hạ tầng.\n"
            "2. Vận hành thực tế trên toàn bộ 361 hồ sơ trong Sheet của Khoa: 330 hồ sơ được cấp "
            "Giấy giới thiệu PDF tự động (91,41%), 31 hồ sơ bị chặn kèm nguyên nhân (8,59%), trong "
            "đó có 5 trường hợp mã số thuế đúng nhưng tên doanh nghiệp lệch với đăng ký thuế.\n"
            "3. Rút ngắn thời gian xử lý từ ước lượng 5 phút/hồ sơ (thủ công) xuống 13,7 giây/hồ sơ "
            "(tự động), tương đương tiết kiệm khoảng 95,4% thời gian; 361 hồ sơ hoàn tất trong 82,7 "
            "phút máy chạy chia thành 4 đợt nền.\n"
            "4. Kiểm thử tự động 93 ca đạt 100%, phủ các nhóm chuẩn hóa dữ liệu, đối soát tên, ánh "
            "xạ cột, trạng thái, lọc thời gian và ghi kết quả về Sheet.\n"
            "5. Bổ sung công cụ quản trị ngay trên bảng tính: menu GGT-CNTT, hàng đợi xử lý nền có "
            "checkpoint, Dashboard tổng hợp, quy trình duyệt và cấp số tự động (có công tắc cấu "
            "hình), cùng cơ chế kiểm soát hạn mức API tra cứu thuế.")
replace_par(doc, "Thuật toán luật chưa bao quát hết biến thể tên doanh nghiệp",
            "Thuật toán luật chưa bao quát hết biến thể tên doanh nghiệp; lớp AI hỗ trợ vùng xám "
            "đã được lập trình nhưng chưa bật trong đợt thực nghiệm nên chưa có đánh giá định lượng.")
replace_par(doc, "Bộ dữ liệu thử có cỡ mẫu hạn chế",
            "Mốc so sánh thời gian thủ công (5 phút/hồ sơ) là ước lượng từ khảo sát Văn phòng Khoa, "
            "không phải số đo bấm giờ; 31 hồ sơ lỗi cũng chưa được xác nhận lại với từng sinh viên "
            "về việc đã sửa hay chưa.")
replace_par(doc, "Mở rộng AI: trích xuất thông tin từ văn bản tự do",
            "Bật và đánh giá lớp AI: chạy AI_ENABLED = true trên một tập con dữ liệu đã gán nhãn "
            "để đo chính xác của vùng xám; đồng thời mở rộng trích xuất thông tin từ văn bản tự do "
            "(email nhận việc, thư xác nhận của doanh nghiệp) và tự soạn phản hồi lỗi cho sinh viên.")

# --------------------------------------------------------------------------- #
#  6. Danh mục hình vẽ / bảng biểu
# --------------------------------------------------------------------------- #
FIGS = [
    "Hình 2.1: Sơ đồ cơ cấu tổ chức và vị trí thực tập tại Phòng Thí nghiệm AI Lab",
    "Hình 2.2: Quy trình cấp giấy giới thiệu thực tập trước khi tự động hóa",
    "Hình 3.1: Kiến trúc dịch vụ của Google Apps Script",
    "Hình 3.2: Cơ chế giới hạn tốc độ và thử lại lũy thừa (exponential backoff)",
    "Hình 3.3: Mô hình Transformer và Scaled Dot-Product Attention",
    "Hình 4.1: Kiến trúc tổng thể của hệ thống (bản cập nhật theo Code.gs mới)",
    "Hình 4.2: Lưu đồ xử lý một dòng dữ liệu",
    "Hình 4.3: Lưu đồ đối soát tên doanh nghiệp",
    "Hình 4.4: Máy trạng thái của hồ sơ",
    "Hình 4.5: Vòng đời của một lô xử lý nền trong hệ thống",
    "Hình 5.1: Cấu hình tập trung CONFIG của hệ thống",
    "Hình 5.2: Biểu mẫu Google Form thu thập thông tin cấp Giấy giới thiệu",
    "Hình 5.3: Bảng tính Google Sheet sau khi hệ thống xử lý",
    "Hình 5.4: Giấy giới thiệu thực tập PDF đúng 1 trang, bảo toàn khoảng ký tên",
    "Hình 5.5: Menu quản trị GGT-CNTT trên Google Sheet",
    "Hình 5.6: Tỷ lệ xử lý và phân loại 31 hồ sơ bị chặn",
    "Hình 5.7: Các phương thức đối soát đã xác nhận 330 hồ sơ hợp lệ",
    "Hình 5.8: Top 10 doanh nghiệp tiếp nhận và phân bố hồ sơ theo học phần",
    "Hình 5.9: So sánh thời gian và các chỉ số hiệu năng đo trên dữ liệu thật",
    "Hình 5.10: Dashboard tổng hợp kết quả vận hành",
]
TABLES = [
    "Bảng 2.1: Kế hoạch và nhật ký công việc thực tập",
    "Bảng 2.2: Yêu cầu chức năng và phi chức năng",
    "Bảng 3.1: So sánh số hóa, chuyển đổi số và tự động hóa",
    "Bảng 3.2: Các dịch vụ Apps Script được sử dụng",
    "Bảng 3.3: Ma trận nhầm lẫn",
    "Bảng 4.1: Cấu trúc các cột trong Google Sheet",
    "Bảng 4.2: Danh sách placeholder trong văn bản mẫu",
    "Bảng 5.1: Môi trường triển khai và thực nghiệm",
    "Bảng 5.2: Các nhóm kiểm thử tự động và kết quả",
    "Bảng 5.3: Kết quả đối soát trên các tình huống thực tế",
    "Bảng 5.4: Kết quả xử lý 361 hồ sơ thực tế",
    "Bảng 5.5: Phân loại 31 hồ sơ bị hệ thống chặn",
    "Bảng 5.6: Năm hồ sơ có mã số thuế đúng nhưng tên doanh nghiệp khác biệt",
    "Bảng 5.7: So sánh hiệu năng giữa cách làm thủ công và hệ thống tự động",
    "Bảng 5.8: Các vấn đề thực tế và giải pháp đã áp dụng",
    "Bảng 6.1: Đối chiếu mục tiêu và kết quả",
]
# thay danh mục hình vẽ
items = body_items(doc)
i_fig = find_body_index(items, lambda t: t == "DANH MỤC HÌNH VẼ")
i_tab = find_body_index(items, lambda t: t == "DANH MỤC BẢNG BIỂU")
i_abbr = find_body_index(items, lambda t: t == "DANH MỤC CÁC CHỮ VIẾT TẮT")
assert -1 not in (i_fig, i_tab, i_abbr)
# xóa các mục cũ
for k, o in items[i_fig + 1:i_tab]:
    (o._p if k == "p" else o._tbl).getparent().remove(o._p if k == "p" else o._tbl)
for k, o in items[i_tab + 1:i_abbr]:
    (o._p if k == "p" else o._tbl).getparent().remove(o._p if k == "p" else o._tbl)
ref_tab = items[i_tab][1]._p
for line in FIGS:
    ref_tab.addprevious(make_par(TPL["caption"], line))
ref_abbr = items[i_abbr][1]._p
for line in TABLES:
    ref_abbr.addprevious(make_par(TPL["caption"], line))
print("Đã cập nhật danh mục hình vẽ và bảng biểu.")

# --------------------------------------------------------------------------- #
#  7. Phụ lục A + B, xóa trang hướng dẫn tác giả
# --------------------------------------------------------------------------- #
replace_par(doc, "Toàn bộ mã nguồn hoàn chỉnh của hệ thống được công khai",
            "Toàn bộ mã nguồn của hệ thống được công khai tại kho lưu trữ GitHub: "
            "https://github.com/nhan1232004/tools (nhánh main).\n"
            "Các tệp chính:\n"
            "• ggt-cntt-auto/sources/Code.gs — mã nguồn Google Apps Script (1.977 dòng).\n"
            "• ggt-cntt-auto/sources/test_suite.js — bộ kiểm thử tự động 93 ca, chạy bằng Node.js.\n"
            "• ggt-cntt-auto/templates/ — mẫu Word hành chính gốc và mẫu đã gắn placeholder.\n"
            "• figures/make_figures.py — sinh toàn bộ biểu đồ từ dữ liệu Sheet thực tế.\n"
            "• tools/update_report.py — dựng lại báo cáo từ template của Khoa và số liệu nêu trên.\n"
            "Các hàm trọng tâm trong Code.gs:\n"
            "• processRow_(sheet, row, isTestMode, approvedForRelease): điều phối xử lý một dòng.\n"
            "• cleanAndValidate_ + normalizeMssv_/normalizeMst_/normalizePhone_: chuẩn hóa dữ liệu, "
            "khử dạng số khoa học E+.\n"
            "• fetchTaxPayerInfo_(mst): gọi API Thuế kèm cache 6 giờ và retry lũy thừa khi gặp 429.\n"
            "• compareCompanyNames_ + computeWeightedSimilarity_: đối soát tên có trọng số, bóc tách "
            "chi nhánh, tên viết tắt, tên nước ngoài.\n"
            "• createPdf_ + trimTrailingBlank_: sinh PDF một trang và bảo toàn khoảng trống ký tên.\n"
            "• enqueueBatch_/continueBatchQueue_/cancelBatchQueue_: hàng đợi xử lý nền có checkpoint.\n"
            "• refreshDashboard() và approveSelectedRow(): cập nhật Dashboard, duyệt hồ sơ chờ.\n"
            "• getColumns_ + findColByExactOrPartial_: nhận diện và tạo cột quản trị không trùng lặp.")
replace_par(doc, "Biểu mẫu Google Form thu thập 6 trường thông tin bắt buộc",
            "Biểu mẫu Google Form thu thập 15 trường thông tin, tương ứng 15 cột đầu của Sheet “Form "
            "Responses 1”:\n"
            "1. Dấu thời gian (Timestamp) — hệ thống tự sinh.\n"
            "2. Họ và Tên (bắt buộc).\n"
            "3. MSSV (bắt buộc, 6–12 ký tự chữ và số).\n"
            "4. Số điện thoại (bắt buộc, 10 chữ số bắt đầu bằng 0).\n"
            "5. Email cá nhân (bắt buộc).\n"
            "6. Tên học phần (Kiến tập công nghiệp / Tập sự nghề nghiệp / Dự án CNTT 1).\n"
            "7. Tên doanh nghiệp thực tập (bắt buộc).\n"
            "8. Tên doanh nghiệp viết bằng tiếng nước ngoài (tùy chọn).\n"
            "9. Tên doanh nghiệp viết tắt (tùy chọn).\n"
            "10. Mã số doanh nghiệp / mã số thuế (bắt buộc, 10 hoặc 13 chữ số).\n"
            "11. Thời gian bắt đầu thực tập (tùy chọn).\n"
            "12. Họ và tên người hướng dẫn tại doanh nghiệp (tùy chọn).\n"
            "13. Số điện thoại người hướng dẫn (tùy chọn).\n"
            "14. Email người hướng dẫn (tùy chọn).\n"
            "15. Vị trí thực tập (tùy chọn; để trống hệ thống dùng giá trị mặc định).\n"
            "Ngoài ra hệ thống tự thêm 5 cột quản trị: TRẠNG THÁI GỬI GGT, LINK PDF GGT, "
            "THỜI GIAN XỬ LÝ GGT, Tên Doanh Nghiệp (Tra cứu MST) và Kết quả đối soát MST.")

# xóa trang PHỤ LỤC HƯỚNG DẪN CHO TÁC GIẢ trở về cuối
items = body_items(doc)
i_guide = find_body_index(items, lambda t: t.startswith("PHỤ LỤC HƯỚNG DẪN CHO TÁC GIẢ"))
if i_guide != -1:
    removed = 0
    for k, o in items[i_guide:]:
        el = o._p if k == "p" else o._tbl
        if el.getparent() is not None:
            el.getparent().remove(el)
            removed += 1
    print(f"Đã xóa {removed} phần tử của trang hướng dẫn tác giả.")
else:
    print("  [cảnh báo] không tìm thấy trang hướng dẫn tác giả.")

# --------------------------------------------------------------------------- #
#  8. Bật cập nhật field (mục lục, số trang) khi mở Word
# --------------------------------------------------------------------------- #
settings = doc.settings.element
uf = settings.find(w("updateFields"))
if uf is None:
    uf = OxmlElement("w:updateFields")
    settings.append(uf)
uf.set(w("val"), "true")

doc.save(OUT)
print("Đã lưu:", os.path.relpath(OUT, ROOT))
