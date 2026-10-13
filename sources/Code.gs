/**
 * TỰ ĐỘNG TẠO GIẤY GIỚI THIỆU THỰC TẬP (GGT-CNTT)
 * Luồng: Google Form -> Google Sheet -> điền Google Doc mẫu -> xuất PDF -> gửi Gmail cho giảng viên
 *
 * 6 câu hỏi của Form (tiêu đề câu hỏi phải chứa các từ khóa bên dưới):
 *   1. Họ và tên
 *   2. Mã số sinh viên - MSSV
 *   3. Email cá nhân
 *   4. Số điện thoại
 *   5. Vị trí thực tập
 *   6. Tên Công ty / Doanh nghiệp thực tập
 *
 * Placeholder trong Google Doc mẫu:
 *   {{HO_TEN}}  {{MSSV}}  {{EMAIL}}  {{SDT}}  {{VI_TRI}}  {{CONG_TY}}
 */

// ======================= 1. CẤU HÌNH (CHỈ CẦN SỬA PHẦN NÀY) =======================
const CONFIG = {
  TEMPLATE_DOC_ID: 'DAN_ID_GOOGLE_DOC_MAU_VAO_DAY',
  OUTPUT_FOLDER_ID: 'DAN_ID_THU_MUC_DRIVE_LUU_PDF_VAO_DAY',
  LECTURER_EMAIL: 'email.giangvien@example.com',  // người nhận PDF
  CC_EMAILS: '',                                  // nhiều mail cách nhau dấu phẩy, để trống nếu không cần
  ADMIN_EMAIL: '',                                // nhận mail báo lỗi; để trống = chính chủ tài khoản chạy script
  SENDER_NAME: 'Hệ thống GGT - Khoa CNTT',
  SHEET_NAME: '',                                 // để trống = sheet đầu tiên của file

  // Tên 3 cột do script tự thêm vào cuối Sheet
  STATUS_HEADER: 'Trạng thái',
  LINK_HEADER: 'Link PDF',
  TIME_HEADER: 'Thời gian xử lý'
};

// Từ khóa (viết thường) để tìm đúng cột theo tiêu đề câu hỏi. Từ khóa đứng trước được ưu tiên.
const FIELD_KEYWORDS = {
  HO_TEN:  ['họ và tên', 'họ tên', 'ho va ten'],
  MSSV:    ['mssv', 'mã số sinh viên', 'ma so sinh vien'],
  EMAIL:   ['email cá nhân', 'email'],
  SDT:     ['số điện thoại', 'điện thoại', 'sdt', 'so dien thoai'],
  VI_TRI:  ['vị trí', 'vi tri'],
  CONG_TY: ['tên công ty', 'công ty', 'doanh nghiệp', 'cong ty']
};
const FIELD_KEYS = Object.keys(FIELD_KEYWORDS);

// ======================= 2. HÀM CHÍNH (TRIGGER) =======================

/** Trigger "Khi gửi biểu mẫu": chạy mỗi khi có sinh viên nộp Form. */
function handleFormSubmit(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // tránh 2 sinh viên nộp cùng lúc làm chồng xử lý
  try {
    if (e && e.range) {
      processRow_(e.range.getSheet(), e.range.getRow());
    } else {
      // Chạy tay trong editor (không có sự kiện từ Form): xử lý dòng cuối cùng
      const sheet = getResponseSheet_();
      if (sheet.getLastRow() >= 2) processRow_(sheet, sheet.getLastRow());
    }
  } finally {
    lock.releaseLock();
  }
}

/** Xử lý 1 dòng: kiểm tra dữ liệu -> tạo PDF -> gửi mail -> ghi trạng thái. */
function processRow_(sheet, row) {
  const cols = getColumns_(sheet);
  const rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  const currentStatus = String(rowValues[cols.STATUS - 1] || '');
  if (currentStatus.indexOf('Đã gửi') === 0) return; // chống gửi trùng

  try {
    const raw = {};
    FIELD_KEYS.forEach(function (k) { raw[k] = rowValues[cols[k] - 1]; });

    const data = cleanAndValidate_(raw);
    const pdf = createPdf_(data);
    sendMail_(data, pdf.blob, pdf.url);

    writeStatus_(sheet, row, cols, 'Đã gửi', pdf.url);
  } catch (err) {
    writeStatus_(sheet, row, cols, 'Lỗi: ' + err.message, '');
    notifyAdmin_(row, err);
  }
}

// ======================= 3. CHUẨN HÓA & KIỂM TRA DỮ LIỆU =======================

function cleanAndValidate_(raw) {
  const errors = [];

  const hoTen = titleCase_(collapse_(raw.HO_TEN));
  const mssv = collapse_(raw.MSSV).replace(/\s/g, '').toUpperCase();
  const email = collapse_(raw.EMAIL).toLowerCase();
  const viTri = collapse_(raw.VI_TRI);
  const congTy = collapse_(raw.CONG_TY);

  // Số điện thoại: bỏ ký tự thừa, đổi +84/84 -> 0, bù số 0 bị Google Sheet làm mất
  let sdt = String(raw.SDT === undefined || raw.SDT === null ? '' : raw.SDT).replace(/[\s.\-()]/g, '');
  if (sdt.indexOf('+84') === 0) sdt = '0' + sdt.slice(3);
  else if (sdt.indexOf('84') === 0 && sdt.length === 11) sdt = '0' + sdt.slice(2);
  else if (/^[1-9]\d{8}$/.test(sdt)) sdt = '0' + sdt;

  if (!hoTen) errors.push('thiếu Họ và tên');
  if (!/^[A-Z0-9]{6,12}$/.test(mssv)) errors.push('MSSV không hợp lệ (' + mssv + ')');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) errors.push('Email không hợp lệ (' + email + ')');
  if (!/^0\d{9}$/.test(sdt)) errors.push('Số điện thoại không hợp lệ (' + sdt + ')');
  if (!viTri) errors.push('thiếu Vị trí thực tập');
  if (!congTy) errors.push('thiếu Tên công ty');

  if (errors.length) throw new Error('Dữ liệu không hợp lệ: ' + errors.join('; '));

  return { hoTen: hoTen, mssv: mssv, email: email, sdt: sdt, viTri: viTri, congTy: congTy };
}

function collapse_(v) {
  return String(v === undefined || v === null ? '' : v).replace(/\s+/g, ' ').trim();
}

function titleCase_(s) {
  return s.toLowerCase().split(' ').map(function (w) {
    return w ? w.charAt(0).toUpperCase() + w.slice(1) : w;
  }).join(' ');
}

function removeAccents_(s) {
  return s.normalize('NFD').replace(/[\u0300-\u036f]/g, '')
          .replace(/đ/g, 'd').replace(/Đ/g, 'D');
}

// ======================= 4. ĐIỀN MẪU & XUẤT PDF =======================

function createPdf_(data) {
  const template = DriveApp.getFileById(CONFIG.TEMPLATE_DOC_ID);
  const folder = DriveApp.getFolderById(CONFIG.OUTPUT_FOLDER_ID);

  const baseName = 'GGT_' + data.mssv + '_' + removeAccents_(data.hoTen).replace(/\s+/g, '');
  const tmp = template.makeCopy(baseName + '_tmp', folder);

  try {
    const doc = DocumentApp.openById(tmp.getId());
    const body = doc.getBody();

    const map = {
      '{{HO_TEN}}': data.hoTen,
      '{{MSSV}}': data.mssv,
      '{{EMAIL}}': data.email,
      '{{SDT}}': data.sdt,
      '{{VI_TRI}}': data.viTri,
      '{{CONG_TY}}': data.congTy
    };
    Object.keys(map).forEach(function (key) {
      body.replaceText(escapeRegex_(key), escapeReplacement_(map[key]));
    });
    trimTrailingBlank_(body); // bỏ đoạn trống cuối để PDF không bị thừa trang trắng
    doc.saveAndClose();
    Utilities.sleep(1500); // chờ Drive ghi xong trước khi xuất PDF

    const blob = tmp.getAs(MimeType.PDF).setName(baseName + '.pdf');
    const pdfFile = folder.createFile(blob);
    return { blob: pdfFile.getBlob(), url: pdfFile.getUrl() };
  } finally {
    tmp.setTrashed(true); // xóa bản Doc tạm, chỉ giữ lại PDF
  }
}

/**
 * Loại bỏ phần thừa ở cuối Doc gây ra trang trắng:
 *  - xóa các đoạn trống nằm sau nội dung cuối cùng
 *  - đoạn cuối cùng (Google Docs bắt buộc phải có) nếu trống thì thu nhỏ còn 1pt và bỏ giãn cách
 *  - xóa dấu ngắt trang còn sót ở đoạn cuối
 */
function trimTrailingBlank_(body) {
  const PARA = DocumentApp.ElementType.PARAGRAPH;

  // 1) Xóa ngược các đoạn trống đứng trước đoạn cuối cùng
  let idx = body.getNumChildren() - 2;
  while (idx >= 0) {
    const child = body.getChild(idx);
    if (child.getType() !== PARA || !isBlankParagraph_(child.asParagraph())) break;
    child.removeFromParent();
    idx--;
  }

  // 2) Xử lý đoạn cuối cùng
  const last = body.getChild(body.getNumChildren() - 1);
  if (last.getType() !== PARA) return;
  const p = last.asParagraph();

  for (let i = p.getNumChildren() - 1; i >= 0; i--) {
    if (p.getChild(i).getType() === DocumentApp.ElementType.PAGE_BREAK) {
      p.getChild(i).removeFromParent();
    }
  }
  p.setSpacingAfter(0);
  if (isBlankParagraph_(p)) {
    p.setFontSize(1).setSpacingBefore(0).setLineSpacing(1);
  }
}

function isBlankParagraph_(p) {
  if (p.getText().replace(/[\s\u00A0]/g, '') !== '') return false;
  for (let i = 0; i < p.getNumChildren(); i++) {
    const t = p.getChild(i).getType();
    if (t === DocumentApp.ElementType.INLINE_IMAGE ||
        t === DocumentApp.ElementType.INLINE_DRAWING ||
        t === DocumentApp.ElementType.HORIZONTAL_RULE) return false;
  }
  return true;
}

function escapeRegex_(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function escapeReplacement_(s) {
  return String(s).replace(/\\/g, '\\\\').replace(/\$/g, '\\$');
}

// ======================= 5. GỬI MAIL =======================

function sendMail_(data, pdfBlob, pdfUrl) {
  const subject = '[GGT-CNTT] Giấy giới thiệu thực tập - ' + data.hoTen + ' (' + data.mssv + ')';

  const plain =
    'Kính gửi Thầy/Cô,\n\n' +
    'Hệ thống đã tạo giấy giới thiệu thực tập cho sinh viên:\n' +
    '- Họ và tên: ' + data.hoTen + '\n' +
    '- MSSV: ' + data.mssv + '\n' +
    '- Email: ' + data.email + '\n' +
    '- Số điện thoại: ' + data.sdt + '\n' +
    '- Vị trí thực tập: ' + data.viTri + '\n' +
    '- Công ty: ' + data.congTy + '\n\n' +
    'File PDF được đính kèm. Link Drive: ' + pdfUrl;

  const html =
    '<p>Kính gửi Thầy/Cô,</p>' +
    '<p>Hệ thống đã tạo giấy giới thiệu thực tập cho sinh viên:</p>' +
    '<table style="border-collapse:collapse">' +
    row_('Họ và tên', data.hoTen) +
    row_('MSSV', data.mssv) +
    row_('Email', data.email) +
    row_('Số điện thoại', data.sdt) +
    row_('Vị trí thực tập', data.viTri) +
    row_('Công ty', data.congTy) +
    '</table>' +
    '<p>File PDF được đính kèm trong mail này. ' +
    '<a href="' + pdfUrl + '">Mở trên Google Drive</a></p>';

  const options = { attachments: [pdfBlob], htmlBody: html, name: CONFIG.SENDER_NAME };
  if (CONFIG.CC_EMAILS) options.cc = CONFIG.CC_EMAILS;

  GmailApp.sendEmail(CONFIG.LECTURER_EMAIL, subject, plain, options);
}

function row_(label, value) {
  return '<tr><td style="padding:2px 12px 2px 0"><b>' + escapeHtml_(label) + ':</b></td>' +
         '<td style="padding:2px 0">' + escapeHtml_(value) + '</td></tr>';
}

function escapeHtml_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
                  .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function notifyAdmin_(row, err) {
  try {
    const to = CONFIG.ADMIN_EMAIL || Session.getEffectiveUser().getEmail();
    MailApp.sendEmail(to, '[GGT-CNTT] Lỗi xử lý dòng ' + row,
      'Không xử lý được dòng ' + row + ' trong Sheet.\n\nNguyên nhân: ' + err.message +
      '\n\nSửa dữ liệu rồi chạy hàm retryPending() để xử lý lại.');
  } catch (ignore) {}
}

// ======================= 6. TIỆN ÍCH SHEET =======================

/** Tìm vị trí các cột (1-based); tự thêm 3 cột Trạng thái / Link PDF / Thời gian nếu chưa có. */
function getColumns_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                       .map(function (h) { return String(h).trim(); });

  [CONFIG.STATUS_HEADER, CONFIG.LINK_HEADER, CONFIG.TIME_HEADER].forEach(function (name) {
    if (headers.indexOf(name) === -1) {
      sheet.getRange(1, headers.length + 1).setValue(name);
      headers.push(name);
    }
  });

  const lower = headers.map(function (h) { return h.toLowerCase(); });
  const cols = {
    STATUS: headers.indexOf(CONFIG.STATUS_HEADER) + 1,
    LINK: headers.indexOf(CONFIG.LINK_HEADER) + 1,
    TIME: headers.indexOf(CONFIG.TIME_HEADER) + 1
  };

  FIELD_KEYS.forEach(function (key) {
    let found = -1;
    FIELD_KEYWORDS[key].some(function (kw) {
      for (let i = 0; i < lower.length; i++) {
        if (lower[i].indexOf(kw) !== -1) { found = i + 1; return true; }
      }
      return false;
    });
    if (found === -1) {
      throw new Error('Không tìm thấy cột cho "' + key + '". Kiểm tra tiêu đề câu hỏi trong Form.');
    }
    cols[key] = found;
  });
  return cols;
}

function writeStatus_(sheet, row, cols, status, link) {
  sheet.getRange(row, cols.STATUS).setValue(status);
  sheet.getRange(row, cols.LINK).setValue(link);
  sheet.getRange(row, cols.TIME).setValue(
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
}

function getResponseSheet_() {
  const ss = SpreadsheetApp.getActive();
  if (CONFIG.SHEET_NAME) return ss.getSheetByName(CONFIG.SHEET_NAME);

  // Tự tìm sheet có đủ 6 cột của Form (không phụ thuộc vị trí sheet)
  const sheets = ss.getSheets();
  for (let s = 0; s < sheets.length; s++) {
    const sh = sheets[s];
    if (sh.getLastColumn() === 0) continue;
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
                      .map(function (h) { return String(h).trim().toLowerCase(); });
    const ok = FIELD_KEYS.every(function (k) {
      return FIELD_KEYWORDS[k].some(function (kw) {
        return headers.some(function (h) { return h.indexOf(kw) !== -1; });
      });
    });
    if (ok) return sh;
  }
  throw new Error('Không tìm thấy sheet câu trả lời của Form. Hãy liên kết Form có đủ 6 câu hỏi với Sheet này.');
}

// ======================= 7. CÀI ĐẶT & KIỂM THỬ (CHẠY TAY TRONG EDITOR) =======================

/** Chạy 1 lần duy nhất: tạo trigger, thêm cột trạng thái, kiểm tra cấu hình. */
function setup() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'handleFormSubmit') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('handleFormSubmit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onFormSubmit()
    .create();

  getColumns_(getResponseSheet_());
  checkConfig_();
  Logger.log('Cài đặt xong. Hãy nộp thử Form hoặc chạy testLastRow().');
}

/** Kiểm tra ID mẫu, ID thư mục và các placeholder trong Doc mẫu. */
function checkConfig_() {
  DriveApp.getFolderById(CONFIG.OUTPUT_FOLDER_ID);
  const text = DocumentApp.openById(CONFIG.TEMPLATE_DOC_ID).getBody().getText();
  const missing = ['{{HO_TEN}}', '{{MSSV}}', '{{EMAIL}}', '{{SDT}}', '{{VI_TRI}}', '{{CONG_TY}}']
    .filter(function (p) { return text.indexOf(p) === -1; });
  if (missing.length) throw new Error('Doc mẫu thiếu placeholder: ' + missing.join(', '));
  Logger.log('Cấu hình hợp lệ.');
}

/** Chạy thử với dòng dữ liệu cuối cùng (reset trạng thái để gửi lại). */
function testLastRow() {
  const sheet = getResponseSheet_();
  const cols = getColumns_(sheet);
  const row = sheet.getLastRow();
  if (row < 2) throw new Error('Sheet chưa có dòng dữ liệu nào.');
  sheet.getRange(row, cols.STATUS).clearContent();
  processRow_(sheet, row);
}

/** Xử lý lại tất cả dòng chưa gửi hoặc đang báo lỗi. */
function retryPending() {
  const sheet = getResponseSheet_();
  const cols = getColumns_(sheet);
  const last = sheet.getLastRow();
  for (let r = 2; r <= last; r++) {
    const status = String(sheet.getRange(r, cols.STATUS).getValue() || '');
    if (status === '' || status.indexOf('Lỗi') === 0) processRow_(sheet, r);
  }
}
