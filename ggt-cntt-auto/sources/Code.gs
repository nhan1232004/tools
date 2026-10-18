/**
 * ==============================================================================
 * GGT-CNTT AUTO — HỆ THỐNG TỰ ĐỘNG TẠO GIẤY GIỚI THIỆU THỰC TẬP (AI_LAB)
 * Khoa Công Nghệ Thông Tin — Trường Đại học Tôn Đức Thắng
 *
 * Luồng hoạt động:
 * 1. Sinh viên điền thông tin qua Google Form (lưu vào Google Sheet).
 * 2. Tự động bóc tách, chuẩn hóa dữ liệu (hỗ trợ cả số khoa học E+ do Sheet format).
 * 3. Kiểm chứng Mã số thuế (MST) qua API Tổng cục Thuế:
 *    https://api.xinvoice.vn/gdt-api/tax-payer/{mst}
 * 4. Đối chứng so sánh tên công ty sinh viên nhập với tên đăng ký thuế:
 *    - Nhận diện chi nhánh: ưu tiên tên công ty tổng ở phía trước.
 *    - Hỗ trợ đối chiếu tên tiếng Anh, tên viết tắt và từ viết tắt thông dụng.
 *    - Nếu sai MST hoặc tên công ty khác biệt rõ rệt -> trả về không hợp lệ.
 * 5. Điền dữ liệu vào Google Doc mẫu -> Xuất file PDF (đảm bảo đúng 1 trang).
 * 6. Gửi email đính kèm PDF và link Google Drive đến giảng viên phụ trách.
 * 7. Ghi nhận trạng thái xử lý, link PDF và kết quả đối soát MST vào Google Sheet.
 * ==============================================================================
 */

// ==============================================================================
// 1. CẤU HÌNH HỆ THỐNG (NGƯỜI DÙNG THIẾT LẬP TẠI ĐÂY)
// ==============================================================================
const CONFIG = {
  // ID file Google Doc mẫu (tạo từ file templates/GGT_mau_placeholder.docx)
  TEMPLATE_DOC_ID: '1-d1-EzqqMn--gtmMR4dX2fvxXl6Q3zFMb1-JufniQyc', // Giữ cấu hình như bản gốc để không cần nhập lại

  // ID thư mục Google Drive để lưu các file PDF xuất ra
  OUTPUT_FOLDER_ID: '1RgRcem14hvJdTa_Cq7As02RnYS8idFQ6', // Giữ cấu hình như bản gốc để không cần nhập lại

  // Chỉ email giảng viên nếu được bật; hệ thống không gửi email cho sinh viên.
  SEND_EMAIL: false,

  // Mặc định chỉ người có quyền với thư mục Drive mới xem được PDF.
  // Chỉ đổi sang ANYONE_WITH_LINK khi đã được đơn vị cho phép chia sẻ công khai bằng liên kết.
  PDF_SHARING_MODE: 'RESTRICTED',

  // Giới hạn số dòng mỗi lần chạy thủ công để tránh vượt thời gian chạy Apps Script.
  MAX_BATCH_ROWS: 25,
  MAX_QUEUE_ROWS: 500,
  BATCH_RUNTIME_BUDGET_MS: 240000,
  AI_ENABLED: false,
  AI_DICE_MIN: 0.30,
  AI_DICE_MAX: 0.50,
  AI_CONFIDENCE_THRESHOLD: 0.90,
  AUTO_NUMBERING: false,
  REQUIRE_APPROVAL: false,

  // Lọc thời gian nộp đơn: chỉ xử lý các yêu cầu từ thời điểm này tới nay (định dạng YYYY-MM-DD).
  // Mặc định '2026-08-01' = chỉ xử lý các yêu cầu bắt đầu từ đầu tháng 8/2026 tới mới nhất; để trống = xử lý tất cả
  FILTER_START_DATE: '2026-08-01',

  // Email giảng viên nhận PDF (chỉ dùng khi bật SEND_EMAIL = true)
  LECTURER_EMAIL: 'email.giangvien@example.com',

  // Danh sách email CC (phân cách bằng dấu phẩy, để trống nếu không dùng)
  CC_EMAILS: '',

  // Email nhận cảnh báo lỗi hệ thống (để trống = chính tài khoản chạy script)
  ADMIN_EMAIL: '',

  // Tên hiển thị người gửi email
  SENDER_NAME: 'Hệ thống GGT - Khoa CNTT (TDTU)',

  // Tên sheet chứa dữ liệu (để trống = tự động phát hiện sheet phù hợp)
  SHEET_NAME: '',

  // Endpoint API kiểm tra mã số thuế
  TAX_API_URL: 'https://api.xinvoice.vn/gdt-api/tax-payer/',

  // Bật/tắt kiểm tra mã số thuế (true: bắt buộc kiểm tra; false: bỏ qua kiểm tra)
  VERIFY_TAX_CODE: true,

  // Sử dụng tên công ty chính thức từ Tổng cục Thuế trên Giấy giới thiệu (true)
  // hoặc ưu tiên giữ tên do sinh viên nhập nếu đã đối chứng hợp lệ (false)
  USE_OFFICIAL_NAME_ON_DOC: false,

  // Vị trí thực tập mặc định nếu sinh viên để trống
  DEFAULT_VI_TRI: 'Thực tập sinh',

  // Thời gian tạm nghỉ giữa các lần gọi API trong xử lý hàng loạt (tránh lỗi 429 Too Many Requests: giới hạn 10 req/30s)
  BATCH_SLEEP_MS: 3200,

  // Thời gian lưu cache kết quả tra cứu MST (giây) - 6 giờ
  CACHE_EXPIRY_SEC: 21600,

  // Tiêu đề các cột quản lý và đối soát ghi vào Google Sheet
  HEADERS: {
    STATUS: 'TRẠNG THÁI GỬI GGT',
    LINK: 'LINK PDF GGT',
    TIME: 'THỜI GIAN XỬ LÝ GGT',
    DOC_NO: 'SỐ GIẤY GGT',
    TAX_NAME: 'Tên Doanh Nghiệp (Tra cứu MST)',
    TAX_NOTE: 'Kết quả đối soát MST'
  }
};

// ==============================================================================
// 2. TỪ KHÓA NHẬN DIỆN CỘT TRONG SHEET / FORM
// Hỗ trợ tự động nhận diện cả Form 6-7 trường và Sheet mẫu 16 cột của Khoa
// ==============================================================================
const FIELD_KEYWORDS = {
  // Trường bắt buộc
  HO_TEN:   ['họ và tên', 'họ tên', 'ho va ten', 'tên sinh viên'],
  MSSV:     ['mssv', 'mã số sinh viên', 'ma so sinh vien'],
  EMAIL:    ['email cá nhân', 'email ca nhan', 'email'],
  SDT:      ['số điện thoại', 'sđt', 'sdt', 'điện thoại', 'dien thoai'],
  CONG_TY:  ['tên doanh nghiệp', 'tên công ty', 'doanh nghiệp', 'công ty'],
  MST:      ['mã số doanh nghiệp', 'mã số thuế', 'mst', 'ma so doanh nghiep', 'ma so thue'],

  // Trường tùy chọn (có trong Sheet đầy đủ của Khoa)
  TIMESTAMP:['timestamp', 'dấu thời gian', 'thời gian nộp', 'thời gian gửi đơn', 'dau thoi gian', 'thoi gian nop', 'thoi gian gui don', 'ngay nop'],
  VI_TRI:   ['ví trí thực tập', 'vị trí thực tập', 'ví trí', 'vị trí', 'vi tri'],
  HOC_PHAN: ['tên học phần', 'học phần', 'hoc phan'],
  TEN_TA:   ['tiếng nước ngoài', 'foreign name', 'tên tiếng anh'],
  TEN_TAT:  ['viết tắt', 'tên viết tắt', 'abbreviation'],
  NGAY_BD:  ['thời gian bắt đầu', 'ngày bắt đầu', 'ngay bat dau']
};

const REQUIRED_FIELDS = ['HO_TEN', 'MSSV', 'EMAIL', 'SDT', 'CONG_TY'];

// ==============================================================================
// 3. TRIGGER VÀ ĐIỀU PHỐI XỬ LÝ (CORE DISPATCHER)
// ==============================================================================

/**
 * Trigger tự động: chạy mỗi khi có sinh viên nộp Google Form mới.
 */
function handleFormSubmit(e) {
  const lock = LockService.getScriptLock();
  // Khóa script tối đa 30s để xử lý tuần tự khi nhiều sinh viên nộp cùng lúc
  if (!lock.tryLock(30000)) {
    Logger.log('Hệ thống đang bận xử lý yêu cầu khác, thử lại sau.');
    return;
  }

  try {
    if (e && e.range) {
      processRow_(e.range.getSheet(), e.range.getRow());
    } else {
      const sheet = getResponseSheet_();
      if (sheet.getLastRow() >= 2) {
        processRow_(sheet, sheet.getLastRow());
      }
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * Xử lý một dòng dữ liệu cụ thể trong Sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 
 * @param {number} row 
 */
function processRow_(sheet, row, isTestMode, approvedForRelease) {
  const cols = getColumns_(sheet);
  const rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  const currentStatus = String(rowValues[cols.STATUS - 1] || '').trim();
  // Không gửi lại nếu đã hoàn thành hoặc đã được duyệt trước đó trong Sheet
  if (isAlreadyProcessed_(currentStatus)) {
    Logger.log('Dòng ' + row + ' đã được xử lý/phê duyệt trước đó ("' + currentStatus + '"). Bỏ qua.');
    return false;
  }

  // Lọc theo thời gian nộp đơn: chỉ xử lý các yêu cầu bắt đầu từ đầu tháng 8/2026 tới mới nhất
  // (Nếu đang trong chế độ kiểm thử thủ công testRow/testLastRow, bỏ qua lọc ngày để hỗ trợ test mọi dòng)
  if (!isTestMode && CONFIG.FILTER_START_DATE) {
    const rawTimestamp = (cols.TIMESTAMP && cols.TIMESTAMP <= rowValues.length) ? rowValues[cols.TIMESTAMP - 1] : '';
    if (rawTimestamp) {
      const subDate = parseSubmissionDate_(rawTimestamp);
      const filterStart = parseSubmissionDate_(CONFIG.FILTER_START_DATE);
      if (subDate && filterStart && subDate < filterStart) {
        Logger.log('Dòng ' + row + ' nộp ngày ' + Utilities.formatDate(subDate, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss') +
                   ' (trước mốc ' + CONFIG.FILTER_START_DATE + '). Bỏ qua theo yêu cầu chỉ xử lý từ đầu tháng 8/2026.');
        return false;
      }
    }
  }

  let taxResult = null;
  try {
    // Đánh dấu trước để có thể nhận diện và chạy lại nếu execution bị dừng giữa chừng.
    if (cols.STATUS && cols.STATUS > 0) {
      sheet.getRange(row, cols.STATUS).setValue('Processing');
      SpreadsheetApp.flush();
    }

    // 1. Trích xuất dữ liệu thô
    const raw = {};
    Object.keys(FIELD_KEYWORDS).forEach(function (k) {
      const colIdx = cols[k];
      raw[k] = (colIdx && colIdx <= rowValues.length) ? rowValues[colIdx - 1] : '';
    });

    // 2. Chuẩn hóa & kiểm tra tính hợp lệ cơ bản
    const studentData = cleanAndValidate_(raw);

    // 3. Kiểm chứng Mã số thuế & đối chiếu tên công ty (theo yêu cầu bắt buộc)
    if (CONFIG.VERIFY_TAX_CODE) {
      taxResult = verifyTaxAndCompany_(studentData);
      studentData.taxOfficialName = taxResult.name;
      studentData.taxAddress = taxResult.address;
      studentData.taxStatus = taxResult.status;

      // Nếu cấu hình dùng tên chính thức theo MST trên giấy giới thiệu
      if (CONFIG.USE_OFFICIAL_NAME_ON_DOC && taxResult.name) {
        studentData.congTy = taxResult.name;
      }
    }

    if (CONFIG.REQUIRE_APPROVAL && !isTestMode && !approvedForRelease) {
      writePendingApproval_(sheet, row, cols, taxResult);
      Logger.log('Dòng ' + row + ' hợp lệ và đang chờ người phụ trách duyệt.');
      return true;
    }

    // Chế độ kiểm thử không cấp số văn bản thật.
    studentData.soGgt = (CONFIG.AUTO_NUMBERING && !isTestMode) ? nextDocumentNumber_() : '';

    // 4. Điền mẫu Google Doc & xuất file PDF (lưu vào Drive)
    const pdf = createPdf_(studentData);

    // 5. Gửi email kèm file PDF cho giảng viên (chỉ thực hiện nếu bật SEND_EMAIL = true)
    if (CONFIG.SEND_EMAIL) {
      sendMail_(studentData, pdf.blob, pdf.url);
    } else {
      Logger.log('Đã tạo file PDF thành công. Bỏ qua gửi email theo yêu cầu (SEND_EMAIL = false).');
    }

    // 6. Ghi nhận kết quả thành công và link PDF vào Sheet
    writeSuccess_(sheet, row, cols, pdf.url, taxResult, studentData.soGgt || '');
    Logger.log('Đã xử lý thành công dòng ' + row + ' cho MSSV: ' + studentData.mssv + ' | Link PDF: ' + pdf.url);
    return true;

  } catch (err) {
    Logger.log('Lỗi tại dòng ' + row + ': ' + err.message);
    const taxRes = taxResult || (err.taxInfo ? { name: err.taxInfo.name } : null);
    writeError_(sheet, row, cols, err.message, taxRes);
    notifyAdmin_(row, err);
    return false;
  }
}

// ==============================================================================
// 4. CHUẨN HÓA VÀ KIỂM TRA DỮ LIỆU ĐẦU VÀO
// ==============================================================================

/**
 * Chuẩn hóa các trường thông tin của sinh viên.
 * Giải quyết triệt để lỗi số khoa học (E+) do Excel/Sheets tự định dạng.
 */
function cleanAndValidate_(raw) {
  const errors = [];

  const hoTen = titleCase_(collapse_(raw.HO_TEN));
  const mssv = normalizeMssv_(raw.MSSV);
  const email = collapse_(raw.EMAIL).toLowerCase();
  const congTy = collapse_(raw.CONG_TY);
  const mst = normalizeMst_(raw.MST);
  const sdt = normalizePhone_(raw.SDT);

  // Vị trí thực tập: nếu để trống thì lấy mặc định hoặc từ Tên học phần
  let viTri = collapse_(raw.VI_TRI);
  if (!viTri) {
    viTri = collapse_(raw.HOC_PHAN) ? ('Sinh viên thực tập (' + collapse_(raw.HOC_PHAN) + ')') : CONFIG.DEFAULT_VI_TRI;
  }

  // Tùy chọn mở rộng
  const hocPhan = collapse_(raw.HOC_PHAN);
  const tenTiengAnh = collapse_(raw.TEN_TA);
  const tenVietTat = collapse_(raw.TEN_TAT);

  // Kiểm tra tính hợp lệ
  if (!hoTen) errors.push('thiếu Họ và tên');
  if (!/^[A-Z0-9]{6,12}$/.test(mssv)) errors.push('MSSV không hợp lệ (' + mssv + ')');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) errors.push('Email không hợp lệ (' + email + ')');
  if (!sdt || !/^0\d{9}$/.test(sdt)) errors.push('Số điện thoại không hợp lệ (' + sdt + ')');
  if (!congTy) errors.push('thiếu Tên công ty');

  if (CONFIG.VERIFY_TAX_CODE) {
    if (!mst) {
      errors.push('thiếu Mã số thuế công ty');
    } else if (!/^[0-9]{10}(-[0-9]{3})?$/.test(mst) && !/^[0-9]{13}$/.test(mst)) {
      errors.push('mã số thuế không hợp lệ (' + mst + ')');
    }
  }

  if (errors.length > 0) {
    throw new Error('Dữ liệu không hợp lệ: ' + errors.join('; '));
  }

  return {
    hoTen: hoTen,
    mssv: mssv,
    email: email,
    sdt: sdt,
    congTy: congTy,
    mst: mst,
    viTri: viTri,
    hocPhan: hocPhan,
    tenTiengAnh: tenTiengAnh,
    tenVietTat: tenVietTat
  };
}

/**
 * Chuẩn hóa MSSV, chuyển đổi định dạng số khoa học (e.g. 5.1900811E7 -> 51900811).
 */
function normalizeMssv_(val) {
  if (val === undefined || val === null || val === '') return '';
  if (typeof val === 'number') {
    return Math.round(val).toString();
  }
  const s = String(val).trim();
  if (/^[0-9]+(\.[0-9]+)?[eE]\+?[0-9]+$/.test(s)) {
    const num = Number(s);
    if (!isNaN(num)) return Math.round(num).toString();
  }
  return s.replace(/\s+/g, '').toUpperCase();
}

/**
 * Chuẩn hóa Mã số thuế công ty:
 * - Khắc phục số khoa học (ví dụ 1.100109154E9 -> 1100109154).
 * - Bù số 0 ở đầu nếu bị bảng tính làm mất (9 chữ số -> 10 chữ số).
 * - Loại bỏ khoảng trắng và dấu chấm phân cách.
 */
function normalizeMst_(val) {
  if (val === undefined || val === null || val === '') return '';
  if (typeof val === 'number') {
    let numStr = Math.round(val).toString();
    if (numStr.length === 9) numStr = '0' + numStr;
    return numStr;
  }

  let s = String(val).trim();
  if (/^[0-9]+(\.[0-9]+)?[eE]\+?[0-9]+$/.test(s)) {
    const num = Number(s);
    if (!isNaN(num)) s = Math.round(num).toString();
  }

  s = s.replace(/[\s.]/g, '');
  // Nếu bị mất số 0 ở đầu do format số 9 chữ số
  if (/^[1-9]\d{8}$/.test(s)) {
    s = '0' + s;
  }
  return s;
}

/**
 * Chuẩn hóa số điện thoại Việt Nam:
 * - Chuyển đổi số khoa học E+ (ví dụ: 7.75380223E8 -> 775380223 -> 0775380223).
 * - Chuyển +84 hoặc 84 về đầu 0.
 * - Bù số 0 nếu bị định dạng số làm mất 9 chữ số.
 */
function normalizePhone_(val) {
  if (val === undefined || val === null || val === '') return '';
  if (typeof val === 'number') {
    let numStr = Math.round(val).toString();
    if (numStr.length === 9) numStr = '0' + numStr;
    return numStr;
  }
  let s = String(val).trim();
  // Khắc phục số khoa học do Excel/Sheets format (ví dụ 7.75380223E8)
  if (/^[0-9]+(\.[0-9]+)?[eE]\+?[0-9]+$/.test(s)) {
    const num = Number(s);
    if (!isNaN(num)) s = Math.round(num).toString();
  }
  s = s.replace(/[\s.\-()]/g, '');
  if (s.indexOf('+84') === 0) s = '0' + s.slice(3);
  else if (s.indexOf('84') === 0 && s.length === 11) s = '0' + s.slice(2);
  else if (/^[1-9]\d{8}$/.test(s)) s = '0' + s;
  return s;
}

/**
 * Phân tích và chuyển đổi thời gian nộp đơn từ nhiều dạng format khác nhau:
 * - Đối tượng Date gốc trong Google Sheets
 * - Số ngày Excel serial (ví dụ 46240.355)
 * - Chuỗi ngày tháng dạng: 'dd/MM/yyyy HH:mm:ss', 'dd/MM/yyyy', 'yyyy-MM-dd'
 */
function parseSubmissionDate_(val) {
  if (val === undefined || val === null || val === '') return null;
  if (val instanceof Date) return val;
  if (typeof val === 'number') {
    // Chuyển đổi Excel serial date sang JavaScript Date
    return new Date((val - 25569) * 86400 * 1000);
  }
  const s = String(val).trim();
  // Khắc phục số serial dạng chuỗi: "46240.355"
  if (/^[0-9]+(\.[0-9]+)?$/.test(s)) {
    const num = Number(s);
    if (!isNaN(num) && num > 30000 && num < 70000) {
      return new Date((num - 25569) * 86400 * 1000);
    }
  }
  const dmy = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/);
  if (dmy) {
    return new Date(parseInt(dmy[3], 10), parseInt(dmy[2], 10) - 1, parseInt(dmy[1], 10),
                    parseInt(dmy[4] || 0, 10), parseInt(dmy[5] || 0, 10), parseInt(dmy[6] || 0, 10));
  }
  const ymd = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (ymd) {
    return new Date(parseInt(ymd[1], 10), parseInt(ymd[2], 10) - 1, parseInt(ymd[3], 10));
  }
  const parsed = new Date(s);
  return isNaN(parsed.getTime()) ? null : parsed;
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
  if (!s) return '';
  return s.normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .replace(/đ/g, 'd')
          .replace(/Đ/g, 'D');
}

// ==============================================================================
// 5. KIỂM CHỨNG MÃ SỐ THUẾ & ĐỐI CHỨNG SO SÁNH TÊN DOANH NGHIỆP
// Thực thi API https://api.xinvoice.vn/gdt-api/tax-payer/
// ==============================================================================

/**
 * Kiểm tra MST qua API và đối chứng so sánh với tên do sinh viên nhập.
 * @param {Object} data 
 * @returns {Object} Thông tin chi tiết từ cơ quan thuế
 */
function verifyTaxAndCompany_(data) {
  const taxInfo = fetchTaxPayerInfo_(data.mst);

  if (!taxInfo) {
    throw new Error('mã số thuế không hợp lệ hoặc không tồn tại trên hệ thống Thuế (' + data.mst + ')');
  }

  // Kiểm tra tình trạng hoạt động của người nộp thuế
  const statusStr = String(taxInfo.status || '');
  if (statusStr.indexOf('ngừng hoạt động') !== -1 || statusStr.indexOf('đã đóng') !== -1) {
    const err = new Error('mã số thuế thuộc doanh nghiệp đã ngừng hoạt động (' + statusStr + ')');
    err.taxInfo = taxInfo;
    throw err;
  }

  // Tiến hành đối chiếu so sánh tên
  const matchResult = compareCompanyNames_(
    data.congTy,
    taxInfo.name,
    taxInfo.orgType,
    data.tenTiengAnh,
    data.tenVietTat
  );

  if (!matchResult.isMatch) {
    const err = new Error('mã số thuế hoặc tên công ty không hợp lệ (MST: ' + data.mst + ' đăng ký là "' + taxInfo.name + '", không tương đồng với "' + data.congTy + '")');
    err.taxInfo = taxInfo;
    throw err;
  }

  return {
    name: taxInfo.name,
    address: taxInfo.address || '',
    status: taxInfo.status || 'NNT đang hoạt động',
    orgType: taxInfo.orgType || '',
    matchNote: matchResult.reason
  };
}

/**
 * Gọi API https://api.xinvoice.vn/gdt-api/tax-payer/{mst} có hỗ trợ Cache và Retry backoff
 */
function fetchTaxPayerInfo_(mst) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'TAX_' + mst;
  const cached = cache.get(cacheKey);

  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (ignore) {}
  }

  const url = CONFIG.TAX_API_URL + encodeURIComponent(mst);
  let lastError = null;

  // Thử gọi tối đa 3 lần nếu gặp nghẽn mạng hoặc giới hạn tần suất 429
  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      const response = UrlFetchApp.fetch(url, {
        method: 'get',
        muteHttpExceptions: true,
        headers: {
          'Accept': 'application/json',
          'User-Agent': 'TDTU-GGT-Auto/2.0'
        }
      });

      const statusCode = response.getResponseCode();

      if (statusCode === 200) {
        const bodyText = response.getContentText();
        const json = JSON.parse(bodyText);
        // Lưu vào cache 6 giờ
        cache.put(cacheKey, JSON.stringify(json), CONFIG.CACHE_EXPIRY_SEC);
        return json;
      }

      if (statusCode === 404) {
        // Mã số thuế không tồn tại
        return null;
      }

      if (statusCode === 429) {
        // Bị giới hạn tốc độ truy vấn (10 req/30s) -> đọc Retry-After hoặc chờ tối thiểu 6-12s
        let waitSec = 6 * attempt;
        try {
          const respHeaders = response.getHeaders();
          const ra = respHeaders['Retry-After'] || respHeaders['retry-after'];
          if (ra) {
            const parsed = parseInt(ra, 10);
            if (!isNaN(parsed) && parsed > 0) waitSec = Math.max(parsed, waitSec);
          }
        } catch (ignore) {}
        Logger.log('Gặp giới hạn tần suất API (HTTP 429) cho MST ' + mst + '. Chờ ' + waitSec + ' giây trước khi thử lại...');
        Utilities.sleep(waitSec * 1000);
        continue;
      }

      lastError = new Error('API thuế phản hồi mã ' + statusCode);
    } catch (err) {
      lastError = err;
      Utilities.sleep(1000 * attempt);
    }
  }

  throw new Error('Không thể kết nối đến hệ thống kiểm tra MST: ' + (lastError ? lastError.message : 'Timeout'));
}

function judgeCompanyNames_(studentName, taxName, foreignName, abbreviation) {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('GEMINI_API_KEY');
  const model = props.getProperty('GEMINI_MODEL');
  if (!apiKey || !model) throw new Error('Chưa cấu hình GEMINI_API_KEY/GEMINI_MODEL trong Script Properties.');
  const prompt = [
    'So sánh tên doanh nghiệp để xác định có cùng một tổ chức hay không.',
    'Chỉ dùng tên được cung cấp; không suy đoán. Trả về JSON: {"match":boolean,"confidence":number 0-1,"reason":string}.',
    'Tên khai báo: ' + String(studentName || '').slice(0, 250),
    'Tên đăng ký thuế: ' + String(taxName || '').slice(0, 250),
    'Tên nước ngoài: ' + String(foreignName || '').slice(0, 250),
    'Tên viết tắt: ' + String(abbreviation || '').slice(0, 120)
  ].join('\n');
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent';
  const payload = {contents: [{parts: [{text: prompt}]}], generationConfig: {temperature: 0, responseMimeType: 'application/json'}};
  let lastError = '';
  for (let attempt = 1; attempt <= 3; attempt++) {
    const response = UrlFetchApp.fetch(url, {method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: {'x-goog-api-key': apiKey}, payload: JSON.stringify(payload)});
    const code = response.getResponseCode();
    if (code === 200) {
      const data = JSON.parse(response.getContentText());
      const text = data.candidates[0].content.parts.map(function (part) { return part.text || ''; }).join('');
      const result = JSON.parse(text);
      if (typeof result.match !== 'boolean' || !isFinite(Number(result.confidence))) throw new Error('AI trả về JSON sai cấu trúc.');
      return {match: result.match, confidence: Math.max(0, Math.min(1, Number(result.confidence))), reason: String(result.reason || '').slice(0, 300)};
    }
    lastError = 'Gemini HTTP ' + code;
    if (code !== 429 && code < 500) break;
    Utilities.sleep(2000 * attempt);
  }
  throw new Error(lastError || 'Gemini không phản hồi');
}

/**
 * Thuật toán đối chứng so sánh tên công ty:
 * - Chuẩn hóa các từ viết tắt doanh nghiệp thông dụng (CTY, TNHH, MTV, CP, JSC, LTD, CORP).
 * - Yêu cầu bài toán: "nếu công ty là chi nhánh thì chỉ cần phần tên công ty tổng ở phía trước
 *   ưu tiên chính xác là được, phần chi nhánh có thể khác đôi chút".
 * - Bóc tách từ khóa cốt lõi (loại bỏ stopwords chung và giảm trọng số từ ngành nghề chung).
 * - Đối chiếu cả tên tiếng nước ngoài và tên viết tắt nếu có.
 */
function compareCompanyNames_(studentName, apiName, orgType, foreignName, abbrev) {
  const normSv = normalizeCorporateText_(studentName);
  const normApi = normalizeCorporateText_(apiName);

  // 1. Trường hợp trùng khớp chính xác hoặc chứa toàn bộ
  if (normSv === normApi || normApi.indexOf(normSv) !== -1 || normSv.indexOf(normApi) !== -1) {
    return { isMatch: true, reason: 'Khớp chính xác tên doanh nghiệp' };
  }

  // 2. Xử lý yêu cầu chi nhánh (Branch Logic)
  const isBranch = (orgType && orgType.toLowerCase().indexOf('chi nhánh') !== -1) ||
                   normApi.indexOf('chi nhanh') !== -1 ||
                   normSv.indexOf('chi nhanh') !== -1;

  if (isBranch) {
    // Tách phần tên công ty tổng ở phía trước (trước dấu gạch ngang '-', hoặc trước chữ 'chi nhánh')
    const parentApi = extractParentCompanyName_(apiName);
    const parentSv = extractParentCompanyName_(studentName);

    const normParentApi = normalizeCorporateText_(parentApi);
    const normParentSv = normalizeCorporateText_(parentSv);

    if (normParentApi && normParentSv) {
      if (normParentApi.indexOf(normParentSv) !== -1 || normParentSv.indexOf(normParentApi) !== -1) {
        return { isMatch: true, reason: 'Khớp tên công ty tổng của chi nhánh' };
      }
    }

    // So sánh từ khóa trọng số phần công ty tổng
    const tokParentApi = getWeightedTokens_(parentApi);
    const tokParentSv = getWeightedTokens_(parentSv);
    const simParent = computeWeightedSimilarity_(tokParentSv, tokParentApi);

    if (simParent.dice >= 0.5 || (simParent.inclusion1 >= 0.7 && simParent.common >= 1.0)) {
      const pct = Math.round(Math.max(simParent.dice, simParent.inclusion1) * 100);
      return { isMatch: true, reason: 'Khớp phần công ty tổng (độ tương đồng ' + pct + '%)' };
    }
  }

  // 3. Đối chiếu từ khóa cốt lõi giữa tên sinh viên và tên từ Thuế
  const tokSv = getWeightedTokens_(studentName);
  const tokApi = getWeightedTokens_(apiName);
  const directSim = computeWeightedSimilarity_(tokSv, tokApi);

  // Hợp lệ nếu độ tương đồng Dice >= 50% HOẶC tên sinh viên nhập được bao hàm >= 70% trong tên chính thức (với ít nhất 1 từ định danh trọng số cao)
  if (directSim.dice >= 0.5 || (directSim.inclusion1 >= 0.7 && directSim.common >= 1.0)) {
    const pct = Math.round(Math.max(directSim.dice, directSim.inclusion1) * 100);
    return { isMatch: true, reason: 'Tên công ty tương đồng cao (' + pct + '%)' };
  }

  // 4. Kiểm tra thêm đối chiếu tên Tiếng Anh / Tên Nước Ngoài (nếu có)
  if (foreignName) {
    const tokForeign = getWeightedTokens_(foreignName);
    const simForeign = computeWeightedSimilarity_(tokForeign, tokApi);
    if (simForeign.dice >= 0.5 || (simForeign.inclusion1 >= 0.7 && simForeign.common >= 1.0)) {
      return { isMatch: true, reason: 'Khớp với tên doanh nghiệp viết bằng tiếng nước ngoài' };
    }
  }

  // 5. Kiểm tra tên viết tắt (Ví dụ ISD, TMA, ITD, FPT, VNG, BOSCH, CEP...)
  if (abbrev) {
    const abbrevWords = normalizeCorporateText_(abbrev).split(' ');
    for (let i = 0; i < abbrevWords.length; i++) {
      const w = abbrevWords[i];
      if (w.length >= 2 && LEGAL_STOPWORDS.indexOf(w) === -1) {
        if (normApi.indexOf(w) !== -1 || normSv.indexOf(w) !== -1) {
          return { isMatch: true, reason: 'Khớp tên viết tắt thương hiệu: ' + w.toUpperCase() };
        }
      }
    }
  }

  // AI chỉ xử lý vùng xám và chỉ nhận tên doanh nghiệp, không nhận dữ liệu sinh viên.
  if (CONFIG.AI_ENABLED && directSim.dice >= CONFIG.AI_DICE_MIN && directSim.dice <= CONFIG.AI_DICE_MAX) {
    try {
      const aiResult = judgeCompanyNames_(studentName, apiName, foreignName, abbrev);
      if (aiResult.match === true && Number(aiResult.confidence) >= CONFIG.AI_CONFIDENCE_THRESHOLD) {
        return { isMatch: true, reason: 'AI hỗ trợ đối soát vùng xám (' + Math.round(aiResult.confidence * 100) + '%): ' + aiResult.reason };
      }
      return { isMatch: false, reason: 'AI không xác nhận đủ ngưỡng tin cậy: ' + (aiResult.reason || 'cần kiểm tra thủ công') };
    } catch (aiError) { Logger.log('AI không khả dụng, quay về kết quả luật: ' + aiError.message); }
  }
  return { isMatch: false, reason: 'Không tương đồng (điểm số: ' + Math.round(directSim.dice * 100) + '%)' };
}

/**
 * Trích xuất phần tên công ty tổng từ tên chi nhánh.
 */
function extractParentCompanyName_(name) {
  if (!name) return '';
  let clean = name;

  // Xóa phần chú thích phân loại doanh nghiệp trong ngoặc, ví dụ: (LOẠI HÌNH DOANH NGHIỆP: CÔNG TY TNHH)
  clean = clean.replace(/\(loại\s*hình\s*doanh\s*nghiệp[^)]*\)/gi, ' ');

  // Nếu có ngoặc đơn chứa thông tin chi nhánh / trung tâm -> loại bỏ phần trong ngoặc
  clean = clean.replace(/\([^)]*chi\s*nhánh[^)]*\)/gi, ' ')
               .replace(/\([^)]*trung\s*tâm[^)]*\)/gi, ' ');

  // Tách theo dấu gạch ngang '-' hoặc '–' hoặc '—'
  const parts = clean.split(/[-–—]/);
  if (parts.length >= 2) {
    let firstPart = parts[0].trim();
    // Bỏ tiền tố "Chi nhánh" ở đầu nếu có
    firstPart = firstPart.replace(/^chi\s*nhánh\s+/i, '').trim();
    if (firstPart.length > 5) return firstPart;
  }

  // Mẫu: "Công ty ABC chi nhánh ..."
  const cnIdx = clean.toLowerCase().indexOf('chi nhánh');
  if (cnIdx > 5) {
    return clean.slice(0, cnIdx).trim();
  } else if (cnIdx === 0) {
    return clean.slice(9).trim();
  }

  return clean.trim();
}

/**
 * Chuẩn hóa tên doanh nghiệp: chữ thường, bỏ dấu, thay thế từ viết tắt.
 */
function normalizeCorporateText_(str) {
  if (!str) return '';
  let s = removeAccents_(str).toLowerCase();

  // Mở rộng các từ viết tắt pháp lý doanh nghiệp
  s = s.replace(/\bcty\b|\bct\b/g, 'cong ty')
       .replace(/\btnhh\b/g, 'trach nhiem huu han')
       .replace(/\bmtv\b/g, 'mot thanh vien')
       .replace(/\bcp\b|\bjsc\b/g, 'co phan')
       .replace(/\bco\.\,?\s*ltd\b|\bltd\b/g, 'trach nhiem huu han')
       .replace(/\bcorp\b|\bcorporation\b/g, 'tap doan')
       .replace(/\bcn\b/g, 'chi nhanh')
       .replace(/\bvn\b/g, 'viet nam');

  // Xóa các ký tự đặc biệt, chỉ giữ lại chữ cái, số và khoảng trắng
  s = s.replace(/[^a-z0-9\s]/g, ' ');
  return s.replace(/\s+/g, ' ').trim();
}

// Từ khóa pháp lý chung không mang tính định danh thương hiệu
const LEGAL_STOPWORDS = [
  'cong', 'ty', 'trach', 'nhiem', 'huu', 'han', 'mot', 'thanh', 'vien',
  'co', 'phan', 'tap', 'doan', 'chi', 'nhanh', 'viet', 'nam', 'va', 'cua', 'tai',
  'company', 'limited', 'corporation', 'corp', 'inc', 'group', 'jsc', 'llc',
  'loai', 'hinh', 'doanh', 'nghiep', 'vp', 'van', 'phong'
];

// Từ khóa mô tả ngành nghề chung (cho trọng số thấp hơn định danh riêng)
const DOMAIN_KEYWORDS = [
  'phan', 'mem', 'cong', 'nghe', 'thong', 'tin', 'giai', 'phap',
  'dich', 'vu', 'thuong', 'mai', 'tu', 'van', 'dau', 'tu', 'san', 'xuat',
  'truyen', 'thong', 'quoc', 'te', 'toan', 'cau', 'ngan', 'hang', 'tai', 'chinh',
  'software', 'technology', 'solutions', 'services', 'trading', 'consulting'
];

/**
 * Tách danh sách token có gán trọng số:
 * - Tên thương hiệu riêng biệt (vd: Tường Minh, FPT, SiGlaz): trọng số 1.0.
 * - Từ ngành nghề chung (vd: phần mềm, công nghệ): trọng số 0.35.
 */
function getWeightedTokens_(str) {
  const norm = normalizeCorporateText_(str);
  const words = norm.split(' ');
  const tokens = {};

  words.forEach(function (w) {
    if (w.length <= 1 || LEGAL_STOPWORDS.indexOf(w) !== -1) return;
    if (DOMAIN_KEYWORDS.indexOf(w) !== -1) {
      tokens[w] = 0.35;
    } else {
      tokens[w] = 1.0;
    }
  });

  return tokens;
}

/**
 * Tính toán độ tương đồng giữa hai tập token có trọng số (Weighted Dice-Sorensen và Inclusion Ratio)
 */
function computeWeightedSimilarity_(tok1, tok2) {
  let commonWeight = 0.0;
  let weight1 = 0.0;
  let weight2 = 0.0;

  Object.keys(tok1).forEach(function (k) {
    weight1 += tok1[k];
    if (tok2[k] !== undefined) {
      commonWeight += Math.min(tok1[k], tok2[k]);
    }
  });

  Object.keys(tok2).forEach(function (k) {
    weight2 += tok2[k];
  });

  if (weight1 + weight2 === 0) return { dice: 0, inclusion1: 0, common: 0 };
  const dice = (2.0 * commonWeight) / (weight1 + weight2);
  const inclusion1 = weight1 > 0 ? (commonWeight / weight1) : 0;
  return { dice: dice, inclusion1: inclusion1, common: commonWeight, weight1: weight1, weight2: weight2 };
}

// ==============================================================================
// 6. ĐIỀN MẪU GOOGLE DOC VÀ XUẤT FILE PDF
// ==============================================================================

/**
 * Tạo file PDF từ mẫu Google Doc với dữ liệu của sinh viên.
 */
function nextDocumentNumber_() {
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const props = PropertiesService.getScriptProperties();
    const year = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy');
    const key = 'GGT_SEQUENCE_' + year;
    const next = Number(props.getProperty(key) || 0) + 1;
    props.setProperty(key, String(next));
    return Utilities.formatString('%04d-GGT-CNTT/%s', next, year);
  } finally { lock.releaseLock(); }
}

function createPdf_(data) {
  const template = DriveApp.getFileById(getPrivateConfig_('TEMPLATE_DOC_ID'));
  const folder = DriveApp.getFolderById(getPrivateConfig_('OUTPUT_FOLDER_ID'));

  const cleanName = removeAccents_(data.hoTen).replace(/[^A-Za-z0-9]/g, '');
  const baseName = 'GGT_' + data.mssv + '_' + cleanName;
  const tmpCopy = template.makeCopy(baseName + '_tmp', folder);

  try {
    const doc = DocumentApp.openById(tmpCopy.getId());
    const body = doc.getBody();

    // Chuẩn hóa khoảng trắng sau dấu hai chấm nếu mẫu chưa có dấu cách
    body.replaceText(':\\{\\{', ': {{');

    // Bản đồ thay thế placeholder
    const now = new Date();
    const tz = Session.getScriptTimeZone();
    const map = {
      '{{HO_TEN}}': data.hoTen,
      '{{MSSV}}': data.mssv,
      '{{EMAIL}}': data.email,
      '{{SDT}}': data.sdt,
      '{{VI_TRI}}': data.viTri,
      '{{CONG_TY}}': data.congTy,
      '{{SO_GGT}}': data.soGgt || '',
      '{{MST}}': data.mst || '',
      '{{MA_SO_THUE}}': data.mst || '',
      '{{HOC_PHAN}}': data.hocPhan || '',
      '{{DIA_CHI_CTY}}': data.taxAddress || '',
      '{{NGAY}}': Utilities.formatDate(now, tz, 'dd'),
      '{{THANG}}': Utilities.formatDate(now, tz, 'MM'),
      '{{NAM}}': Utilities.formatDate(now, tz, 'yyyy')
    };

    Object.keys(map).forEach(function (key) {
      body.replaceText(escapeRegex_(key), escapeReplacement_(map[key]));
    });

    // Dọn dẹp khoảng trống cuối tài liệu để PDF luôn vừa khít 1 trang
    trimTrailingBlank_(body);

    doc.saveAndClose();
    // Chờ Drive đồng bộ xong trước khi chuyển đổi sang PDF
    Utilities.sleep(1500);

    const blob = tmpCopy.getAs(MimeType.PDF).setName(baseName + '.pdf');
    const pdfFile = folder.createFile(blob);

    // Chỉ bật quyền xem bằng liên kết khi cấu hình được chủ động thay đổi.
    if (CONFIG.PDF_SHARING_MODE === 'ANYONE_WITH_LINK') {
      try {
        pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (sharingError) {
        pdfFile.setTrashed(true);
        throw new Error('Không thể cấp quyền xem PDF theo cấu hình chia sẻ: ' + sharingError.message);
      }
    }

    return {
      blob: pdfFile.getBlob(),
      url: pdfFile.getUrl(),
      fileId: pdfFile.getId()
    };
  } finally {
    // Luôn xóa file Google Doc tạm, chỉ lưu file PDF thành phẩm
    tmpCopy.setTrashed(true);
  }
}

/**
 * Dọn dẹp cuối tài liệu Google Doc để PDF không bị tràn sang trang 2.
 * CHÚ Ý: Chỉ xóa page break và các dòng trắng THỪA phía SAU phần nội dung cuối cùng
 * (phần gây tràn trang), KHÔNG xóa các dòng trắng nằm giữa nội dung
 * (ví dụ: khoảng trống để ký tên giữa chữ ký và phần đóng dấu).
 *
 * Thuật toán:
 * 1. Tìm vị trí đoạn văn có nội dung cuối cùng (last non-blank paragraph).
 * 2. Chỉ xóa các đoạn trắng phía SAU đoạn có nội dung cuối cùng đó,
 *    nhưng giữ lại tối thiểu một số dòng trắng hợp lý (để không phá vỡ bố cục ký tên).
 * 3. Luôn xóa mọi page break để tránh PDF bị 2 trang.
 */
function trimTrailingBlank_(body) {
  const PARA = DocumentApp.ElementType.PARAGRAPH;
  const numChildren = body.getNumChildren();

  // Bước 1: Xóa tất cả page break trong toàn bộ tài liệu (ngăn PDF tràn trang 2)
  for (let i = numChildren - 1; i >= 0; i--) {
    const child = body.getChild(i);
    if (child.getType() !== PARA) continue;
    const p = child.asParagraph();
    for (let j = p.getNumChildren() - 1; j >= 0; j--) {
      if (p.getChild(j).getType() === DocumentApp.ElementType.PAGE_BREAK) {
        p.getChild(j).removeFromParent();
      }
    }
  }

  // Bước 2: Tìm vị trí đoạn có nội dung cuối cùng (chữ, hình ảnh, đường kẻ)
  let lastContentIdx = -1;
  for (let i = body.getNumChildren() - 1; i >= 0; i--) {
    const child = body.getChild(i);
    if (child.getType() !== PARA || !isBlankParagraph_(child.asParagraph())) {
      lastContentIdx = i;
      break;
    }
  }

  if (lastContentIdx === -1) return; // Toàn bộ tài liệu trống

  // Bước 3: Đếm số dòng trắng phía sau đoạn có nội dung cuối cùng
  const totalChildren = body.getNumChildren();
  const trailingBlanks = totalChildren - 1 - lastContentIdx;

  // Giữ lại tối đa 2 dòng trắng phía sau nội dung cuối (đủ cho bố cục ký tên)
  // Chỉ xóa các dòng trắng thừa vượt quá 2 dòng
  const MAX_TRAILING_BLANKS = 2;
  const blanksToRemove = Math.max(0, trailingBlanks - MAX_TRAILING_BLANKS);

  if (blanksToRemove > 0) {
    // Xóa từ cuối lên, chỉ xóa phần thừa
    for (let i = 0; i < blanksToRemove; i++) {
      const lastIdx = body.getNumChildren() - 2; // -2 vì Google Doc luôn giữ phần tử cuối
      if (lastIdx > lastContentIdx) {
        const child = body.getChild(lastIdx);
        if (child.getType() === PARA && isBlankParagraph_(child.asParagraph())) {
          child.removeFromParent();
        }
      }
    }
  }

  // Bước 4: Thu nhỏ đoạn trắng cuối cùng (phần tử bắt buộc của Google Doc) để không chiếm thêm diện tích
  const last = body.getChild(body.getNumChildren() - 1);
  if (last.getType() === PARA) {
    const p = last.asParagraph();
    p.setSpacingAfter(0);
    if (isBlankParagraph_(p)) {
      p.setFontSize(1).setSpacingBefore(0).setLineSpacing(1);
    }
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

// ==============================================================================
// 7. GỬI EMAIL THÔNG BÁO CHO GIẢNG VIÊN KÈM FILE PDF
// ==============================================================================

/**
 * Gửi email kèm file PDF và bảng thông tin chi tiết đến giảng viên.
 */
function sendMail_(data, pdfBlob, pdfUrl) {
  const subject = '[GGT-CNTT] Giấy giới thiệu thực tập - ' + data.hoTen + ' (' + data.mssv + ')';

  const plain =
    'Kính gửi Thầy/Cô,\n\n' +
    'Hệ thống tự động gửi Giấy giới thiệu thực tập của sinh viên sau khi kiểm chứng hợp lệ:\n' +
    '- Họ và tên: ' + data.hoTen + '\n' +
    '- MSSV: ' + data.mssv + '\n' +
    '- Email cá nhân: ' + data.email + '\n' +
    '- Số điện thoại: ' + data.sdt + '\n' +
    '- Học phần: ' + (data.hocPhan || 'KTCN / TSNN') + '\n' +
    '- Vị trí thực tập: ' + data.viTri + '\n' +
    '- Doanh nghiệp: ' + data.congTy + '\n' +
    '- Mã số thuế: ' + (data.mst || 'Không có') + '\n' +
    (data.taxOfficialName ? ('- Tên đăng ký Thuế: ' + data.taxOfficialName + '\n') : '') +
    (data.taxAddress ? ('- Địa chỉ doanh nghiệp: ' + data.taxAddress + '\n') : '') +
    '\nFile PDF Giấy giới thiệu được đính kèm theo thư này.\n' +
    'Xem trực tiếp trên Google Drive: ' + pdfUrl + '\n\n' +
    'Trân trọng,\n' + CONFIG.SENDER_NAME;

  const html =
    '<div style="font-family:Arial,sans-serif;line-height:1.6;color:#333;">' +
    '<h3 style="color:#004085;margin-bottom:8px;">Kính gửi Thầy/Cô,</h3>' +
    '<p>Hệ thống Khoa CNTT đã tự động bóc tách, đối soát Mã số thuế và tạo Giấy giới thiệu thực tập cho sinh viên:</p>' +
    '<table style="border-collapse:collapse;width:100%;max-width:650px;margin:15px 0;background:#fdfdfd;border:1px solid #dee2e6;">' +
    row_('Họ và tên sinh viên', data.hoTen) +
    row_('Mã số sinh viên (MSSV)', data.mssv) +
    row_('Email cá nhân', data.email) +
    row_('Số điện thoại', data.sdt) +
    row_('Học phần', data.hocPhan || 'Kiến tập công nghiệp / Tập sự nghề nghiệp') +
    row_('Vị trí thực tập', data.viTri) +
    row_('Doanh nghiệp đăng ký', data.congTy) +
    row_('Mã số thuế doanh nghiệp', data.mst ? ('<code>' + escapeHtml_(data.mst) + '</code>') : 'Chưa có') +
    (data.taxOfficialName ? row_('Tên đăng ký theo MST', '<b>' + escapeHtml_(data.taxOfficialName) + '</b>') : '') +
    (data.taxAddress ? row_('Địa chỉ doanh nghiệp (MST)', escapeHtml_(data.taxAddress)) : '') +
    (data.taxStatus ? row_('Tình trạng hoạt động', '<span style="color:green;">' + escapeHtml_(data.taxStatus) + '</span>') : '') +
    '</table>' +
    '<p><b>File PDF Giấy giới thiệu thực tập</b> đã được tạo tự động và đính kèm trong email này.</p>' +
    '<p><a href="' + pdfUrl + '" style="display:inline-block;padding:8px 16px;background:#0066cc;color:#fff;text-decoration:none;border-radius:4px;font-weight:bold;">Mở xem trên Google Drive</a></p>' +
    '<hr style="border:none;border-top:1px solid #eee;margin:20px 0;">' +
    '<p style="font-size:12px;color:#777;">Thư được gửi tự động từ Hệ thống Quản lý Thực tập — Khoa Công nghệ Thông tin, Đại học Tôn Đức Thắng.</p>' +
    '</div>';

  const options = {
    attachments: [pdfBlob],
    htmlBody: html,
    name: CONFIG.SENDER_NAME
  };

  if (CONFIG.CC_EMAILS) {
    options.cc = CONFIG.CC_EMAILS;
  }

  GmailApp.sendEmail(CONFIG.LECTURER_EMAIL, subject, plain, options);
}

function row_(label, value) {
  return '<tr style="border-bottom:1px solid #eee;">' +
         '<td style="padding:6px 12px;font-weight:bold;width:38%;background:#f8f9fa;">' + escapeHtml_(label) + '</td>' +
         '<td style="padding:6px 12px;">' + value + '</td>' +
         '</tr>';
}

function escapeHtml_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
                  .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function notifyAdmin_(row, err) {
  try {
    const to = CONFIG.ADMIN_EMAIL || Session.getEffectiveUser().getEmail();
    MailApp.sendEmail(
      to,
      '[GGT-CNTT] Cảnh báo lỗi xử lý tại dòng ' + row,
      'Hệ thống không thể xử lý dòng ' + row + ' trong Sheet.\n\n' +
      'Nguyên nhân lỗi: ' + err.message + '\n\n' +
      'Vui lòng kiểm tra lại thông tin sinh viên tại dòng ' + row + ' và chạy lại hàm retryPending().'
    );
  } catch (ignore) {}
}

// ==============================================================================
// 8. QUẢN LÝ CỘT VÀ GHI TRẠNG THÁI VÀO GOOGLE SHEET
// ==============================================================================

/**
 * Tự động tìm kiếm vị trí các cột hoặc thêm các cột quản trị nếu chưa có.
 * Đảm bảo:
 * 1. Nhận diện các cột đã có sẵn trên Sheet (kể cả tên biến thể: "TRẠNG THÁI GỬI GGT", "LINK PDF GGT", "THỜI GIAN XỬ LÝ GGT")
 * 2. Không tạo cột mới trùng lặp khi đã có cột tương đương.
 * 3. Chỉ thêm cột mới khi cột quản trị đó hoàn toàn chưa có trong Sheet.
 */
function getColumns_(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headerValues = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const headers = headerValues.map(function (h) { return String(h).trim(); });

  // Danh mục cấu hình cột quản trị và các biến thể tên cột thực tế:
  // - STATUS: Nhận diện "TRẠNG THÁI GỬI GGT", "GỬI GGT", "TRẠNG THÁI GGT", "TRẠNG THÁI", "Status"...
  // - LINK: Nhận diện "LINK PDF GGT", "PDF GGT", "LINK GGT", "Link PDF", "PDF Link"...
  // - TIME: Nhận diện "THỜI GIAN XỬ LÝ GGT", "XỬ LÝ GGT", "THỜI GIAN GỬI GGT", "Thời gian xử lý", "Thời gian gửi"...
  // - TAX_NAME: Nhận diện "Tên Doanh Nghiệp (Tra cứu MST)", "Tên DN Tra cứu MST", "Tra cứu MST"...
  // - TAX_NOTE: Nhận diện "Kết quả đối soát MST", "Đối soát MST", "Kết quả đối soát", "Ghi chú đối soát"...
  const adminDefs = [
    {
      key: 'STATUS',
      defaultHeader: CONFIG.HEADERS.STATUS,
      exact: [
        CONFIG.HEADERS.STATUS,
        'TRẠNG THÁI GỬI GGT',
        'TRẠNG THÁI GGT',
        'TRẠNG THÁI GỬI',
        'TRẠNG THÁI',
        'GỬI GGT',
        'STATUS'
      ],
      partial: [
        'gửi ggt',
        'gui ggt',
        'trạng thái gửi ggt',
        'trang thai gui ggt',
        'trạng thái ggt',
        'trang thai ggt'
      ]
    },
    {
      key: 'LINK',
      defaultHeader: CONFIG.HEADERS.LINK,
      exact: [
        CONFIG.HEADERS.LINK,
        'LINK PDF GGT',
        'LINK GGT',
        'PDF GGT',
        'Link PDF',
        'PDF Link',
        'LINK FILE GGT',
        'FILE PDF GGT'
      ],
      partial: [
        'pdf ggt',
        'link pdf',
        'pdf link',
        'link ggt',
        'file ggt'
      ]
    },
    {
      key: 'TIME',
      defaultHeader: CONFIG.HEADERS.TIME,
      exact: [
        CONFIG.HEADERS.TIME,
        'THỜI GIAN XỬ LÝ GGT',
        'THỜI GIAN GỬI GGT',
        'THỜI GIAN XỬ LÝ',
        'XỬ LÝ GGT',
        'Thời gian xử lý',
        'Thời gian gửi'
      ],
      partial: [
        'xử lý ggt',
        'xu ly ggt',
        'thời gian xử lý ggt',
        'thoi gian xu ly ggt',
        'thời gian xử lý',
        'thoi gian xu ly',
        'thời gian gửi ggt',
        'thoi gian gui ggt'
      ]
    },
    {
      key: 'TAX_NAME',
      defaultHeader: CONFIG.HEADERS.TAX_NAME,
      exact: [
        CONFIG.HEADERS.TAX_NAME,
        'Tên Doanh Nghiệp (Tra cứu MST)',
        'Tên DN Tra cứu MST',
        'Tên DN (Tra cứu MST)',
        'Tên tra cứu MST',
        'Tên DN theo MST',
        'Tra cứu MST'
      ],
      partial: [
        'tra cứu mst',
        'tra cuu mst',
        'dn theo mst'
      ]
    },
    {
      key: 'TAX_NOTE',
      defaultHeader: CONFIG.HEADERS.TAX_NOTE,
      exact: [
        CONFIG.HEADERS.TAX_NOTE,
        'Kết quả đối soát MST',
        'Kết quả đối soát',
        'Đối soát MST',
        'Ghi chú đối soát',
        'Đối soát'
      ],
      partial: [
        'đối soát mst',
        'doi soat mst',
        'đối soát',
        'doi soat'
      ]
    }
  ];

  if (CONFIG.AUTO_NUMBERING) adminDefs.push({key:'DOC_NO', defaultHeader:CONFIG.HEADERS.DOC_NO, exact:[CONFIG.HEADERS.DOC_NO, 'SỐ GGT'], partial:['số giấy ggt', 'số ggt']});

  const cols = {};
  const claimedAdminCols = [];

  // Bước 1: Quét tìm các cột quản trị ĐÃ TỒN TẠI trong Sheet (không tạo mới nếu đã có)
  adminDefs.forEach(function (def) {
    const foundCol = findColByExactOrPartial_(headers, def.exact, def.partial, claimedAdminCols);
    if (foundCol !== -1) {
      cols[def.key] = foundCol;
      claimedAdminCols.push(foundCol);
    }
  });

  // Bước 2: Chỉ tạo thêm cột mới đối với những cột quản trị THỰC SỰ CHƯA CÓ
  adminDefs.forEach(function (def) {
    if (!cols[def.key]) {
      // Tìm vị trí cột trống đầu tiên chưa bị chiếm dụng
      let targetCol = -1;
      for (let i = 0; i < headers.length; i++) {
        if (!headers[i] || headers[i].trim() === '') {
          if (claimedAdminCols.indexOf(i + 1) === -1) {
            targetCol = i + 1;
            break;
          }
        }
      }

      if (targetCol === -1) {
        targetCol = headers.length + 1;
        headers.push(def.defaultHeader);
      } else {
        headers[targetCol - 1] = def.defaultHeader;
      }

      sheet.getRange(1, targetCol).setValue(def.defaultHeader);
      cols[def.key] = targetCol;
      claimedAdminCols.push(targetCol);
    }
  });

  const lowerHeaders = headers.map(function (h) {
    return removeAccents_(h).toLowerCase().trim();
  });

  // Bước 3: Tìm các cột nghiệp vụ của sinh viên (loại trừ các cột quản trị để tránh xung đột)
  const claimedStudentCols = [];
  Object.keys(FIELD_KEYWORDS).forEach(function (key) {
    let found = -1;
    const keywords = FIELD_KEYWORDS[key];

    // Ưu tiên 3.1: Khớp chính xác hoàn toàn tiêu đề cột
    for (let k = 0; k < keywords.length; k++) {
      const kw = removeAccents_(keywords[k]).toLowerCase().trim();
      for (let i = 0; i < lowerHeaders.length; i++) {
        if (claimedAdminCols.indexOf(i + 1) !== -1 || claimedStudentCols.indexOf(i + 1) !== -1) continue;
        if (lowerHeaders[i] === kw) {
          found = i + 1;
          break;
        }
      }
      if (found !== -1) break;
    }

    // Ưu tiên 3.2: Khớp một phần nếu chưa tìm thấy ở khớp chính xác
    if (found === -1) {
      for (let k = 0; k < keywords.length; k++) {
        const kw = removeAccents_(keywords[k]).toLowerCase().trim();
        for (let i = 0; i < lowerHeaders.length; i++) {
          if (claimedAdminCols.indexOf(i + 1) !== -1 || claimedStudentCols.indexOf(i + 1) !== -1) continue;
          if (lowerHeaders[i].indexOf(kw) !== -1) {
            found = i + 1;
            break;
          }
        }
        if (found !== -1) break;
      }
    }

    if (found === -1 && REQUIRED_FIELDS.indexOf(key) !== -1) {
      throw new Error('Không tìm thấy cột dữ liệu cho trường bắt buộc "' + key + '". Kiểm tra tiêu đề bảng tính.');
    }
    cols[key] = found;
    if (found !== -1) {
      claimedStudentCols.push(found);
    }
  });

  return cols;
}

/**
 * Tìm vị trí cột (1-based) khớp chính xác hoặc khớp một phần theo danh sách ứng viên.
 * - Ưu tiên 1: Khớp chính xác hoàn toàn (exact match, không dấu/hoa thường).
 * - Ưu tiên 2: Khớp một phần (partial substring match) với các từ khóa an toàn.
 * Hỗ trợ cả 2 signature:
 * - findColByExactOrPartial_(headers, exactList, partialList, excludeCols)
 * - findColByExactOrPartial_(headers, candidatesList, excludeCols)
 * @param {Array<string>} headers Danh sách tiêu đề trong Sheet
 * @param {Array<string>} exactCandidates Danh sách từ khóa khớp chính xác
 * @param {Array<string>|Array<number>} [partialCandidates] Danh sách từ khóa khớp một phần, hoặc excludeCols nếu dùng signature cũ
 * @param {Array<number>} [excludeCols] Danh sách các chỉ số cột (1-based) đã được nhận trước đó
 * @returns {number} Chỉ số cột (1-based) hoặc -1 nếu không tìm thấy
 */
function findColByExactOrPartial_(headers, exactCandidates, partialCandidates, excludeCols) {
  let exactList = [];
  let partialList = [];
  let excluded = [];

  if (arguments.length <= 3 && Array.isArray(partialCandidates) &&
      (partialCandidates.length === 0 || typeof partialCandidates[0] === 'number')) {
    // Signature cũ: (headers, candidates, excludeCols)
    exactList = exactCandidates || [];
    partialList = exactCandidates || [];
    excluded = partialCandidates || [];
  } else {
    exactList = exactCandidates || [];
    partialList = partialCandidates || [];
    excluded = excludeCols || [];
  }

  // Ưu tiên 1: Khớp chính xác (exact match)
  for (let c = 0; c < exactList.length; c++) {
    const cand = removeAccents_(String(exactList[c] || '')).toLowerCase().trim();
    if (!cand) continue;
    for (let i = 0; i < headers.length; i++) {
      if (excluded.indexOf(i + 1) !== -1) continue;
      const h = removeAccents_(String(headers[i] || '')).toLowerCase().trim();
      if (h === cand) {
        return i + 1;
      }
    }
  }

  // Ưu tiên 2: Khớp một phần (substring match) với từ khóa an toàn
  for (let c = 0; c < partialList.length; c++) {
    const cand = removeAccents_(String(partialList[c] || '')).toLowerCase().trim();
    if (!cand) continue;
    for (let i = 0; i < headers.length; i++) {
      if (excluded.indexOf(i + 1) !== -1) continue;
      const h = removeAccents_(String(headers[i] || '')).toLowerCase().trim();
      if (h.indexOf(cand) !== -1) {
        return i + 1;
      }
    }
  }

  return -1;
}

/**
 * Ghi kết quả thành công và thông tin đối soát vào đúng các cột quản trị trong Sheet.
 */
function writePendingApproval_(sheet, row, cols, taxResult) {
  sheet.getRange(row, cols.STATUS).setValue('Chờ duyệt');
  if (cols.TIME) sheet.getRange(row, cols.TIME).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  if (cols.TAX_NAME && taxResult) sheet.getRange(row, cols.TAX_NAME).setValue(taxResult.name || '');
  if (cols.TAX_NOTE) sheet.getRange(row, cols.TAX_NOTE).setValue(taxResult ? taxResult.matchNote : 'Đã qua kiểm tra dữ liệu; chờ duyệt');
}

function writeSuccess_(sheet, row, cols, pdfUrl, taxResult, docNumber) {
  const statusText = CONFIG.SEND_EMAIL ? 'Đã gửi' : 'Đã tạo PDF';
  if (cols.STATUS && cols.STATUS > 0) {
    sheet.getRange(row, cols.STATUS).setValue(statusText);
  }
  if (cols.LINK && cols.LINK > 0) {
    sheet.getRange(row, cols.LINK).setValue(pdfUrl);
  }
  if (cols.TIME && cols.TIME > 0) {
    sheet.getRange(row, cols.TIME).setValue(
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss')
    );
  }

  if (CONFIG.AUTO_NUMBERING && cols.DOC_NO && cols.DOC_NO > 0) sheet.getRange(row, cols.DOC_NO).setValue(docNumber || '');
  if (cols.TAX_NAME && cols.TAX_NAME > 0 && taxResult) {
    sheet.getRange(row, cols.TAX_NAME).setValue(taxResult.name || '');
  }
  if (cols.TAX_NOTE && cols.TAX_NOTE > 0) {
    sheet.getRange(row, cols.TAX_NOTE).setValue(taxResult ? (taxResult.matchNote || 'Hợp lệ') : 'Bỏ qua kiểm tra MST');
  }
}

/**
 * Ghi lỗi và chi tiết đối soát vào đúng các cột quản trị trong Sheet.
 */
function writeError_(sheet, row, cols, errorMsg, taxResult) {
  if (cols.STATUS && cols.STATUS > 0) {
    sheet.getRange(row, cols.STATUS).setValue('Lỗi: ' + errorMsg);
  }
  if (cols.TIME && cols.TIME > 0) {
    sheet.getRange(row, cols.TIME).setValue(
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss')
    );
  }
  if (cols.TAX_NOTE && cols.TAX_NOTE > 0) {
    sheet.getRange(row, cols.TAX_NOTE).setValue(errorMsg);
  }
  if (cols.TAX_NAME && cols.TAX_NAME > 0 && taxResult && taxResult.name) {
    sheet.getRange(row, cols.TAX_NAME).setValue(taxResult.name);
  }
}

/**
 * Tự động tìm sheet câu trả lời trong Workbook
 */
function getResponseSheet_() {
  const ss = SpreadsheetApp.getActive();
  if (CONFIG.SHEET_NAME) {
    const s = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (s) return s;
  }

  const sheets = ss.getSheets();
  // Ưu tiên sheet có tên chứa "Form Responses" hoặc "Câu trả lời"
  for (let i = 0; i < sheets.length; i++) {
    const name = sheets[i].getName().toLowerCase();
    if (name.indexOf('form responses') !== -1 || name.indexOf('câu trả lời') !== -1) {
      return sheets[i];
    }
  }

  // Tự quét tìm sheet đầu tiên có các cột bắt buộc
  for (let s = 0; s < sheets.length; s++) {
    const sh = sheets[s];
    if (sh.getLastColumn() === 0) continue;
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
                      .map(function (h) { return removeAccents_(String(h)).toLowerCase(); });

    const ok = REQUIRED_FIELDS.every(function (k) {
      return FIELD_KEYWORDS[k].some(function (kw) {
        const normKw = removeAccents_(kw).toLowerCase();
        return headers.some(function (h) { return h.indexOf(normKw) !== -1; });
      });
    });
    if (ok) return sh;
  }

  // Mặc định sheet đầu tiên nếu không khớp
  return sheets[0];
}

// ==============================================================================
// 9. CÁC HÀM CÀI ĐẶT, KIỂM THỬ VÀ VẬN HÀNH BẢO TRÌ
// ==============================================================================

/**
 * Thiết lập hệ thống (Chạy 1 lần trong Google Apps Script Editor).
 * - Cài đặt Trigger tự động khi có Form Submit.
 * - Kiểm tra mẫu Doc, thư mục Drive và các cột trong Sheet.
 */
function getPrivateConfig_(key) {
  const localValue = CONFIG[key];
  if (localValue && String(localValue).trim()) return String(localValue).trim();
  const stores = [PropertiesService.getScriptProperties()];
  try { stores.push(PropertiesService.getDocumentProperties()); } catch (ignore) {}
  try { stores.push(PropertiesService.getUserProperties()); } catch (ignore) {}
  for (let i = 0; i < stores.length; i++) {
    const value = stores[i].getProperty(key);
    if (value && String(value).trim()) return String(value).trim();
  }
  throw new Error('Chưa cấu hình ' + key + '. Điền ID vào CONFIG.' + key + ' hoặc lưu thuộc tính tương ứng trong Apps Script.');
}

function kiemTraCauHinhDrive() {
  const result = {};
  ['TEMPLATE_DOC_ID', 'OUTPUT_FOLDER_ID'].forEach(function (key) {
    try { result[key] = 'OK (' + getPrivateConfig_(key).length + ' ký tự)'; }
    catch (err) { result[key] = 'THIẾU'; }
  });
  Logger.log(JSON.stringify(result));
  return result;
}

function setup() {
  // Kiểm tra cấu hình và sheet trước khi thay đổi trigger hiện có.
  const sheet = getResponseSheet_();
  getColumns_(sheet);
  checkConfig_();

  // Chỉ thay trigger sau khi mọi kiểm tra cấu hình đã qua.
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'handleFormSubmit') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Tạo trigger "onFormSubmit" mới
  ScriptApp.newTrigger('handleFormSubmit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onFormSubmit()
    .create();

  Logger.log('>>> Cài đặt thành công! Hệ thống sẵn sàng tự động tiếp nhận Form sinh viên.');
}

/**
 * Kiểm tra tính hợp lệ của cấu hình Drive, Doc mẫu và placeholder
 */
function checkConfig_() {
  if (CONFIG.PDF_SHARING_MODE !== 'RESTRICTED' && CONFIG.PDF_SHARING_MODE !== 'ANYONE_WITH_LINK') {
    throw new Error('PDF_SHARING_MODE chỉ nhận RESTRICTED hoặc ANYONE_WITH_LINK.');
  }
  if (!Number.isInteger(CONFIG.MAX_BATCH_ROWS) || CONFIG.MAX_BATCH_ROWS < 1 || CONFIG.MAX_BATCH_ROWS > 50) {
    throw new Error('MAX_BATCH_ROWS phải là số nguyên từ 1 đến 50.');
  }
  if (!Number.isInteger(CONFIG.MAX_QUEUE_ROWS) || CONFIG.MAX_QUEUE_ROWS < 1 || CONFIG.MAX_QUEUE_ROWS > 500) throw new Error('MAX_QUEUE_ROWS phải từ 1 đến 500.');
  if (CONFIG.AI_ENABLED && (CONFIG.AI_CONFIDENCE_THRESHOLD < 0 || CONFIG.AI_CONFIDENCE_THRESHOLD > 1 || CONFIG.AI_DICE_MIN > CONFIG.AI_DICE_MAX)) throw new Error('Cấu hình ngưỡng AI không hợp lệ.');
  DriveApp.getFolderById(getPrivateConfig_('OUTPUT_FOLDER_ID'));
  const text = DocumentApp.openById(getPrivateConfig_('TEMPLATE_DOC_ID')).getBody().getText();
  if (CONFIG.AUTO_NUMBERING && text.indexOf('{{SO_GGT}}') === -1) throw new Error('Đã bật AUTO_NUMBERING nhưng Google Doc mẫu thiếu placeholder {{SO_GGT}}.');

  const requiredPlaceholders = ['{{HO_TEN}}', '{{MSSV}}', '{{EMAIL}}', '{{SDT}}', '{{CONG_TY}}', '{{VI_TRI}}'];
  const missing = requiredPlaceholders.filter(function (p) {
    return text.indexOf(p) === -1;
  });

  if (missing.length > 0) {
    throw new Error('Google Doc mẫu còn thiếu các placeholder bắt buộc: ' + missing.join(', '));
  }

  Logger.log('Cấu hình Google Doc và Thư mục Drive hoàn toàn chính xác.');
}

/**
 * Kiểm tra xem một trạng thái đã hoàn tất hoặc đã được giảng viên phê duyệt trước đó hay chưa.
 * Hỗ trợ các định dạng trong Sheet thực tế của Khoa:
 * - "Done", "[10/7/2025] Done", "2/4 - Done"
 * - "Đã gửi"
 * - "ĐỒNG Ý CẤP GIẤY GTTT (14/4)"
 * - "22/5/2024 - Đã xử lý"
 * - "Đã cấp ngày 31/10/2023"
 */
function isAlreadyProcessed_(status) {
  if (!status) return false;
  const s = removeAccents_(String(status)).toLowerCase().trim();
  if (!s) return false;

  // Nếu trạng thái bắt đầu bằng tiền tố lỗi/thất bại -> không bao giờ coi là đã hoàn thành
  if (s.indexOf('loi') === 0 || s.indexOf('error') === 0 || s.indexOf('fail') === 0) {
    return false;
  }

  if (s.indexOf('done') !== -1) return true;
  if (s.indexOf('da gui') !== -1) return true;
  if (s.indexOf('da tao pdf') !== -1) return true;
  if (s.indexOf('da xu ly') !== -1) return true;
  if (s.indexOf('dong y cap') !== -1) return true;
  if (s.indexOf('da cap') !== -1) return true;
  return false;
}

/**
 * Chạy thử nghiệm với dòng cuối cùng trong Sheet
 */
function testLastRow() {
  const sheet = getResponseSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('Sheet chưa có dữ liệu sinh viên nào.');

  processRow_(sheet, lastRow, true);
}

/**
 * Hiển thị hộp thoại UI (SpreadsheetApp.getUi().prompt() hoặc Browser.inputBox())
 * để nhận số dòng cần kiểm thử từ người dùng.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {number|null} Số dòng hợp lệ hoặc null nếu người dùng hủy
 */
function promptForRowNumber_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    throw new Error('Sheet chưa có dữ liệu sinh viên nào (chỉ có dòng tiêu đề hoặc rỗng).');
  }

  let input = '';
  // 1. Thử dùng SpreadsheetApp.getUi() (chuẩn container-bound Apps Script)
  try {
    const ui = SpreadsheetApp.getUi();
    const promptTitle = 'Kiểm thử dòng dữ liệu GGT';
    const promptMsg = 'Nhập số thứ tự dòng cần kiểm thử trong Sheet (từ 2 đến ' + lastRow + '):';
    const response = ui.prompt(promptTitle, promptMsg, ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() !== ui.Button.OK) {
      Logger.log('Người dùng đã bấm Hủy bỏ trong hộp thoại kiểm thử.');
      return null;
    }
    input = response.getResponseText().trim();
  } catch (e1) {
    // 2. Fallback sang Browser.inputBox nếu getUi() không khả dụng
    try {
      input = Browser.inputBox(
        'Kiểm thử dòng dữ liệu GGT',
        'Nhập số thứ tự dòng cần kiểm thử (từ 2 đến ' + lastRow + '):',
        Browser.Buttons.OK_CANCEL
      );
      if (input.toLowerCase() === 'cancel' || !input || input.trim() === '') {
        Logger.log('Người dùng đã hủy bỏ trong hộp thoại kiểm thử.');
        return null;
      }
      input = input.trim();
    } catch (e2) {
      // 3. Fallback khi không có giao diện người dùng (chạy standalone hoặc headless test)
      Logger.log('Không có UI context (do chạy trực tiếp từ Apps Script Editor). Mặc định chọn dòng 2 để kiểm thử. Để chọn dòng khác trên giao diện, hãy mở Google Sheet và chọn menu "GGT-CNTT > Kiểm thử dòng cụ thể...".');
      return 2;
    }
  }

  if (!/^\d+$/.test(input)) {
    throw new Error('Số dòng không hợp lệ: "' + input + '". Vui lòng nhập số nguyên từ 2 đến ' + lastRow + '.');
  }

  const row = parseInt(input, 10);
  if (row < 2 || row > lastRow) {
    throw new Error('Số dòng không hợp lệ: "' + input + '". Vui lòng nhập số nguyên từ 2 đến ' + lastRow + '.');
  }
  return row;
}

/**
 * Chạy thử nghiệm với hộp thoại prompt hỏi số dòng từ người dùng.
 * Thích hợp chạy trực tiếp từ dropdown Apps Script editor hoặc custom menu.
 */
function testRowPrompt() {
  const sheet = getResponseSheet_();
  const row = promptForRowNumber_(sheet);
  if (row) {
    Logger.log('Bắt đầu kiểm thử dòng ' + row + ' theo yêu cầu từ prompt...');
    testRow(row);
  }
}

/**
 * Hàm wrapper testRow2() (không tham số) sử dụng UI prompt để hỏi số dòng từ người dùng.
 * Khắc phục triệt để lỗi khi người dùng chọn chạy từ dropdown "Select function" trong Apps Script Editor.
 */
function testRow2() {
  testRowPrompt();
}

/**
 * Chạy thử nghiệm với một dòng số bất kỳ do người dùng chỉ định.
 * @param {number} [rowNumber] Ví dụ: testRow(2). Nếu không truyền (undefined do chạy từ dropdown), tự động bật UI prompt hỏi số dòng.
 */
function testRow(rowNumber) {
  const sheet = getResponseSheet_();
  if (rowNumber === undefined || rowNumber === null || rowNumber === '') {
    // Tự động kích hoạt UI prompt nếu chạy từ dropdown Apps Script mà không truyền tham số
    rowNumber = promptForRowNumber_(sheet);
    if (!rowNumber) return;
  }

  const num = Number(rowNumber);
  if (isNaN(num) || num < 2 || num > sheet.getLastRow()) {
    throw new Error('Số dòng không hợp lệ: ' + rowNumber);
  }
  processRow_(sheet, num, true);
}

/**
 * Phân tích chuỗi danh sách dòng linh hoạt do người dùng nhập.
 * Hỗ trợ các định dạng:
 * - Dòng đơn lẻ:      "2, 5, 15"
 * - Khoảng dòng:       "7-10"
 * - Kết hợp:           "2, 5, 7-10, 15, 20-25"
 * @param {string} input Chuỗi người dùng nhập
 * @param {number} maxRow Số dòng lớn nhất hợp lệ trong Sheet
 * @returns {number[]} Mảng số dòng đã sắp xếp, loại bỏ trùng lặp
 */
function parseRowList_(input, maxRow) {
  if (!input || !input.trim()) return [];

  var parts = input.split(/[,;]+/);
  var rows = {};

  for (var i = 0; i < parts.length; i++) {
    var part = parts[i].trim();
    if (!part) continue;

    // Kiểm tra dạng khoảng: "7-10", "7–10", "7—10"
    var rangeMatch = part.match(/^(\d+)\s*[–—-]\s*(\d+)$/);
    if (rangeMatch) {
      var rStart = parseInt(rangeMatch[1], 10);
      var rEnd = parseInt(rangeMatch[2], 10);
      if (rStart > rEnd) { var tmp = rStart; rStart = rEnd; rEnd = tmp; }
      // Giới hạn khoảng tối đa 500 dòng để tránh treo hệ thống
      if (rEnd - rStart > 500) {
        throw new Error('Khoảng dòng quá lớn (' + rStart + '-' + rEnd + '). Tối đa 500 dòng mỗi lần.');
      }
      for (var r = rStart; r <= rEnd; r++) {
        if (r >= 2 && r <= maxRow) rows[r] = true;
      }
    } else if (/^\d+$/.test(part)) {
      var num = parseInt(part, 10);
      if (num >= 2 && num <= maxRow) rows[num] = true;
    } else {
      throw new Error('Không hiểu đoạn nhập: "' + part + '". Hãy nhập dạng: 2, 5, 7-10, 15');
    }
  }

  var result = Object.keys(rows).map(function (k) { return parseInt(k, 10); });
  result.sort(function (a, b) { return a - b; });
  return result;
}

/**
 * Hiển thị hộp thoại UI để nhận danh sách nhiều dòng từ người dùng.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} title Tiêu đề hộp thoại
 * @param {string} actionLabel Mô tả hành động (ví dụ: "kiểm thử", "in PDF hàng loạt")
 * @returns {number[]|null} Mảng số dòng hoặc null nếu người dùng hủy
 */
function promptForMultipleRows_(sheet, title, actionLabel) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    throw new Error('Sheet chưa có dữ liệu sinh viên nào (chỉ có dòng tiêu đề hoặc rỗng).');
  }

  var msg = 'Nhập danh sách các dòng cần ' + actionLabel + ' (từ 2 đến ' + lastRow + ').\n\n' +
            'Hỗ trợ nhiều định dạng:\n' +
            '• Dòng đơn lẻ: 2, 5, 15\n' +
            '• Khoảng dòng: 7-10\n' +
            '• Kết hợp: 2, 5, 7-10, 15, 20-25';

  var input = '';
  try {
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt(title, msg, ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() !== ui.Button.OK) {
      Logger.log('Người dùng đã bấm Hủy bỏ.');
      return null;
    }
    input = response.getResponseText().trim();
  } catch (e1) {
    try {
      input = Browser.inputBox(title, msg, Browser.Buttons.OK_CANCEL);
      if (!input || input.toLowerCase() === 'cancel' || input.trim() === '') {
        Logger.log('Người dùng đã hủy bỏ.');
        return null;
      }
      input = input.trim();
    } catch (e2) {
      Logger.log('Hãy chạy thao tác từ menu trong Google Sheet để chọn dòng.');
      return null;
    }
  }

  var rowList = parseRowList_(input, lastRow);
  if (rowList.length === 0) {
    throw new Error('Không có dòng hợp lệ nào trong danh sách đã nhập: "' + input + '"');
  }
  if (rowList.length > CONFIG.MAX_QUEUE_ROWS) {
    throw new Error('Danh sách vượt quá ' + CONFIG.MAX_QUEUE_ROWS + ' dòng cho một yêu cầu.');
  }
  return rowList;
}

/**
 * Kiểm thử nhiều dòng cùng lúc (có hộp thoại UI).
 * Hỗ trợ nhập dạng: "2, 5, 7-10, 15" để kiểm thử tất cả các dòng đó.
 * Mỗi dòng bỏ qua lọc ngày (chế độ test), nhưng vẫn giữ chốt chống xử lý trùng.
 */
/** Lưu danh sách đã chọn và xử lý nhiều đợt nền, tiếp tục tự động sau mỗi phút. */
function enqueueBatch_(sheet, rowList, isTestMode, label) {
  const props = PropertiesService.getScriptProperties();
  const key = 'GGT_BATCH_QUEUE';
  if (props.getProperty(key)) throw new Error('Đang có một lô xử lý nền. Hãy đợi lô hiện tại hoàn tất.');
  if (!rowList || !rowList.length) return;
  if (rowList.length > CONFIG.MAX_QUEUE_ROWS) throw new Error('Tối đa ' + CONFIG.MAX_QUEUE_ROWS + ' dòng cho mỗi lô.');
  const queue = {spreadsheetId: sheet.getParent().getId(), sheetName: sheet.getName(), rows: rowList,
    index: 0, success: 0, failed: 0, testMode: Boolean(isTestMode), label: label || 'xử lý'};
  props.setProperty(key, JSON.stringify(queue));
  try {
    ScriptApp.newTrigger('continueBatchQueue_').timeBased().everyMinutes(1).create();
  } catch (err) { props.deleteProperty(key); throw err; }
  continueBatchQueue_(); // chạy ngay đợt đầu; phần còn lại tự chạy nền
}

function continueBatchQueue_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  const props = PropertiesService.getScriptProperties();
  const key = 'GGT_BATCH_QUEUE';
  const started = Date.now();
  let queue;
  try {
    const raw = props.getProperty(key);
    if (!raw) { removeQueueTriggers_(); return; }
    queue = JSON.parse(raw);
    const ss = SpreadsheetApp.openById(queue.spreadsheetId);
    const sheet = ss.getSheetByName(queue.sheetName);
    if (!sheet) throw new Error('Không tìm thấy sheet của lô đang chờ: ' + queue.sheetName);
    let processed = 0;
    while (queue.index < queue.rows.length && processed < CONFIG.MAX_BATCH_ROWS && Date.now() - started < CONFIG.BATCH_RUNTIME_BUDGET_MS) {
      const row = Number(queue.rows[queue.index]);
      try {
        if (processRow_(sheet, row, queue.testMode)) queue.success++;
        else queue.failed++;
      } catch (err) {
        queue.failed++;
        Logger.log('Lỗi lô ' + queue.label + ', dòng ' + row + ': ' + err.message);
      }
      queue.index++;
      processed++;
      props.setProperty(key, JSON.stringify(queue)); // checkpoint để lần chạy sau tiếp tục đúng dòng
      if (queue.index < queue.rows.length && processed < CONFIG.MAX_BATCH_ROWS) Utilities.sleep(CONFIG.BATCH_SLEEP_MS);
    }
    if (queue.index >= queue.rows.length) {
      props.deleteProperty(key);
      removeQueueTriggers_();
      const message = 'Hoàn tất ' + queue.label + ': ' + queue.success + ' thành công, ' + queue.failed + ' lỗi/bỏ qua (' + queue.rows.length + ' dòng).';
      Logger.log(message);
      try { SpreadsheetApp.openById(queue.spreadsheetId).toast(message, 'GGT-CNTT', 10); } catch (ignore) {}
    } else {
      Logger.log('Đã xử lý ' + queue.index + '/' + queue.rows.length + ' dòng; phần còn lại sẽ tự chạy tiếp.');
    }
  } catch (err) {
    if (queue) props.setProperty(key, JSON.stringify(queue));
    Logger.log('Lỗi điều phối lô nền: ' + err.message);
    throw err;
  } finally { lock.releaseLock(); }
}

function cancelBatchQueue() {
  const props = PropertiesService.getScriptProperties();
  const exists = props.getProperty('GGT_BATCH_QUEUE');
  props.deleteProperty('GGT_BATCH_QUEUE');
  removeQueueTriggers_();
  SpreadsheetApp.getActive().toast(exists ? 'Đã hủy phần còn lại của lô.' : 'Không có lô nào đang chờ.', 'GGT-CNTT', 5);
}

function removeQueueTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'continueBatchQueue_') ScriptApp.deleteTrigger(trigger);
  });
}

function testMultipleRows() {
  const sheet = getResponseSheet_();
  const rows = promptForMultipleRows_(sheet, 'Kiểm thử nhiều dòng GGT', 'kiểm thử');
  if (rows) enqueueBatch_(sheet, rows, true, 'kiểm thử');
}

/**
 * In PDF hàng loạt nhiều dòng cùng lúc (có hộp thoại UI).
 * Hỗ trợ nhập dạng: "2, 5, 7-10, 15" để xử lý tất cả các dòng đó.
 * Khác với kiểm thử: hàm này tuân thủ lọc ngày và KHÔNG reset trạng thái đã hoàn tất.
 */
function processBatchPrompt() {
  const sheet = getResponseSheet_();
  const rows = promptForMultipleRows_(sheet, 'In PDF hàng loạt GGT', 'in PDF hàng loạt');
  if (rows) enqueueBatch_(sheet, rows, false, 'in PDF hàng loạt');
}

/**
 * Trigger tự động tạo Menu thao tác nhanh trên Google Sheet khi mở file.
 * Giúp người dùng thực hiện kiểm thử và quản trị trực tiếp từ giao diện bảng tính.
 */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('GGT-CNTT')
      .addItem('Kiểm thử dòng cụ thể...', 'testRowPrompt')
      .addItem('Kiểm thử nhiều dòng...', 'testMultipleRows')
      .addItem('Kiểm thử dòng cuối cùng', 'testLastRow')
      .addSeparator()
      .addItem('In PDF hàng loạt...', 'processBatchPrompt')
      .addItem('Xử lý lại các dòng lỗi / chờ', 'retryPending')
      .addItem('Cập nhật Dashboard', 'refreshDashboard')
      .addSeparator()
      .addItem('Cài đặt hệ thống (Trigger)', 'setup')
      .addToUi();
  } catch (ignore) {}
}

function columnLetter_(column) {
  let n = column, out = '';
  while (n > 0) { const r = (n - 1) % 26; out = String.fromCharCode(65 + r) + out; n = Math.floor((n - 1) / 26); }
  return out;
}

function refreshDashboard() {
  const source = getResponseSheet_();
  const cols = getColumns_(source);
  const ss = source.getParent();
  let dash = ss.getSheetByName('Dashboard GGT');
  if (!dash) dash = ss.insertSheet('Dashboard GGT');
  dash.clear();
  const col = columnLetter_(cols.STATUS);
  const sheetRef = "'" + source.getName().replace(/'/g, "''") + "'!";
  const range = sheetRef + col + '2:' + col;
  const totalCol = columnLetter_(cols.HO_TEN);
  const rows = [
    ['BẢNG TỔNG HỢP GGT', 'Giá trị'],
    ['Tổng số hồ sơ', '=COUNTA(' + sheetRef + totalCol + '2:' + totalCol + ')'],
    ['Đã tạo PDF', '=COUNTIF(' + range + ',"Đã tạo PDF")+COUNTIF(' + range + ',"Đã gửi")'],
    ['Đang xử lý', '=COUNTIF(' + range + ',"Processing")'],
    ['Chờ duyệt', '=COUNTIF(' + range + ',"Chờ duyệt")'],
    ['Lỗi', '=COUNTIF(' + range + ',"Lỗi:*")'],
    ['Cập nhật lúc', new Date()]
  ];
  dash.getRange(1,1,rows.length,2).setValues(rows);
  dash.getRange(2,2,5,1).setNumberFormat('0');
  dash.getRange(7,2).setNumberFormat('dd/MM/yyyy HH:mm:ss');
  dash.getRange(1,1,1,2).setFontWeight('bold').setBackground('#d9ead3');
  dash.setFrozenRows(1);
  dash.autoResizeColumns(1,2);
  ss.setActiveSheet(dash);
  ss.toast('Đã cập nhật Dashboard GGT.', 'GGT-CNTT', 5);
}

function approveSelectedRow() {
  const sheet = getResponseSheet_();
  const row = sheet.getActiveRange().getRow();
  if (row < 2) throw new Error('Hãy chọn một dòng hồ sơ, không chọn dòng tiêu đề.');
  const cols = getColumns_(sheet);
  const status = String(sheet.getRange(row, cols.STATUS).getDisplayValue() || '').trim();
  if (status !== 'Chờ duyệt') throw new Error('Dòng đang chọn không ở trạng thái Chờ duyệt.');
  processRow_(sheet, row, false, true);
  sheet.getParent().toast('Đã xử lý thao tác duyệt dòng ' + row + '.', 'GGT-CNTT', 5);
}

/**
 * Hàm kiểm tra nhanh đối soát MST và Tên công ty trực tiếp trên Editor console.
 * Bao gồm toàn bộ các case thực tế từ dữ liệu mẫu của Khoa:
 * - Doanh nghiệp chuẩn, chi nhánh công ty tổng, viết tắt, và các ca sai lệch.
 */
function testTaxValidationManual() {
  const testCases = [
    { mst: '0301425912', name: 'Công ty TNHH Giải Pháp Phần Mềm Tường Minh', desc: 'Khớp chính xác TMA Solutions' },
    { mst: '0313577878', name: 'CTY TNHH MTV INNOVATIVE SOFTWARE DEVELOPMENT', desc: 'Viết tắt CTY TNHH MTV khớp ISD' },
    { mst: '0300942001-008', name: 'Tổng công ty điện lực miền Nam', desc: 'Chi nhánh khớp tên công ty tổng' },
    { mst: '0300954529', name: 'Tập đoàn Bưu chính Viễn thông Việt Nam (Trung Tâm Viễn Thông Nam Sài Gòn)', desc: 'VNPT chi nhánh Nam Sài Gòn (mẫu ngoặc đơn)' },
    { mst: '0305140259', name: 'Tổ chức tài chính vi mô CEP', desc: 'Tổ chức tài chính vi mô CEP (Inclusion ratio)' },
    { mst: '0301596604', name: 'CÔNG TY CỔ PHẦN CÔNG NGHỆ TIÊN PHONG', abbrev: 'ITD Corporation', desc: 'Khớp viết tắt ITD' },
    { mst: '0301425912', name: 'Công ty Cổ phần Công nghệ FPT', desc: 'Sẽ báo lỗi không khớp (FPT vs Tường Minh)' },
    { mst: '9999999999', name: 'Công ty Không Tồn Tại', desc: 'Sẽ báo lỗi MST không tồn tại' }
  ];

  testCases.forEach(function (tc) {
    Logger.log('----------------------------------------------------');
    Logger.log('Kiểm thử: [' + tc.mst + '] - ' + tc.name + ' (' + tc.desc + ')');
    try {
      const res = verifyTaxAndCompany_({
        mst: tc.mst,
        congTy: tc.name,
        tenTiengAnh: '',
        tenVietTat: tc.abbrev || ''
      });
      Logger.log('=> HỢP LỆ: Tên Thuế="' + res.name + '" | ' + res.matchNote);
    } catch (err) {
      Logger.log('=> BÁO LỖI: ' + err.message);
    }
  });
}

/**
 * Xử lý lại toàn bộ các dòng chưa hoàn thành hoặc đang báo lỗi.
 * Có cơ chế giãn cách thời gian (sleep) để không vi phạm rate-limit của API Thuế.
 */
function retryPending() {
  const sheet = getResponseSheet_();
  const cols = getColumns_(sheet);
  const last = sheet.getLastRow();
  if (last < 2) { SpreadsheetApp.getUi().alert('Sheet chưa có hồ sơ cần xử lý.'); return; }
  const statuses = sheet.getRange(2, cols.STATUS, last - 1, 1).getDisplayValues();
  const rows = [];
  for (let i = 0; i < statuses.length && rows.length < CONFIG.MAX_QUEUE_ROWS; i++) {
    const status = String(statuses[i][0] || '').trim();
    if (!isAlreadyProcessed_(status) && (status === '' || /^lỗi/i.test(status) || /^processing/i.test(status))) rows.push(i + 2);
  }
  if (!rows.length) { SpreadsheetApp.getUi().alert('Không có dòng lỗi hoặc đang chờ cần xử lý.'); return; }
  enqueueBatch_(sheet, rows, false, 'xử lý lại');
}

/**
 * Xử lý hàng loạt theo khoảng dòng (Ví dụ từ dòng 2 đến dòng 20)
 */
function processBatch(startRow, endRow) {
  const sheet = getResponseSheet_();
  const last = sheet.getLastRow();
  const start = Math.max(Number(startRow) || 2, 2);
  const end = Math.min(Number(endRow) || last, last);
  if (end < start) throw new Error('Khoảng dòng không hợp lệ.');
  const rows = [];
  for (let r = start; r <= end; r++) rows.push(r);
  enqueueBatch_(sheet, rows, false, 'xử lý theo khoảng dòng');
}
