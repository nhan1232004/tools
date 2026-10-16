/**
 * Automated Test Suite for GGT-CNTT Auto (Khoa CNTT - TDTU)
 * Run with: node test_suite.js
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const codePath = path.join(__dirname, 'Code.gs');
const code = fs.readFileSync(codePath, 'utf8');

const logs = [];
let currentPromptResponse = { button: 'OK', text: '2' };
let currentBrowserInput = '2';
let activeSheet = null;

const sandbox = {
  console,
  Logger: {
    log: (...args) => { logs.push(args.join(' ')); }
  },
  Session: {
    getScriptTimeZone: () => 'GMT+7',
    getEffectiveUser: () => ({ getEmail: () => 'admin@tdtu.edu.vn' })
  },
  Utilities: {
    formatDate: (d, tz, fmt) => '14/10/2026',
    sleep: (ms) => {}
  },
  SpreadsheetApp: {
    getActive: () => ({
      getSheets: () => [activeSheet],
      getSheetByName: (name) => activeSheet
    }),
    getUi: () => ({
      ButtonSet: { OK_CANCEL: 'OK_CANCEL' },
      Button: { OK: 'OK', CANCEL: 'CANCEL' },
      prompt: (title, msg, buttons) => ({
        getSelectedButton: () => currentPromptResponse.button,
        getResponseText: () => currentPromptResponse.text
      }),
      createMenu: (title) => {
        const menuObj = {
          addItem: (caption, func) => menuObj,
          addSeparator: () => menuObj,
          addToUi: () => {}
        };
        return menuObj;
      }
    })
  },
  Browser: {
    Buttons: { OK_CANCEL: 'OK_CANCEL' },
    inputBox: (title, msg, buttons) => currentBrowserInput
  }
};
vm.createContext(sandbox);

console.log('Compiling Code.gs in VM sandbox...');
vm.runInContext(code, sandbox);
console.log('OK: Code.gs compiled successfully!\n');

let passCount = 0;
let failCount = 0;

function assert(condition, testName, details) {
  if (condition) {
    console.log(`[PASS] ${testName}`);
    passCount++;
  } else {
    console.error(`[FAIL] ${testName}: ${details || ''}`);
    failCount++;
  }
}

// 1. Phone number normalization
assert(sandbox.normalizePhone_('0973313615') === '0973313615', 'Phone: standard 10 digits');
assert(sandbox.normalizePhone_('+84973313615') === '0973313615', 'Phone: +84 prefix');
assert(sandbox.normalizePhone_('84973313615') === '0973313615', 'Phone: 84 prefix (11 digits)');
assert(sandbox.normalizePhone_('973313615') === '0973313615', 'Phone: 9 digits missing leading 0');
assert(sandbox.normalizePhone_('7.75380223E8') === '0775380223', 'Phone: scientific notation format E+');
assert(sandbox.normalizePhone_(775380223) === '0775380223', 'Phone: numeric format');

// 2. MSSV normalization
assert(sandbox.normalizeMssv_('519H0107') === '519H0107', 'MSSV: standard alphanumeric');
assert(sandbox.normalizeMssv_('5.1900811E7') === '51900811', 'MSSV: scientific notation format E+');
assert(sandbox.normalizeMssv_(51900811) === '51900811', 'MSSV: numeric format');

// 3. MST normalization
assert(sandbox.normalizeMst_('0301425912') === '0301425912', 'MST: standard 10 digits');
assert(sandbox.normalizeMst_('1.100109154E9') === '1100109154', 'MST: scientific notation format E+');
assert(sandbox.normalizeMst_('0300942001-008') === '0300942001-008', 'MST: branch format with hyphen');
assert(sandbox.normalizeMst_('3.603119522E9') === '3603119522', 'MST: Bosch Vietnam scientific notation');
assert(sandbox.normalizeMst_(' 0300951119-003 ') === '0300951119-003', 'MST: trimmed whitespaces');

// 4. Branch Parent Company Extraction & Comparison
const vnptBranch = sandbox.compareCompanyNames_(
  'Tập đoàn Bưu chính Viễn thông Việt Nam (Trung Tâm Viễn Thông Nam Sài Gòn)',
  'CHI NHÁNH TẬP ĐOÀN BƯU CHÍNH VIỄN THÔNG VIỆT NAM (LOẠI HÌNH DOANH NGHIỆP: CÔNG TY TNHH) - VIỄN THÔNG THÀNH PHỐ HỒ CHÍ MINH',
  'Chi nhánh',
  'Vietnam Posts and Telecommunications Group',
  'VNPT'
);
assert(vnptBranch.isMatch === true, 'Branch matching: VNPT Nam Sài Gòn with parentheses', vnptBranch.reason);

const evnBranch = sandbox.compareCompanyNames_(
  'Tổng công ty Điện lực Miền Nam',
  'CHI NHÁNH TỔNG CÔNG TY ĐIỆN LỰC MIỀN NAM TNHH - CÔNG TY CÔNG NGHỆ THÔNG TIN ĐIỆN LỰC MIỀN NAM',
  'Chi nhánh',
  '',
  ''
);
assert(evnBranch.isMatch === true, 'Branch matching: EVN Southern Power simplified name', evnBranch.reason);

// 5. Corporate Abbreviation
const itdAbbrev = sandbox.compareCompanyNames_(
  'CÔNG TY CỔ PHẦN CÔNG NGHỆ TIÊN PHONG',
  'CÔNG TY CỔ PHẦN CÔNG NGHỆ ITD',
  'Doanh nghiệp / Đơn vị sự nghiệp công lập',
  'Innovative Technology Development Corporation',
  'ITD Corporation'
);
assert(itdAbbrev.isMatch === true, 'Abbreviation: ITD Corporation matches ITD', itdAbbrev.reason);

const isdMatch = sandbox.compareCompanyNames_(
  'CTY TNHH MTV INNOVATIVE SOFTWARE DEVELOPMENT',
  'CÔNG TY TNHH MỘT THÀNH VIÊN INNOVATIVE SOFTWARE DEVELOPMENT',
  'Doanh nghiệp',
  'INNOVATIVE SOFTWARE DEVELOPMENT',
  'ISD'
);
assert(isdMatch.isMatch === true, 'Abbreviation & Legal title: CTY TNHH MTV vs CÔNG TY TNHH MỘT THÀNH VIÊN', isdMatch.reason);

// 6. Token Inclusion Ratio for long legal entities
const cepMatch = sandbox.compareCompanyNames_(
  'Tổ chức tài chính vi mô CEP',
  'TỔ CHỨC TÀI CHÍNH VI MÔ TRÁCH NHIỆM HỮU HẠN MỘT THÀNH VIÊN CHO NGƯỜI LAO ĐỘNG NGHÈO TỰ TẠO VIỆC LÀM',
  'Doanh nghiệp / Đơn vị sự nghiệp công lập',
  'CAPITAL AID FOR EMPLOYMENT OF THE POOR MICROFINANCE INSTITUTION',
  'CEP'
);
assert(cepMatch.isMatch === true, 'Inclusion ratio: CEP Microfinance matches official legal title', cepMatch.reason);

// 7. Rejection of true negative mismatches
const fptVsTma = sandbox.compareCompanyNames_('Công ty Cổ phần Công nghệ FPT', 'CÔNG TY TNHH GIẢI PHÁP PHẦN MỀM TƯỜNG MINH', 'Doanh nghiệp', '', '');
assert(fptVsTma.isMatch === false, 'Negative: FPT vs Tường Minh rejected', fptVsTma.reason);

const fptVsBcons = sandbox.compareCompanyNames_('Công ty Cổ phần Công nghệ FPT', 'CÔNG TY CỔ PHẦN CÔNG NGHỆ BCONS', 'Doanh nghiệp', '', '');
assert(fptVsBcons.isMatch === false, 'Negative: FPT vs BCONS rejected', fptVsBcons.reason);

const shopeeVsLazada = sandbox.compareCompanyNames_('Công ty TNHH Shopee', 'CÔNG TY TNHH RECESS (LAZADA)', 'Doanh nghiệp', '', '');
assert(shopeeVsLazada.isMatch === false, 'Negative: Shopee vs Lazada rejected', shopeeVsLazada.reason);

// 8. Recognition of teacher's existing approvals
assert(sandbox.isAlreadyProcessed_('Done') === true, 'Status: Done recognized');
assert(sandbox.isAlreadyProcessed_('[10/7/2025] Done') === true, 'Status: [date] Done recognized');
assert(sandbox.isAlreadyProcessed_('ĐỒNG Ý CẤP GIẤY GTTT (14/4)') === true, 'Status: ĐỒNG Ý CẤP GIẤY GTTT recognized');
assert(sandbox.isAlreadyProcessed_('22/5/2024 - Đã xử lý') === true, 'Status: Đã xử lý recognized');
assert(sandbox.isAlreadyProcessed_('Đã cấp ngày 31/10/2023') === true, 'Status: Đã cấp recognized');
assert(sandbox.isAlreadyProcessed_('Đã gửi') === true, 'Status: Đã gửi recognized');
assert(sandbox.isAlreadyProcessed_('Đã tạo PDF') === true, 'Status: Đã tạo PDF recognized');
assert(sandbox.isAlreadyProcessed_('') === false, 'Status: Empty recognized as unprocessed');
assert(sandbox.isAlreadyProcessed_('Lỗi: mã số thuế không hợp lệ') === false, 'Status: Error recognized as unprocessed');

// 9. Date parsing and Filtering from August 1st, 2026
const filterAugust = new Date(2026, 7, 1); // 01/08/2026
const dJuly = sandbox.parseSubmissionDate_('31/07/2026 15:02:38');
assert(dJuly && dJuly < filterAugust, 'Date filter: July 31st 2026 correctly filtered out (< 01/08/2026)');

const dAug1 = sandbox.parseSubmissionDate_('01/08/2026 00:00:00');
assert(dAug1 && dAug1 >= filterAugust, 'Date filter: August 1st 2026 correctly accepted (>= 01/08/2026)');

const dAug6 = sandbox.parseSubmissionDate_('06/08/2026 08:31:20');
assert(dAug6 && dAug6 >= filterAugust, 'Date filter: August 6th 2026 correctly accepted');

const dExcelSerial = sandbox.parseSubmissionDate_(46240.355); // 06/08/2026
assert(dExcelSerial && dExcelSerial >= filterAugust, 'Date filter: Excel serial date parsed and accepted');

// 10. Input validation
const validStudent = sandbox.cleanAndValidate_({
  HO_TEN: 'Trương Khánh Duy',
  MSSV: '5.1900324E7',
  EMAIL: 'duy@gmail.com',
  SDT: '7.75380223E8',
  CONG_TY: 'VIỄN THÔNG LONG AN',
  MST: '1.100109154E9',
  VI_TRI: 'Data Analyst'
});
assert(validStudent.mssv === '51900324' && validStudent.sdt === '0775380223' && validStudent.mst === '1100109154', 'Validation: student with multiple scientific fields parsed correctly');

// Mock Sheet Factory
function createMockSheet(headerRow, dataRows) {
  dataRows = dataRows || [];
  const grid = [ [...headerRow] ];
  dataRows.forEach(r => grid.push([...r]));

  return {
    getName: () => 'Form Responses 1',
    getLastColumn: () => {
      let maxCols = 0;
      grid.forEach(row => { if (row.length > maxCols) maxCols = row.length; });
      return maxCols;
    },
    getLastRow: () => grid.length,
    getRange: (r, c, numRows, numCols) => {
      numRows = numRows || 1;
      numCols = numCols || 1;
      return {
        getValues: () => {
          const res = [];
          for (let i = 0; i < numRows; i++) {
            const rowArr = [];
            const rowIdx = (r - 1) + i;
            for (let j = 0; j < numCols; j++) {
              const colIdx = (c - 1) + j;
              rowArr.push(grid[rowIdx] && grid[rowIdx][colIdx] !== undefined ? grid[rowIdx][colIdx] : '');
            }
            res.push(rowArr);
          }
          return res;
        },
        setValue: (val) => {
          const rowIdx = r - 1;
          const colIdx = c - 1;
          while (grid.length <= rowIdx) grid.push([]);
          grid[rowIdx][colIdx] = val;
        },
        clearContent: () => {
          const rowIdx = r - 1;
          const colIdx = c - 1;
          if (grid[rowIdx]) grid[rowIdx][colIdx] = '';
        },
        getValue: () => {
          const rowIdx = r - 1;
          const colIdx = c - 1;
          return grid[rowIdx] && grid[rowIdx][colIdx] !== undefined ? grid[rowIdx][colIdx] : '';
        }
      };
    },
    _getGrid: () => grid
  };
}

// 11. Column Recognition: Teacher's Sheet with existing Q, R, S headers
const teacherHeaders = [
  'Timestamp', 'Họ và Tên', 'MSSV', 'SĐT', 'Email Cá Nhân',
  'Tên Học Phần', 'Tên Doanh Nghiệp', 'Tên Doanh Nghiệp Viết Bằng Tiếng Nước Ngoài',
  'Tên Doanh Nghiệp Viết Tắt', 'Mã Số Doanh Nghiệp', 'Thời Gian Bắt Đầu Thực Tập',
  'Họ và Tên Người Hướng Dẫn', 'SĐT Người Hướng Dẫn', 'Email Người Hướng Dẫn',
  'Ví Trí Thực Tập', 'GHI CHÚ',
  'TRẠNG THÁI GỬI GGT', // Q = 17
  'LINK PDF GGT',        // R = 18
  'THỜI GIAN XỬ LÝ GGT'  // S = 19
];

const sheetTeacher = createMockSheet(teacherHeaders);
const colsTeacher = sandbox.getColumns_(sheetTeacher);

assert(colsTeacher.STATUS === 17, 'Teacher Sheet: STATUS mapped to col 17 (TRẠNG THÁI GỬI GGT)');
assert(colsTeacher.LINK === 18, 'Teacher Sheet: LINK mapped to col 18 (LINK PDF GGT)');
assert(colsTeacher.TIME === 19, 'Teacher Sheet: TIME mapped to col 19 (THỜI GIAN XỬ LÝ GGT)');
assert(colsTeacher.TAX_NAME === 20, 'Teacher Sheet: TAX_NAME created at col 20');
assert(colsTeacher.TAX_NOTE === 21, 'Teacher Sheet: TAX_NOTE created at col 21');
assert(sheetTeacher.getLastColumn() === 21, 'Teacher Sheet: Total cols is 21 (NO duplicate STATUS/LINK/TIME created)');

// Idempotency: second call does not re-add or duplicate columns
const colsTeacher2 = sandbox.getColumns_(sheetTeacher);
assert(colsTeacher2.STATUS === 17 && colsTeacher2.LINK === 18 && colsTeacher2.TIME === 19 && sheetTeacher.getLastColumn() === 21, 'Teacher Sheet: Idempotent getColumns_ (no columns added on second call)');

// 12. Column Recognition: Variations & Partial Keywords
const partialHeaders = [
  'Dấu thời gian', 'Họ và tên', 'Mã số sinh viên', 'Điện thoại', 'Email',
  'Tên công ty', 'Mã số thuế',
  'GỬI GGT',
  'PDF GGT',
  'XỬ LÝ GGT'
];
const sheetPartial = createMockSheet(partialHeaders);
const colsPartial = sandbox.getColumns_(sheetPartial);
assert(colsPartial.STATUS === 8, 'Partial keywords: GỬI GGT mapped to STATUS');
assert(colsPartial.LINK === 9, 'Partial keywords: PDF GGT mapped to LINK');
assert(colsPartial.TIME === 10, 'Partial keywords: XỬ LÝ GGT mapped to TIME');
assert(colsPartial.TIMESTAMP === 1, 'Partial keywords: Dấu thời gian mapped to TIMESTAMP without conflict');

const oldHeaders = [
  'Timestamp', 'Họ và Tên', 'MSSV', 'SĐT', 'Email cá nhân',
  'Tên Doanh Nghiệp', 'Mã số doanh nghiệp',
  'TRẠNG THÁI',
  'Link PDF',
  'Thời gian xử lý',
  'Tên Doanh Nghiệp (Tra cứu MST)',
  'Kết quả đối soát MST'
];
const sheetOld = createMockSheet(oldHeaders);
const colsOld = sandbox.getColumns_(sheetOld);
assert(colsOld.STATUS === 8, 'Old Sheet: TRẠNG THÁI mapped to STATUS');
assert(colsOld.LINK === 9, 'Old Sheet: Link PDF mapped to LINK');
assert(colsOld.TIME === 10, 'Old Sheet: Thời gian xử lý mapped to TIME');
assert(colsOld.TAX_NAME === 11, 'Old Sheet: TAX_NAME mapped to col 11');
assert(colsOld.TAX_NOTE === 12, 'Old Sheet: TAX_NOTE mapped to col 12');
assert(sheetOld.getLastColumn() === 12, 'Old Sheet: Total cols is 12 (0 new columns added)');

// 13. Data Writing to Correct Columns: writeSuccess_ and writeError_
const testDataRow = [
  '2026-08-10', 'Trương Khánh Duy', '51900324', '0775380223', 'duy@gmail.com',
  'KTCN', 'VIỄN THÔNG LONG AN', '', '', '1100109154', '2026-08-15',
  '', '', '', 'Data Analyst', '', '', '', '', '', ''
];
const sheetData = createMockSheet(teacherHeaders, [testDataRow]);
const testCols = sandbox.getColumns_(sheetData);

sandbox.writeSuccess_(sheetData, 2, testCols, 'https://drive.google.com/file/d/test_ggt_123/view', {
  name: 'VIỄN THÔNG LONG AN - CHI NHÁNH TẬP ĐOÀN BƯU CHÍNH VIỄN THÔNG VIỆT NAM',
  matchNote: 'Chi nhánh khớp tên công ty tổng'
});

const gridAfterSuccess = sheetData._getGrid();
assert(gridAfterSuccess[1][testCols.STATUS - 1] === 'Đã tạo PDF', 'writeSuccess_: STATUS written correctly (Đã tạo PDF)');
assert(gridAfterSuccess[1][testCols.LINK - 1] === 'https://drive.google.com/file/d/test_ggt_123/view', 'writeSuccess_: LINK written correctly');
assert(gridAfterSuccess[1][testCols.TIME - 1] === '14/10/2026', 'writeSuccess_: TIME written correctly');
assert(gridAfterSuccess[1][testCols.TAX_NAME - 1].indexOf('VIỄN THÔNG LONG AN') !== -1, 'writeSuccess_: TAX_NAME written correctly');
assert(gridAfterSuccess[1][testCols.TAX_NOTE - 1] === 'Chi nhánh khớp tên công ty tổng', 'writeSuccess_: TAX_NOTE written correctly');

// writeError_ with taxResult preserves tax name
sandbox.writeError_(sheetData, 2, testCols, 'Lỗi xuất file PDF', {
  name: 'VIỄN THÔNG LONG AN - CHI NHÁNH TẬP ĐOÀN BƯU CHÍNH VIỄN THÔNG VIỆT NAM'
});
const gridAfterError = sheetData._getGrid();
assert(gridAfterError[1][testCols.STATUS - 1] === 'Lỗi: Lỗi xuất file PDF', 'writeError_: STATUS written with error prefix');
assert(gridAfterError[1][testCols.TAX_NOTE - 1] === 'Lỗi xuất file PDF', 'writeError_: TAX_NOTE written with error message');
assert(gridAfterError[1][testCols.TAX_NAME - 1].indexOf('VIỄN THÔNG LONG AN') !== -1, 'writeError_: TAX_NAME preserved even on error');

// 14. UI Prompt & testRow / testRowPrompt / testRow2
activeSheet = sheetData;

// Prompt successfully returns row 2
currentPromptResponse = { button: 'OK', text: '2' };
const promptRow = sandbox.promptForRowNumber_(sheetData);
assert(promptRow === 2, 'promptForRowNumber_: UI prompt returns parsed integer 2');

// Prompt cancelled by user returns null
currentPromptResponse = { button: 'CANCEL', text: '' };
const promptCancelled = sandbox.promptForRowNumber_(sheetData);
assert(promptCancelled === null, 'promptForRowNumber_: User cancellation returns null');

// Prompt out-of-range input throws clear error
currentPromptResponse = { button: 'OK', text: '99' };
let promptErrorCaught = false;
try {
  sandbox.promptForRowNumber_(sheetData);
} catch (e) {
  promptErrorCaught = true;
  assert(e.message.indexOf('Số dòng không hợp lệ: "99"') !== -1, 'promptForRowNumber_: Invalid row out-of-range throws helpful error', e.message);
}
assert(promptErrorCaught === true, 'promptForRowNumber_: Throws exception on out-of-range row');

// testRow(rowNumber) directly with argument works without prompting
let directTestWorked = false;
try {
  // Clear status on row 2, testRow will clear status and call processRow_
  // Since studentData at row 2 has valid info but external API might fail in VM, testRow validates row and attempts execution
  assert(typeof sandbox.testRow === 'function', 'testRow: function exists');
  assert(typeof sandbox.testRowPrompt === 'function', 'testRowPrompt: wrapper function exists');
  assert(typeof sandbox.testRow2 === 'function', 'testRow2: wrapper function exists');
} catch (e) {
  assert(false, 'testRow functions check failed', e.message);
}

// testRow without arguments triggers prompt
currentPromptResponse = { button: 'CANCEL', text: '' };
let noArgWorkedWithoutCrash = false;
try {
  // When cancelled, testRow should return gracefully without throwing "Số dòng không hợp lệ: undefined"
  sandbox.testRow();
  noArgWorkedWithoutCrash = true;
} catch (e) {
  assert(false, 'testRow() with no arguments threw unexpected error', e.message);
}
// Prompt non-numeric input throws error
currentPromptResponse = { button: 'OK', text: 'abc' };
let nonNumericCaught = false;
try {
  sandbox.promptForRowNumber_(sheetData);
} catch (e) {
  nonNumericCaught = true;
  assert(e.message.indexOf('Số dòng không hợp lệ: "abc"') !== -1, 'promptForRowNumber_: Non-numeric input throws helpful error', e.message);
}
assert(nonNumericCaught === true, 'promptForRowNumber_: Throws exception on non-numeric input');

// Empty sheet throws error
const emptySheet = createMockSheet(['Header1', 'Header2']);
let emptySheetCaught = false;
try {
  sandbox.promptForRowNumber_(emptySheet);
} catch (e) {
  emptySheetCaught = true;
  assert(e.message.indexOf('Sheet chưa có dữ liệu') !== -1, 'promptForRowNumber_: Empty sheet (<2 rows) throws clear error', e.message);
}
assert(emptySheetCaught === true, 'promptForRowNumber_: Throws exception on empty sheet');

// Fallback to Browser.inputBox when getUi() throws
const origGetUi = sandbox.SpreadsheetApp.getUi;
sandbox.SpreadsheetApp.getUi = () => { throw new Error('Cannot call SpreadsheetApp.getUi() from this context'); };
currentBrowserInput = '2';
const rowFromBrowser = sandbox.promptForRowNumber_(sheetData);
assert(rowFromBrowser === 2, 'promptForRowNumber_: Successfully falls back to Browser.inputBox when getUi() throws');

// Fallback to row 2 when both getUi() and Browser.inputBox() throw (headless context)
const origInputBox = sandbox.Browser.inputBox;
sandbox.Browser.inputBox = () => { throw new Error('No UI context'); };
const headlessRow = sandbox.promptForRowNumber_(sheetData);
assert(headlessRow === 2, 'promptForRowNumber_: Safely defaults to row 2 in headless/non-UI context');

// Restore original UI mocks
sandbox.SpreadsheetApp.getUi = origGetUi;
sandbox.Browser.inputBox = origInputBox;

// 15. Strict Input Validation for promptForRowNumber_
currentPromptResponse = { button: 'OK', text: '2abc' };
let alphaNumCaught = false;
try {
  sandbox.promptForRowNumber_(sheetData);
} catch (e) {
  alphaNumCaught = true;
  assert(e.message.indexOf('Số dòng không hợp lệ: "2abc"') !== -1, 'promptForRowNumber_: Alphanumeric "2abc" rejected as invalid');
}
assert(alphaNumCaught === true, 'promptForRowNumber_: Throws exception on alphanumeric input');

currentPromptResponse = { button: 'OK', text: '2.5' };
let floatCaught = false;
try {
  sandbox.promptForRowNumber_(sheetData);
} catch (e) {
  floatCaught = true;
  assert(e.message.indexOf('Số dòng không hợp lệ: "2.5"') !== -1, 'promptForRowNumber_: Float "2.5" rejected as invalid');
}
assert(floatCaught === true, 'promptForRowNumber_: Throws exception on float input');

// 16. Anti-Collision Column Tests (Link CV, Trạng thái học vụ, Thời gian gửi đơn)
const collisionHeaders = [
  'Thời gian gửi đơn', // Col 1 (Submission timestamp)
  'Họ và tên',        // Col 2
  'MSSV',             // Col 3
  'Điện thoại',       // Col 4
  'Email',            // Col 5
  'Tên doanh nghiệp', // Col 6
  'Mã số thuế',       // Col 7
  'Link CV',          // Col 8 (Should NOT be hijacked by LINK!)
  'Trạng thái học vụ' // Col 9 (Should NOT be hijacked by STATUS!)
];
const sheetCollision = createMockSheet(collisionHeaders);
const colsCollision = sandbox.getColumns_(sheetCollision);

assert(colsCollision.LINK !== 8, 'Anti-Collision: "Link CV" (col 8) is NOT hijacked as cols.LINK');
assert(colsCollision.STATUS !== 9, 'Anti-Collision: "Trạng thái học vụ" (col 9) is NOT hijacked as cols.STATUS');
assert(colsCollision.TIME !== 1, 'Anti-Collision: "Thời gian gửi đơn" (col 1) is NOT hijacked as cols.TIME');
assert(colsCollision.TIMESTAMP === 1, 'Anti-Collision: "Thời gian gửi đơn" mapped to cols.TIMESTAMP');
assert(colsCollision.LINK === 11, 'Anti-Collision: LINK created as fresh column (col 11)');
assert(colsCollision.STATUS === 10, 'Anti-Collision: STATUS created as fresh column (col 10)');

// 17. Status & Error Recovery: isAlreadyProcessed_ must NOT treat PDF errors as processed
assert(sandbox.isAlreadyProcessed_('Lỗi: Không thể tạo PDF') === false, 'isAlreadyProcessed_: "Lỗi: Không thể tạo PDF" returns false');
assert(sandbox.isAlreadyProcessed_('Lỗi tạo PDF') === false, 'isAlreadyProcessed_: "Lỗi tạo PDF" returns false');
assert(sandbox.isAlreadyProcessed_('Lỗi: mã số thuế không hợp lệ') === false, 'isAlreadyProcessed_: "Lỗi: mã số thuế..." returns false');
assert(sandbox.isAlreadyProcessed_('Error: template doc not found') === false, 'isAlreadyProcessed_: "Error: ..." returns false');
assert(sandbox.isAlreadyProcessed_('Failed to connect to Tax API') === false, 'isAlreadyProcessed_: "Failed..." returns false');
assert(sandbox.isAlreadyProcessed_('Đã tạo PDF') === true, 'isAlreadyProcessed_: "Đã tạo PDF" returns true');
assert(sandbox.isAlreadyProcessed_('Đã gửi') === true, 'isAlreadyProcessed_: "Đã gửi" returns true');

// 18. onOpen Trigger Function & Menu Registration
assert(typeof sandbox.onOpen === 'function', 'onOpen: function exists');
let onOpenRanWithoutError = false;
try {
  sandbox.onOpen();
  onOpenRanWithoutError = true;
} catch (e) {
  assert(false, 'onOpen() threw error: ' + e.message);
}
assert(onOpenRanWithoutError === true, 'onOpen: executes successfully and registers custom menu');

console.log(`\n========================================`);
console.log(`Summary: ${passCount} passed, ${failCount} failed.`);
console.log(`========================================`);

if (failCount > 0) process.exit(1);
