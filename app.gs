/**
 * ===================================================================
 *  CỎ MỀM — API QUẢN LÝ HÀNH CHÍNH NHÂN SỰ
 *  Backend Google Apps Script — CRUD tổng quát cho mọi sheet trong
 *  Google Sheet "hanh_chinh_nhan_su_dataset".
 *
 *  CÁCH TRIỂN KHAI (Apps Script tự lưu trữ luôn cả giao diện, chỉ dùng
 *  1 link duy nhất -> cùng origin với API, không còn cần JSONP/CORS):
 *  1. Mở Google Sheet dữ liệu -> Tiện ích mở rộng -> Apps Script.
 *  2. Xoá code mẫu, dán toàn bộ nội dung file này vào Code.gs.
 *  3. Trong project Apps Script: Tệp -> Thêm tệp -> HTML -> đặt tên đúng
 *     là "index" (không thêm đuôi .html, trình soạn thảo tự thêm).
 *     Dán toàn bộ nội dung file giao diện (index.html) vào đó.
 *  4. Sửa hằng số SHEET_ID bên dưới nếu cần (mặc định lấy ID trong
 *     đường link google_sheet_URL của dự án).
 *  5. (Tuỳ chọn) Muốn bật khoá "mật khẩu chỉnh sửa": Project Settings ->
 *     Script properties -> thêm khoá "EDIT_PASSWORD" với giá trị mật
 *     khẩu mong muốn. Để trống/không đặt thì ai cũng ghi được như cũ.
 *  6. Triển khai -> Triển khai dạng mới -> Ứng dụng web.
 *       - Người thực thi: Tôi (chủ sở hữu)
 *       - Người có quyền truy cập: Bất kỳ ai
 *  7. Mở URL /exec ngay trên trình duyệt để dùng app (đây cũng chính là
 *     API_URL bên trong file index.html — không cần sửa gì thêm vì đã
 *     tự trỏ vào chính nó).
 *  8. Mỗi lần sửa Code.gs hoặc index.html: vào Triển khai -> Quản lý
 *     triển khai -> Sửa -> Phiên bản mới -> Triển khai, để /exec cập
 *     nhật code mới nhất (chỉ Lưu (Ctrl+S) thôi thì /exec KHÔNG tự cập nhật).
 * ===================================================================
 */

const SHEET_ID = '1ofxPjVjfWLOJMRpBpsgVQOLsOgtWe0_kpbrV59vTkFk';

// Thư mục Google Drive dùng để lưu ảnh tải lên (ảnh đại diện nhân sự, ảnh
// công cụ, ảnh hoá đơn...) khi người dùng bấm "Chọn ảnh"/"Chụp ảnh" trên
// giao diện. File được set quyền chia sẻ "Bất kỳ ai có đường liên kết" ngay
// sau khi tải lên để có thể hiển thị trực tiếp trong app.
const UPLOAD_FOLDER_ID = '1dIBc1bvmfln1k29Aa8AM_4QWGk5pYwwu';

// LƯU Ý: để dùng mục "Hoá đơn" ở giao diện, cần thêm cột "hoadon_url" vào
// sheet lichsu_congcu (dán link ảnh hoá đơn được lưu công khai, vd Google Drive
// share link dạng xem trực tiếp). Backend không cần sửa gì thêm vì API là
// CRUD tổng quát theo header của từng sheet.

// MẬT KHẨU CHỈNH SỬA (Thêm/Sửa/Xoá/Tải ảnh): nếu để trống ("") thì AI CŨNG
// ghi được, giữ nguyên hành vi cũ. Muốn bật khoá, đặt 1 chuỗi mật khẩu tại
// đây (hoặc tốt hơn, đặt qua Project Settings -> Script properties, khoá
// "EDIT_PASSWORD", để không lộ mật khẩu ngay trong mã nguồn). Người xem
// (chưa nhập đúng mật khẩu) vẫn xem được toàn bộ danh sách bình thường.
const EDIT_PASSWORD = '';

function _requiredEditPassword() {
  const fromProps = PropertiesService.getScriptProperties().getProperty('EDIT_PASSWORD');
  return (fromProps && fromProps.length) ? fromProps : EDIT_PASSWORD;
}

// Trả về null nếu hợp lệ (không khoá, hoặc mật khẩu đúng); trả về thông báo
// lỗi (string) nếu bị từ chối, để hàm gọi return luôn { success:false, ... }.
function _checkEditPassword(pw) {
  const required = _requiredEditPassword();
  if (!required) return null; // chưa bật khoá chỉnh sửa
  if (String(pw || '') !== String(required)) {
    return 'Sai mật khẩu chỉnh sửa. Vui lòng nhập đúng mật khẩu để Thêm/Sửa/Xoá dữ liệu.';
  }
  return null;
}

// Khoá chính (có thể là khoá ghép, phân tách bởi dấu phẩy) cho từng sheet
const PRIMARY_KEYS = {
  nhan_vien: ['nhanvien_ma'],
  phong_ban: ['phongban_ma'],
  cong_cu: ['cong_cu_ma'],
  thongso_kythuat: ['cong_cu_ma', 'thuoc_tinh_ma'],
  thuoc_tinh: ['thuoc_tinh_ma'],
  lichsu_congcu: ['lich_su_ma'],
  danh_muc: ['danh_muc_ma'],
  danhmuc_chitiet: ['danh_muc_ma'],
  qua_tang: ['nhanvien_ma', 'ma_qua_tang'],
  bien_ban: ['bienban_ma']
};

// Header mặc định để TỰ ĐỘNG TẠO sheet "bien_ban" (lưu lại Biên bản bàn giao
// đã lập) nếu sheet dữ liệu chưa có sẵn sheet này — người dùng không cần tự
// tạo sheet + cột bằng tay.
const AUTO_SHEET_HEADERS = {
  bien_ban: ['bienban_ma', 'so_bb', 'ngay_lap', 'nguoi_lap', 'don_vi', 'nguoi_phu_trach',
    'ben_giao', 'ben_nhan', 'dia_diem', 'danh_sach_json', 'ghi_chu']
};

function _ss() {
  return SpreadsheetApp.openById(SHEET_ID);
}

function _sheet(name) {
  let sh = _ss().getSheetByName(name);
  if (!sh && AUTO_SHEET_HEADERS[name]) {
    // Sheet phục vụ tính năng mới (vd lưu biên bản bàn giao) chưa tồn tại
    // trong Google Sheet dữ liệu -> tự tạo kèm dòng tiêu đề, để người dùng
    // không phải thao tác thủ công trên Sheet trước khi dùng tính năng.
    sh = _ss().insertSheet(name);
    sh.getRange(1, 1, 1, AUTO_SHEET_HEADERS[name].length).setValues([AUTO_SHEET_HEADERS[name]]);
  }
  if (!sh) throw new Error('Không tìm thấy sheet: ' + name);
  return sh;
}

function _headers(sh) {
  const lastCol = sh.getLastColumn();
  return sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
}

function _rowsAsObjects(sh) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const headers = _headers(sh);
  if (lastRow < 2) return { headers: headers, rows: [] };
  const values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const rows = values.map((row, idx) => {
    const obj = {};
    headers.forEach((h, i) => {
      let v = row[i];
      if (v instanceof Date) {
        v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy');
      }
      obj[h] = v === '' ? null : v;
    });
    obj._row = idx + 2; // số dòng thực tế trong sheet (1-indexed, có header)
    return obj;
  });
  return { headers: headers, rows: rows };
}

function _matchesKey(obj, keyFields, keyValues) {
  return keyFields.every((f, i) => String(obj[f] ?? '').trim() === String(keyValues[i] ?? '').trim());
}

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Trả kết quả dạng JSONP: bọc trong lời gọi hàm callback(...) thay vì JSON thuần.
// Sở dĩ cần dùng cách này: khi người dùng mở file index.html trực tiếp bằng
// double-click (địa chỉ dạng file://), trình duyệt gắn "origin: null" cho mọi
// yêu cầu fetch()/XMLHttpRequest, và Google chặn hẳn CORS cho origin null ở
// tầng hạ tầng (không có cách nào cấu hình lại từ phía Apps Script để cho
// qua). Thẻ <script src="..."> thì KHÔNG bị luật CORS áp dụng (nó không phải
// request kiểu fetch/XHR) nên gọi được xuyên origin kể cả từ file:// -> đây
// là cách JSONP truyền thống để lách giới hạn này.
function _jsonp(callback, obj) {
  const safeCallback = String(callback).replace(/[^a-zA-Z0-9_$]/g, '');
  return ContentService
    .createTextOutput(safeCallback + '(' + JSON.stringify(obj) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

/** ============== ENTRY POINTS ============== */

function doGet(e) {
  const params = (e && e.parameter) || {};

  // Không có action & không có sheet -> đây là yêu cầu MỞ TRANG WEB
  // (người dùng truy cập trực tiếp link /exec), trả về giao diện HTML.
  // Nhờ vậy trang web và API cùng chạy trên 1 địa chỉ (script.google.com),
  // trình duyệt sẽ KHÔNG còn coi đây là 2 "origin" khác nhau -> hết lỗi CORS
  // "Access-Control-Allow-Origin" / "from origin 'null'" khi mở file .html
  // trực tiếp bằng double-click (file://) trước đây.
  // LƯU Ý: cần có 1 file HTML tên đúng là "index" trong project Apps Script
  // (Tệp -> Thêm tệp -> HTML -> đặt tên "index"), dán nội dung file giao diện vào đó.
  if (!params.action && !params.sheet) {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Cỏ Mềm — Hành chính · Nhân sự')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  try {
    const action = (params.action || 'list').toLowerCase();
    const sheetName = params.sheet;
    // Nếu có tham số "callback" -> trả kết quả dạng JSONP (gọi qua thẻ <script>,
    // dùng khi mở file index.html trực tiếp bằng double-click, xem giải thích ở _jsonp).
    // Ngược lại (gọi qua fetch() bình thường khi đã deploy /exec) -> trả JSON thuần như cũ.
    const respond = (obj) => params.callback ? _jsonp(params.callback, obj) : _json(obj);

    if (action === 'schema') {
      return respond({ success: true, data: _schema() });
    }

    if (action === 'ping') {
      return respond({ success: true, data: _ping() });
    }

    if (!sheetName) throw new Error('Thiếu tham số sheet');
    const sh = _sheet(sheetName);

    if (action === 'list') {
      const { rows } = _rowsAsObjects(sh);
      // Phân trang phía server (tuỳ chọn): nếu FE gửi kèm page & pageSize thì
      // chỉ trả về đúng 1 trang dữ liệu (đỡ tải cả sheet lớn mỗi lần chuyển
      // trang), kèm "total" để FE tự tính số trang. Nếu không gửi 2 tham số
      // này (như trước đây) thì trả về TOÀN BỘ danh sách như cũ, không phá vỡ
      // các màn hình đang cần load hết dữ liệu để lọc/tổng hợp/tra cứu.
      const pageSize = parseInt(params.pageSize, 10);
      if (pageSize > 0) {
        const page = Math.max(1, parseInt(params.page, 10) || 1);
        const start = (page - 1) * pageSize;
        const pageRows = rows.slice(start, start + pageSize);
        return respond({ success: true, data: pageRows, total: rows.length, page: page, pageSize: pageSize });
      }
      return respond({ success: true, data: rows, total: rows.length });
    }

    if (action === 'nextcode') {
      const keyFields = PRIMARY_KEYS[sheetName] || [];
      const code = _nextCode(sh, keyFields);
      return respond({ success: true, data: { code: code } });
    }

    if (action === 'get') {
      const keyFields = PRIMARY_KEYS[sheetName] || ['_row'];
      const keyValues = String(e.parameter.id || '').split('|');
      const { rows } = _rowsAsObjects(sh);
      const found = rows.find(r => _matchesKey(r, keyFields, keyValues));
      if (!found) return respond({ success: false, message: 'Không tìm thấy bản ghi' });
      return respond({ success: true, data: found });
    }

    // Thêm/Sửa/Xoá qua GET: chỉ dùng khi gọi bằng kỹ thuật JSONP (script tag),
    // vì thẻ <script> chỉ gọi được phương thức GET. Vẫn khoá ghi (LockService)
    // giống hệt doPost để tránh 2 người cùng sửa 1 lúc làm lệch dòng.
    if (action === 'create' || action === 'update' || action === 'delete') {
      const lock = LockService.getScriptLock();
      try {
        lock.waitLock(15000);
      } catch (lockErr) {
        return respond({ success: false, message: 'Hệ thống đang bận xử lý yêu cầu khác, vui lòng thử lại sau ít giây.' });
      }
      try {
        const pwErr = _checkEditPassword(params.editPassword);
        if (pwErr) return respond({ success: false, message: pwErr, authError: true });
        const headers = _headers(sh);
        const keyFields = PRIMARY_KEYS[sheetName] || ['_row'];
        const data = params.data ? JSON.parse(params.data) : {};
        let result;
        if (action === 'create') {
          result = _create(sh, headers, data);
        } else if (action === 'update') {
          const keyValues = String(params.id || '').split('|');
          result = _update(sh, headers, keyFields, keyValues, data);
        } else {
          const keyValues = String(params.id || '').split('|');
          result = _delete(sh, headers, keyFields, keyValues);
        }
        return respond(result);
      } finally {
        lock.releaseLock();
      }
    }

    throw new Error('Action không hợp lệ: ' + action);
  } catch (err) {
    const respond = (obj) => params.callback ? _jsonp(params.callback, obj) : _json(obj);
    return respond({ success: false, message: err.message });
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Không nhận được dữ liệu gửi lên (postData rỗng)');
    }
    const body = JSON.parse(e.postData.contents);
    const action = (body.action || '').toLowerCase();

    // Mọi thao tác ghi (kể cả tải ảnh) đều yêu cầu đúng mật khẩu chỉnh sửa
    // nếu tính năng khoá đang được bật (xem _requiredEditPassword). Kiểm tra
    // ngay từ đầu, trước khi đụng tới Drive/Sheet.
    const pwErr = _checkEditPassword(body.editPassword);
    if (pwErr) return _json({ success: false, message: pwErr, authError: true });

    // Tải ảnh lên Google Drive: xử lý riêng, KHÔNG cần khoá sheet (LockService)
    // vì bước này chỉ ghi file vào Drive, chưa đụng tới hàng/cột nào trong
    // Google Sheet -> không tranh chấp với các thao tác CRUD khác.
    if (action === 'uploadimage') {
      return _json(_uploadImage(body));
    }

    const lock = LockService.getScriptLock();
    try {
      // Chờ tối đa 15s để lấy khoá ghi, tránh 2 người cùng sửa 1 lúc làm lệch dòng
      lock.waitLock(15000);
    } catch (lockErr) {
      return _json({ success: false, message: 'Hệ thống đang bận xử lý yêu cầu khác, vui lòng thử lại sau ít giây.' });
    }
    try {
      const sheetName = body.sheet;
      if (!sheetName) throw new Error('Thiếu tham số sheet');
      const sh = _sheet(sheetName);
      const headers = _headers(sh);
      const keyFields = PRIMARY_KEYS[sheetName] || ['_row'];

      if (action === 'create') {
        return _json(_create(sh, headers, body.data || {}));
      }
      if (action === 'update') {
        const keyValues = String(body.id || '').split('|');
        return _json(_update(sh, headers, keyFields, keyValues, body.data || {}));
      }
      if (action === 'delete') {
        const keyValues = String(body.id || '').split('|');
        return _json(_delete(sh, headers, keyFields, keyValues));
      }
      throw new Error('Action không hợp lệ: ' + action);
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return _json({ success: false, message: err.message });
  }
}

/**
 * Nhận ảnh dạng base64 từ giao diện (chọn file trên máy hoặc chụp trực tiếp
 * bằng camera), lưu vào thư mục Drive UPLOAD_FOLDER_ID, bật chia sẻ "Bất kỳ
 * ai có đường liên kết" rồi trả về link để lưu vào cột kiểu *_url trong sheet
 * (vd photo_url, hoadon_url). body = { filename, mimeType, base64 }.
 */
/**
 * HÀM TẠM DÙNG 1 LẦN — chỉ để ép Google hiện đúng hộp thoại xin quyền Drive
 * (popup chuẩn) ngay trong trình soạn thảo, thay vì để nó tự bật lên qua
 * link /exec (dễ vỡ do CSP/extension chặn, gây lỗi "Unexpected identifier $").
 * Cách dùng: Lưu file -> ở thanh trên trình soạn thảo chọn hàm
 * "authorizeDriveAccess" trong dropdown -> bấm nút ▶ Chạy -> một popup xin
 * quyền hiện ra -> chọn tài khoản -> nếu thấy "Ứng dụng chưa xác minh" thì
 * bấm "Nâng cao" (Advanced) -> "Đi tới ... (không an toàn)" -> "Cho phép".
 * Sau khi chạy xong (Log hiện dòng "OK - đã cấp quyền Drive"), có thể xoá
 * hàm này đi rồi Triển khai -> Quản lý triển khai -> Sửa -> Phiên bản mới.
 */
function _authorizeDriveAccess() {
  const folder = DriveApp.getFolderById(UPLOAD_FOLDER_ID);
  Logger.log('OK - đã cấp quyền Drive, thư mục: ' + folder.getName());
}

function _uploadImage(body) {
  if (!body || !body.base64) throw new Error('Thiếu dữ liệu ảnh (base64)');
  let folder;
  try {
    folder = DriveApp.getFolderById(UPLOAD_FOLDER_ID);
  } catch (e) {
    throw new Error('Không mở được thư mục Google Drive lưu ảnh. Kiểm tra lại UPLOAD_FOLDER_ID hoặc quyền truy cập thư mục.');
  }
  const mimeType = body.mimeType || 'image/jpeg';
  const ext = (mimeType.split('/')[1] || 'jpg').replace('jpeg', 'jpg');
  const safeName = body.filename
    ? String(body.filename).replace(/[^\w.\-]/g, '_')
    : ('anh_' + Date.now() + '.' + ext);
  let bytes;
  try {
    bytes = Utilities.base64Decode(body.base64);
  } catch (e) {
    throw new Error('Dữ liệu ảnh gửi lên không hợp lệ (giải mã base64 thất bại).');
  }
  const blob = Utilities.newBlob(bytes, mimeType, safeName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const id = file.getId();
  return {
    success: true,
    data: {
      fileId: id,
      // Link "xem trực tiếp" dạng chia sẻ chuẩn của Drive - đồng nhất với
      // link mà người dùng tự dán tay từ trước tới nay (driveDirectLink ở
      // phía giao diện sẽ tự chuyển link này sang dạng ảnh nhúng được).
      url: 'https://drive.google.com/file/d/' + id + '/view?usp=sharing',
      directUrl: 'https://drive.google.com/thumbnail?id=' + id + '&sz=w1000'
    }
  };
}

/**
 * Ping đơn giản để kiểm tra deployment còn sống và đúng phiên bản mới nhất
 * hay không (không đọc/ghi sheet). Gọi bằng GET: ?action=ping
 */
function _ping() {
  return { ok: true, time: new Date().toISOString() };
}

// Sinh mã mới nối tiếp mã lớn nhất hiện có, giữ nguyên tiền tố chữ và độ dài số
// (vd NV000001, NV000002...). Chỉ áp dụng cho sheet có khoá chính đơn (1 cột).
function _nextCode(sh, keyFields) {
  if (!keyFields || keyFields.length !== 1) return null;
  const field = keyFields[0];
  const { rows } = _rowsAsObjects(sh);
  let maxNum = -1, prefix = '', width = 1;
  rows.forEach(r => {
    const v = String(r[field] ?? '').trim();
    const m = v.match(/^([^\d]*)(\d+)$/);
    if (m) {
      const num = parseInt(m[2], 10);
      if (num > maxNum) { maxNum = num; prefix = m[1]; width = m[2].length; }
    }
  });
  if (maxNum === -1) return null; // chưa có dữ liệu mẫu -> không tự suy ra được tiền tố
  const nextNum = maxNum + 1;
  const nextStr = String(nextNum);
  return prefix + nextStr.padStart(Math.max(width, nextStr.length), '0');
}

function _schema() {
  const out = {};
  _ss().getSheets().forEach(sh => {
    out[sh.getName()] = { headers: _headers(sh), primaryKey: PRIMARY_KEYS[sh.getName()] || [] };
  });
  return out;
}

// ================= VALIDATE DỮ LIỆU ĐẦU VÀO =================
// Kiểm tra định dạng email/số điện thoại và tính hợp lý của các cặp ngày
// tháng có liên quan (vd ngày nghỉ việc phải sau ngày vào làm). Áp dụng
// CHUNG cho mọi sheet dựa theo TÊN CỘT (không cần khai báo riêng cho từng
// sheet) — cột nào có tên "email"/"phone"/"ngay_..." thì tự được kiểm tra
// nếu người dùng có nhập giá trị. Trả về mảng chuỗi lỗi (rỗng = hợp lệ).
const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
// SĐT Việt Nam: 10 số bắt đầu bằng 0, hoặc dạng +84 theo sau 9 số
const PHONE_REGEX = /^(0\d{9,10}|\+84\d{9,10})$/;

function _parseVnDate(v) {
  const s = String(v || '').trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return null;
  const d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  return isNaN(d.getTime()) ? null : d;
}

function _validateData(sheetName, data) {
  const errors = [];
  if (data.email != null && String(data.email).trim() !== '' && !EMAIL_REGEX.test(String(data.email).trim())) {
    errors.push('Email không đúng định dạng: ' + data.email);
  }
  if (data.phone != null && String(data.phone).trim() !== '' && !PHONE_REGEX.test(String(data.phone).trim().replace(/[.\s-]/g, ''))) {
    errors.push('Số điện thoại không hợp lệ: ' + data.phone);
  }
  // Các cặp ngày cần đảm bảo ngày sau >= ngày trước (chỉ kiểm tra khi cả 2 đều có giá trị hợp lệ)
  const DATE_PAIRS = [
    ['ngay_thu_viec', 'ngay_chinh_thuc', 'Ngày chính thức phải từ sau Ngày thử việc trở đi'],
    ['ngay_thu_viec', 'ngay_nghi_viec', 'Ngày nghỉ việc phải từ sau Ngày thử việc trở đi'],
    ['ngay_nhan', 'ngay_sua', 'Ngày sửa phải từ sau Ngày nhận trở đi']
  ];
  DATE_PAIRS.forEach(([fA, fB, msg]) => {
    if (data[fA] == null || data[fB] == null) return;
    const dA = _parseVnDate(data[fA]);
    const dB = _parseVnDate(data[fB]);
    if (dA && dB && dB < dA) errors.push(msg);
  });
  if (data.chi_phi_sua != null && String(data.chi_phi_sua).trim() !== '') {
    const n = Number(String(data.chi_phi_sua).replace(/[^\d.-]/g, ''));
    if (isNaN(n) || n < 0) errors.push('Chi phí sửa phải là một số không âm');
  }
  return errors;
}

function _create(sh, headers, data) {
  const errors = _validateData(sh.getName(), data);
  if (errors.length) return { success: false, message: errors.join('; ') };
  // Chặn trùng khoá chính nếu có
  const sheetName = sh.getName();
  const keyFields = PRIMARY_KEYS[sheetName];
  if (keyFields && keyFields.length) {
    const { rows } = _rowsAsObjects(sh);
    const dup = rows.find(r => keyFields.every(f => String(r[f] ?? '').trim() === String(data[f] ?? '').trim() && String(data[f] ?? '') !== ''));
    if (dup) return { success: false, message: 'Mã đã tồn tại: ' + keyFields.map(f => data[f]).join('|') };
  }
  const rowArr = headers.map(h => (data[h] !== undefined ? data[h] : ''));
  sh.appendRow(rowArr);
  return { success: true, message: 'Đã thêm mới thành công' };
}

function _findRowIndex(sh, headers, keyFields, keyValues) {
  const { rows } = _rowsAsObjects(sh);
  const found = rows.find(r => _matchesKey(r, keyFields, keyValues));
  return found ? found._row : -1;
}

function _update(sh, headers, keyFields, keyValues, data) {
  const errors = _validateData(sh.getName(), data);
  if (errors.length) return { success: false, message: errors.join('; ') };
  const rowIndex = _findRowIndex(sh, headers, keyFields, keyValues);
  if (rowIndex === -1) return { success: false, message: 'Không tìm thấy bản ghi để cập nhật' };
  const currentValues = sh.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  const newValues = headers.map((h, i) => (data[h] !== undefined ? data[h] : currentValues[i]));
  sh.getRange(rowIndex, 1, 1, headers.length).setValues([newValues]);
  return { success: true, message: 'Đã cập nhật thành công' };
}

function _delete(sh, headers, keyFields, keyValues) {
  const rowIndex = _findRowIndex(sh, headers, keyFields, keyValues);
  if (rowIndex === -1) return { success: false, message: 'Không tìm thấy bản ghi để xoá' };
  sh.deleteRow(rowIndex);
  return { success: true, message: 'Đã xoá thành công' };
}
