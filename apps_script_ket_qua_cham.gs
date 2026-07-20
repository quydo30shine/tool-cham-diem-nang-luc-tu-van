/**
 * Kho lưu chung cho Tool chấm điểm năng lực tư vấn.
 * Mọi salesup ghi kết quả vào đây, và ai mở link cũng đọc được danh sách chung.
 *
 * CÁCH CÀI (làm 1 lần, ~5 phút):
 *  1. Vào https://sheets.new  -> đặt tên file, vd "Kết quả chấm năng lực tư vấn".
 *  2. Menu Tiện ích mở rộng (Extensions) -> Apps Script.
 *  3. Xoá hết code mẫu, dán TOÀN BỘ file này vào, bấm Lưu.
 *  4. Bấm Triển khai (Deploy) -> Bản triển khai mới (New deployment).
 *       - Loại (Select type)      : Ứng dụng web (Web app)
 *       - Thực thi với tư cách    : Tôi (Me)
 *       - Ai có quyền truy cập    : Bất kỳ ai (Anyone)      <-- BẮT BUỘC chọn mục này
 *  5. Bấm Triển khai, cấp quyền khi Google hỏi.
 *  6. Copy "URL ứng dụng web" (dạng https://script.google.com/macros/s/..../exec)
 *     và gửi lại cho tôi để nhúng vào tool.
 *
 * Lưu ý "Bất kỳ ai": người có URL này ghi/đọc được dữ liệu chấm điểm.
 * URL không bị Google index, nhưng đừng đăng công khai.
 * Muốn đổi URL sau này: Triển khai -> Quản lý bản triển khai -> tạo bản mới.
 */

var SHEET_NAME = 'ket_qua';
var SO_CAU_TOI_DA = 13; // Skinner/Supporter 13 câu, Stylist 9 câu

var COT = ['thoi_gian_cham', 'salesup', 'id_nhan_su', 'ten_nhan_su', 'vi_tri', 'level', 'salon',
           'nl1_nhan_dien', 'nl2_tu_van', 'nl3_xay_bill', 'nl4_chot_sale',
           'diem_tong_5', 'diem_tong_100', 'xep_loai']
  .concat(function () {
    var a = [];
    for (var i = 1; i <= SO_CAU_TOI_DA; i++) a.push('cau_' + i);
    return a;
  }())
  .concat(['cap_nhat_luc', 'json_raw']);

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(COT);
    sh.setFrozenRows(1);
  }
  if (sh.getLastRow() === 0) {
    sh.appendRow(COT);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Đọc toàn bộ kết quả đã chấm. */
function doGet() {
  try {
    var sh = sheet_();
    var last = sh.getLastRow();
    if (last < 2) return json_({ ok: true, rows: [] });
    var iJson = COT.indexOf('json_raw');
    var vals = sh.getRange(2, 1, last - 1, COT.length).getValues();
    var rows = [];
    for (var i = 0; i < vals.length; i++) {
      var raw = vals[i][iJson];
      if (!raw) continue;
      try { rows.push(JSON.parse(raw)); } catch (e) {}
    }
    return json_({ ok: true, rows: rows });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/** Ghi 1 kết quả (đè nếu nhân sự đó đã được chấm), hoặc xoá. */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
  } catch (err) {
    return json_({ ok: false, error: 'Hệ thống đang bận, thử lại sau.' });
  }
  try {
    var body = JSON.parse(e.postData.contents);
    var sh = sheet_();
    var iId = COT.indexOf('id_nhan_su');
    var last = sh.getLastRow();
    var ids = last > 1 ? sh.getRange(2, iId + 1, last - 1, 1).getValues() : [];

    var timDong = function (id) {
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(id)) return i + 2; // 1-based + header
      }
      return 0;
    };

    if (body.action === 'delete') {
      var d = timDong(body.id);
      if (d) sh.deleteRow(d);
      return json_({ ok: true, deleted: !!d });
    }

    var r = body.rec;
    if (!r || !r.id) return json_({ ok: false, error: 'Thiếu dữ liệu nhân sự.' });

    var g = r.gAvg || [];
    var s = r.scores || [];
    var hang = [r.at, r.salesup || '', String(r.id), r.ten || '', r.vitri || '', r.level || '', r.salon || '',
                g[0] == null ? '' : g[0], g[1] == null ? '' : g[1],
                g[2] == null ? '' : g[2], g[3] == null ? '' : g[3],
                r.on5, r.on100, r.xep_loai || ''];
    for (var i = 0; i < SO_CAU_TOI_DA; i++) hang.push(s[i] == null ? '' : s[i]);
    hang.push(new Date());
    hang.push(JSON.stringify(r));

    var dong = timDong(r.id);
    if (dong) sh.getRange(dong, 1, 1, COT.length).setValues([hang]);
    else sh.appendRow(hang);

    return json_({ ok: true, updated: !!dong });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
