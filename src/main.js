/**
 * src/main.js
 * Controller điều phối Menu hệ thống
 */

function onOpen() {
  SpreadsheetApp.getUi().createMenu("🖼️ Avatar Tools").addItem("1. Cắt ảnh nhân sự 1:1", "menuOpenAvatarCropper").addSeparator();
  // .addItem("2. Tự động sao lưu ảnh 1:1 vào Thư mục", "menuSyncSquareAvatars")
  // .addToUi();
}

/**
 * Điều hướng mở Sidebar tính năng Cắt ảnh
 */
function menuOpenAvatarCropper() {
  AvatarCropperService.openCropperSidebar();
}

/**
 * Điều hướng chạy tính năng sao lưu ảnh gốc 1:1 vào thư mục Drive và cập nhật Sheet
 */
function menuSyncSquareAvatars() {
  const ui = SpreadsheetApp.getUi();
  try {
    const res = AvatarCropperService.syncSquareOriginalAvatars();
    ui.alert("Thông báo", res.message, ui.ButtonSet.OK);
  } catch (err) {
    ui.alert("Lỗi", "Quá trình đồng bộ thất bại: " + err.message, ui.ButtonSet.OK);
  }
}

/**
 * Hàm debug tra cứu Hera
 */
function debugLookupHera() {
  const heraId = CROPPER_CONFIG.HERA_SPREADSHEET_ID;
  const ss = SpreadsheetApp.openById(heraId);

  // 1. In ra toàn bộ danh sách sheet để xem chính xác tên tab
  const sheets = ss.getSheets().map((s) => `${s.getName()} (gid: ${s.getSheetId()})`);
  Logger.log("=== DANH SÁCH SHEET TRÊN HERA ===");
  Logger.log(sheets);

  // 2. Tìm sheet có gid = 335034815 (theo link dev của bạn)
  const targetSheet = ss.getSheets().find((s) => String(s.getSheetId()) === "335034815") || ss.getSheets()[0];
  Logger.log(">>> Đang kiểm tra tab: " + targetSheet.getName());

  // 3. In ra 3 dòng đầu tiên của tab này để xem header thực tế
  const topRows = targetSheet.getRange(1, 1, 3, 35).getValues();
  Logger.log("DÒNG 1: " + JSON.stringify(topRows[0]));
  Logger.log("DÒNG 2: " + JSON.stringify(topRows[1]));

  // 4. Tìm kiếm email của Alissa trên tab này
  const allData = targetSheet.getDataRange().getValues();
  const searchEmail = "alissawoon@gmail.com";
  let found = false;
  for (let r = 0; r < allData.length; r++) {
    for (let c = 0; c < allData[r].length; c++) {
      if (String(allData[r][c]).toLowerCase().trim() === searchEmail) {
        Logger.log(`>>> TÌM THẤY EMAIL tại Dòng ${r + 1}, Cột ${c + 1} (${AvatarCropperUtil.columnToLetter(c + 1)})`);
        Logger.log(`>>> Giá trị cột I (Cột 9 - Squad) cùng dòng là: "${allData[r][8]}"`);
        found = true;
      }
    }
  }
  if (!found) {
    Logger.log("❌ KHÔNG TÌM THẤY email " + searchEmail + " trên sheet này!");
  }
}
