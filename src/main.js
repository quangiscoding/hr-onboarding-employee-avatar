/**
 * src/main.js
 * Controller điều phối Menu hệ thống
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu("🖼️ Avatar Tools").addItem("1. ✂️ Cắt ảnh nhân sự 1:1", "menuOpenAvatarCropper").addItem("2. ⚙️ Cấu hình thư mục lưu ảnh", "menuOpenStorageSettings").addToUi();
  // .addItem("3. Tự động sao lưu ảnh 1:1 vào Thư mục", "menuSyncSquareAvatars")
}

/**
 * Điều hướng mở Sidebar tính năng Cắt ảnh
 */
function menuOpenAvatarCropper() {
  AvatarCropperService.openCropperSidebar();
}

/**
 * Điều hướng mở Dialog Cấu hình Thư mục lưu ảnh
 */
function menuOpenStorageSettings() {
  AvatarCropperService.openStorageFolderDialog();
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
