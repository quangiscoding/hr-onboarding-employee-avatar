/**
 * src/main.js
 * Controller điều phối Menu hệ thống
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🖼️ Avatar Tools")
    .addItem("1. Cắt ảnh nhân sự 1:1", "menuOpenAvatarCropper")
    .addToUi();
}

/**
 * Điều hướng mở Sidebar tính năng Cắt ảnh
 */
function menuOpenAvatarCropper() {
  AvatarCropperService.openCropperSidebar();
}
