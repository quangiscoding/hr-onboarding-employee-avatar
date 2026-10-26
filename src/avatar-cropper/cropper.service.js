/**
 * src/avatar-cropper/cropper.service.js
 */

class AvatarCropperService {
  /**
   * Khởi tạo và hiển thị Sidebar công cụ
   */
  static openCropperSidebar() {
    const htmlOutput = HtmlService.createTemplateFromFile(
      "avatar-cropper/cropper-sidebar",
    )
      .evaluate()
      .setTitle("✂️ Chuẩn hóa Avatar 1:1")
      .setWidth(360);

    SpreadsheetApp.getUi().showSidebar(htmlOutput);
  }

  /**
   * Trích xuất thông tin nhân sự và dữ liệu ảnh từ dòng đang chọn
   */
  static getSelectedRowContext() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const activeRange = sheet.getActiveRange();

    if (!activeRange) {
      throw new Error("Vui lòng click chọn 1 ô trên dòng nhân sự cần xử lý!");
    }

    const rowIndex = activeRange.getRow();
    if (rowIndex < 2) {
      throw new Error(
        "Bạn đang chọn dòng tiêu đề (Dòng 1). Vui lòng chọn dòng dữ liệu nhân sự!",
      );
    }

    const lastCol = sheet.getLastColumn();
    if (lastCol < 1) {
      throw new Error("Bảng tính hiện không có dữ liệu cột!");
    }

    // 1. Quét map tiêu đề cột (chuẩn hóa không dấu, chữ thường)[cite: 6]
    const headerValues = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const colMap = {};
    headerValues.forEach((h, idx) => {
      if (h && String(h).trim() !== "") {
        colMap[AvatarCropperUtil.normalizeHeader(h)] = idx + 1; // 1-indexed
      }
    });

    // 2. Định vị linh hoạt các cột cần thiết
    const possiblePictureHeaders = [
      "picture",
      "avatar",
      "anh",
      "anhdaidien",
      "photo",
      "linkanh",
      "pictureurl",
    ];
    let pictureColIdx = -1;
    for (const key of possiblePictureHeaders) {
      if (colMap[key]) {
        pictureColIdx = colMap[key];
        break;
      }
    }

    const possibleNameHeaders = [
      "hovaten",
      "fullname",
      "hoten",
      "name",
      "ungvien",
    ];
    let nameColIdx = -1;
    for (const key of possibleNameHeaders) {
      if (colMap[key]) {
        nameColIdx = colMap[key];
        break;
      }
    }

    const possibleEmailHeaders = [
      "emailaddress",
      "email",
      "personalemail",
      "workingemail",
      "emailcanhan",
    ];
    let emailColIdx = -1;
    for (const key of possibleEmailHeaders) {
      if (colMap[key]) {
        emailColIdx = colMap[key];
        break;
      }
    }

    // 3. Đọc dữ liệu dòng hiện tại
    const rowValues = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
    const fullName =
      nameColIdx !== -1 ? String(rowValues[nameColIdx - 1] || "").trim() : "";
    const email =
      emailColIdx !== -1 ? String(rowValues[emailColIdx - 1] || "").trim() : "";
    const rawPictureUrl =
      pictureColIdx !== -1
        ? String(rowValues[pictureColIdx - 1] || "").trim()
        : "";

    if (!rawPictureUrl) {
      return {
        rowIndex: rowIndex,
        fullName: fullName,
        email: email,
        rawPictureUrl: "",
        originalImageBase64: null,
        errorMsg: `Dòng #${rowIndex} (${fullName || "Nhân sự"}) chưa có link ảnh ở cột Picture / Avatar!`,
      };
    }

    // 4. Trích xuất ID File Drive
    const driveFileId = AvatarCropperUtil.extractDriveId(rawPictureUrl);
    if (!driveFileId) {
      return {
        rowIndex: rowIndex,
        fullName: fullName,
        email: email,
        rawPictureUrl: rawPictureUrl,
        originalImageBase64: null,
        errorMsg: `Không nhận diện được ID Google Drive từ link: "${rawPictureUrl}"`,
      };
    }

    // 5. Đọc Blob từ Google Drive và mã hóa Base64
    let originalImageBase64 = null;
    try {
      const file = DriveApp.getFileById(driveFileId);
      const blob = file.getBlob();
      const contentType = blob.getContentType() || "image/jpeg";
      originalImageBase64 = `data:${contentType};base64,${Utilities.base64Encode(blob.getBytes())}`;
    } catch (err) {
      return {
        rowIndex: rowIndex,
        fullName: fullName,
        email: email,
        rawPictureUrl: rawPictureUrl,
        originalImageBase64: null,
        errorMsg: `Không thể đọc file Drive (ID: ${driveFileId}): ${err.message}`,
      };
    }

    return {
      rowIndex: rowIndex,
      fullName: fullName,
      email: email,
      rawPictureUrl: rawPictureUrl,
      driveFileId: driveFileId,
      originalImageBase64: originalImageBase64,
      errorMsg: null,
    };
  }

  /**
   * Lưu ảnh Base64 vào Google Drive và ghi link vào Sheet UC06
   */
  static saveCroppedAvatar(rowIndex, base64Data, fileName) {
    if (!base64Data || !rowIndex) {
      throw new Error("Dữ liệu ảnh hoặc chỉ số dòng không hợp lệ!");
    }

    // 1. Tách chuỗi Base64 an toàn
    let base64Content = base64Data;
    if (base64Data.indexOf(",") !== -1) {
      base64Content = base64Data.split(",")[1];
    }

    const decodedBytes = Utilities.base64Decode(base64Content);
    const mimeType =
      (CROPPER_CONFIG.OUTPUT_IMAGE && CROPPER_CONFIG.OUTPUT_IMAGE.MIME_TYPE) ||
      "image/jpeg";
    const finalFileName = fileName || `avatar_row_${rowIndex}.jpg`;

    const blob = Utilities.newBlob(decodedBytes, mimeType, finalFileName);

    // 2. Xác định Folder lưu trữ (Kiểm tra chuỗi rỗng chặt chẽ để tránh Invalid argument)
    let folder;
    const folderId = (CROPPER_CONFIG.STORAGE_FOLDER_ID || "").trim();

    if (folderId !== "") {
      try {
        folder = DriveApp.getFolderById(folderId);
      } catch (err) {
        Logger.log(
          `⚠️ Không tìm thấy folder ID (${folderId}), chuyển sang lưu tại My Drive gốc: ${err.message}`,
        );
        folder = DriveApp.getRootFolder();
      }
    } else {
      // Mặc định lưu tại My Drive gốc nếu không cấu hình folder
      folder = DriveApp.getRootFolder();
    }

    // 3. Tạo file trên Drive
    const savedFile = folder.createFile(blob);

    // 4. Cấp quyền chia sẻ (bọc try-catch chống chặn quyền từ Workspace domain)
    try {
      savedFile.setSharing(
        DriveApp.Access.ANYONE_WITH_LINK,
        DriveApp.Permission.VIEW,
      );
    } catch (shareErr) {
      Logger.log(
        "⚠️ Chính sách tổ chức không cho phép set ANYONE_WITH_LINK: " +
          shareErr.message,
      );
    }

    const fileUrl = savedFile.getUrl();

    // 5. Ghi link vào Sheet UC06 tại cột "Ảnh 1:1"
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    // Chuẩn hóa tên cột kết quả từ CROPPER_CONFIG.COLUMNS.SQUARE_PICTURE ("Ảnh 1:1")
    const targetHeaderKey = AvatarCropperUtil.normalizeHeader(
      CROPPER_CONFIG.COLUMNS.SQUARE_PICTURE,
    );
    let targetColIndex = -1;

    for (let i = 0; i < headers.length; i++) {
      if (AvatarCropperUtil.normalizeHeader(headers[i]) === targetHeaderKey) {
        targetColIndex = i + 1; // 1-indexed
        break;
      }
    }

    // Nếu chưa có cột "Ảnh 1:1", tự động thêm cột mới ở cuối bảng
    if (targetColIndex === -1) {
      targetColIndex = lastCol + 1;
      sheet
        .getRange(1, targetColIndex)
        .setValue(CROPPER_CONFIG.COLUMNS.SQUARE_PICTURE);
    }

    // Cập nhật URL ảnh đã crop vào đúng dòng nhân sự
    sheet.getRange(rowIndex, targetColIndex).setValue(fileUrl);

    return {
      ok: true,
      fileUrl: fileUrl,
    };
  }
}

// =========================================================================
// GLOBAL BRIDGES CHO google.script.run
// =========================================================================
function cropperApi_getSelectedRowContext() {
  return AvatarCropperService.getSelectedRowContext();
}

function cropperApi_saveCroppedAvatar(rowIndex, base64Data, fileName) {
  return AvatarCropperService.saveCroppedAvatar(rowIndex, base64Data, fileName);
}
