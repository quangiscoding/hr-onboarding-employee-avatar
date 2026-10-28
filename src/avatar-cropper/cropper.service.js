/**
 * src/avatar-cropper/cropper.service.js
 */
class AvatarCropperService {
  // Khởi tạo và hiển thị Sidebar công cụ
  static openCropperSidebar() {
    const htmlOutput = HtmlService.createTemplateFromFile(
      "avatar-cropper/cropper-sidebar",
    )
      .evaluate()
      .setTitle("✂️ Chuẩn hóa Avatar 1:1")
      .setWidth(360);
    SpreadsheetApp.getUi().showSidebar(htmlOutput);
  }

  // Trích xuất thông tin nhân sự và dữ liệu ảnh từ dòng đang chọn
  // Lấy CHÍNH XÁC theo cấu hình cột CROPPER_CONFIG, KHÔNG DÒ TÌM LUNG TUNG.
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

    // 1. Quét map tiêu đề cột chính xác (chuẩn hóa không dấu, chữ thường)
    const headerValues = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const colMap = {};
    headerValues.forEach((h, idx) => {
      if (h && String(h).trim() !== "") {
        colMap[AvatarCropperUtil.normalizeHeader(h)] = idx + 1; // 1-indexed
      }
    });

    // 2. Định vị CHÍNH XÁC vị trí các cột từ CROPPER_CONFIG
    const colKeyName = AvatarCropperUtil.normalizeHeader(
      CROPPER_CONFIG.COLUMNS.EMPLOYEE_NAME,
    );
    const colKeyPicture = AvatarCropperUtil.normalizeHeader(
      CROPPER_CONFIG.COLUMNS.PICTURE,
    );
    const colKeyEmail = AvatarCropperUtil.normalizeHeader(
      CROPPER_CONFIG.COLUMNS.EMAIL || "Email Address",
    );

    const nameColIdx = colMap[colKeyName] || -1;
    const pictureColIdx = colMap[colKeyPicture] || -1;
    const emailColIdx = colMap[colKeyEmail] || -1;

    // Bắt buộc phải có cột Tên và cột Picture trên Sheet
    if (nameColIdx === -1) {
      throw new Error(
        `Không tìm thấy cột [${CROPPER_CONFIG.COLUMNS.EMPLOYEE_NAME}] trên hàng tiêu đề của Sheet!`,
      );
    }
    if (pictureColIdx === -1) {
      throw new Error(
        `Không tìm thấy cột [${CROPPER_CONFIG.COLUMNS.PICTURE}] trên hàng tiêu đề của Sheet!`,
      );
    }

    // 3. Đọc dữ liệu dòng hiện tại
    const rowValues = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
    const fullName = String(rowValues[nameColIdx - 1] || "").trim();
    const rawPictureUrl = String(rowValues[pictureColIdx - 1] || "").trim();
    const email =
      emailColIdx !== -1 ? String(rowValues[emailColIdx - 1] || "").trim() : "";

    // Tạo sẵn tên file chuẩn từ AvatarCropperUtil (giữ nguyên dấu tiếng Việt)
    const generatedFileName = AvatarCropperUtil.buildAvatarFileName(
      { fullName, rowIndex },
      "avatar_1x1",
    );

    // Nếu chưa có link ảnh ở cột Picture
    if (!rawPictureUrl) {
      return {
        rowIndex: rowIndex,
        fullName: fullName,
        email: email,
        fileName: generatedFileName,
        rawPictureUrl: "",
        originalImageBase64: null,
        errorMsg: `Dòng #${rowIndex} (${fullName || "Nhân sự"}) chưa có link ảnh ở cột ${CROPPER_CONFIG.COLUMNS.PICTURE}!`,
      };
    }

    // 4. Trích xuất ID File Drive
    const driveFileId = AvatarCropperUtil.extractDriveId(rawPictureUrl);
    if (!driveFileId) {
      return {
        rowIndex: rowIndex,
        fullName: fullName,
        email: email,
        fileName: generatedFileName,
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
        fileName: generatedFileName,
        rawPictureUrl: rawPictureUrl,
        originalImageBase64: null,
        errorMsg: `Không thể đọc file Drive (ID: ${driveFileId}): ${err.message}`,
      };
    }

    // Trả về đầy đủ context kèm fileName chuẩn cho Sidebar sử dụng
    return {
      rowIndex: rowIndex,
      fullName: fullName,
      email: email,
      fileName: generatedFileName,
      rawPictureUrl: rawPictureUrl,
      driveFileId: driveFileId,
      originalImageBase64: originalImageBase64,
      errorMsg: null,
    };
  }

  // Lấy cấu hình folder lưu trữ hiện tại
  // Thứ tự ưu tiên: UserProperties -> CROPPER_CONFIG.STORAGE_FOLDER_ID -> Root My Drive
  static getStorageFolderConfig() {
    try {
      const userFolderId = PropertiesService.getUserProperties().getProperty(
        "AVATAR_STORAGE_FOLDER_ID",
      );
      const activeFolderId = (
        userFolderId ||
        CROPPER_CONFIG.STORAGE_FOLDER_ID ||
        ""
      ).trim();

      let folderName = "My Drive (Mặc định)";
      if (activeFolderId) {
        try {
          folderName = DriveApp.getFolderById(activeFolderId).getName();
        } catch (e) {
          folderName = "Folder ID không hợp lệ hoặc thiếu quyền";
        }
      }

      return {
        ok: true,
        folderId: activeFolderId,
        folderName: folderName,
        isCustom: Boolean(userFolderId),
      };
    } catch (err) {
      return {
        ok: false,
        folderId: "",
        folderName: "My Drive (Mặc định)",
        isCustom: false,
      };
    }
  }

  /**
   * Lưu cấu hình Folder ID vào UserProperties
   */
  static setStorageFolderConfig(folderInput) {
    try {
      const rawInput = String(folderInput || "").trim();

      // Nếu truyền chuỗi rỗng: xóa config người dùng, quay về mặc định
      if (!rawInput) {
        PropertiesService.getUserProperties().deleteProperty(
          "AVATAR_STORAGE_FOLDER_ID",
        );
        const defaultConfig = this.getStorageFolderConfig();
        return {
          ok: true,
          message: "Đã đặt lại thư mục lưu về mặc định.",
          folderId: defaultConfig.folderId,
          folderName: defaultConfig.folderName,
        };
      }

      const cleanId = AvatarCropperUtil.extractDriveId(rawInput);
      if (!cleanId) {
        return {
          ok: false,
          message: "Đường link hoặc ID thư mục Google Drive không hợp lệ!",
        };
      }

      // Xác thực quyền truy cập thư mục
      const folder = DriveApp.getFolderById(cleanId);
      PropertiesService.getUserProperties().setProperty(
        "AVATAR_STORAGE_FOLDER_ID",
        cleanId,
      );

      return {
        ok: true,
        folderId: cleanId,
        folderName: folder.getName(),
        message: `Đã đổi thư mục lưu sang: ${folder.getName()}`,
      };
    } catch (err) {
      return {
        ok: false,
        message: "Không thể truy cập thư mục: " + err.message,
      };
    }
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

    // 2. Xác định Folder lưu trữ (Ưu tiên UserProperties -> CROPPER_CONFIG -> Root)
    let folder;
    const userFolderId = PropertiesService.getUserProperties().getProperty(
      "AVATAR_STORAGE_FOLDER_ID",
    );
    const folderId = (
      userFolderId ||
      CROPPER_CONFIG.STORAGE_FOLDER_ID ||
      ""
    ).trim();

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
      folder = DriveApp.getRootFolder();
    }

    // 3. Tạo file trên Drive
    const savedFile = folder.createFile(blob);

    // 4. Cấp quyền chia sẻ
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

    // Chuẩn hóa tên cột kết quả từ CROPPER_CONFIG.COLUMNS.SQUARE_PICTURE
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

/** Bridge đọc dòng dữ liệu được chọn */
function cropperApi_getSelectedRowContext() {
  return AvatarCropperService.getSelectedRowContext();
}

/** Bridge lưu ảnh và ghi URL lên Sheet */
function cropperApi_saveCroppedAvatar(rowIndex, base64Data, fileName) {
  return AvatarCropperService.saveCroppedAvatar(rowIndex, base64Data, fileName);
}

/** Bridge đọc cấu hình folder lưu ảnh hiện tại */
function cropperApi_getStorageFolder() {
  return AvatarCropperService.getStorageFolderConfig();
}

/** Bridge cập nhật folder lưu ảnh mới */
function cropperApi_setStorageFolder(folderInput) {
  return AvatarCropperService.setStorageFolderConfig(folderInput);
}

/** Bridge tiện ích: sinh tên file avatar chuẩn hóa */
function cropperApi_getAvatarFileName(rowContext, suffix) {
  return AvatarCropperUtil.buildAvatarFileName(rowContext, suffix);
}
