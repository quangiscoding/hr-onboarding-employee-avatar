/**
 * src/avatar-cropper/cropper.service.js
 */
class AvatarCropperService {
  // Khởi tạo và hiển thị Sidebar công cụ
  static openCropperSidebar() {
    const htmlOutput = HtmlService.createTemplateFromFile("avatar-cropper/cropper-sidebar").evaluate().setTitle("✂️ Chuẩn hóa Avatar 1:1").setWidth(360);
    SpreadsheetApp.getUi().showSidebar(htmlOutput);
  }

  // Trích xuất thông tin nhân sự và dữ liệu ảnh từ dòng đang chọn
  static getSelectedRowContext() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const activeRange = sheet.getActiveRange();
    if (!activeRange) {
      throw new Error("Vui lòng click chọn 1 ô trên dòng nhân sự cần xử lý!");
    }

    const rowIndex = activeRange.getRow();
    if (rowIndex < 2) {
      throw new Error("Bạn đang chọn dòng tiêu đề (Dòng 1). Vui lòng chọn dòng dữ liệu nhân sự!");
    }

    const lastCol = sheet.getLastColumn();
    if (lastCol < 1) {
      throw new Error("Bảng tính hiện không có dữ liệu cột!");
    }

    // 1. Quét map tiêu đề cột chính xác
    const headerValues = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const colMap = {};
    headerValues.forEach((h, idx) => {
      if (h && String(h).trim() !== "") {
        colMap[AvatarCropperUtil.normalizeHeader(h)] = idx + 1; // 1-indexed
      }
    });

    // 2. Định vị vị trí các cột từ CROPPER_CONFIG
    const colKeyCode = AvatarCropperUtil.normalizeHeader(CROPPER_CONFIG.COLUMNS.CODE || "Code");
    const colKeyName = AvatarCropperUtil.normalizeHeader(CROPPER_CONFIG.COLUMNS.EMPLOYEE_NAME);
    const colKeyPicture = AvatarCropperUtil.normalizeHeader(CROPPER_CONFIG.COLUMNS.PICTURE);
    const colKeyEmail = AvatarCropperUtil.normalizeHeader(CROPPER_CONFIG.COLUMNS.EMAIL || "Email Address");

    const codeColIdx = colMap[colKeyCode] || -1;
    const nameColIdx = colMap[colKeyName] || -1;
    const pictureColIdx = colMap[colKeyPicture] || -1;
    const emailColIdx = colMap[colKeyEmail] || -1;

    if (nameColIdx === -1) {
      throw new Error(`Không tìm thấy cột [${CROPPER_CONFIG.COLUMNS.EMPLOYEE_NAME}] trên hàng tiêu đề của Sheet!`);
    }
    if (pictureColIdx === -1) {
      throw new Error(`Không tìm thấy cột [${CROPPER_CONFIG.COLUMNS.PICTURE}] trên hàng tiêu đề của Sheet!`);
    }

    // 3. Đọc dữ liệu dòng hiện tại
    const rowValues = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
    const employeeCode = codeColIdx !== -1 ? String(rowValues[codeColIdx - 1] || "").trim() : "";
    const fullName = String(rowValues[nameColIdx - 1] || "").trim();
    const rawPictureUrl = String(rowValues[pictureColIdx - 1] || "").trim();
    const email = emailColIdx !== -1 ? String(rowValues[emailColIdx - 1] || "").trim() : "";

    const generatedFileName = AvatarCropperUtil.buildAvatarFileName({ fullName, rowIndex }, "avatar_1x1");

    if (!rawPictureUrl) {
      return {
        rowIndex: rowIndex,
        employeeCode: employeeCode,
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
        employeeCode: employeeCode,
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
        employeeCode: employeeCode,
        fullName: fullName,
        email: email,
        fileName: generatedFileName,
        rawPictureUrl: rawPictureUrl,
        originalImageBase64: null,
        errorMsg: `Không thể đọc file Drive (ID: ${driveFileId}): ${err.message}`,
      };
    }

    return {
      rowIndex: rowIndex,
      employeeCode: employeeCode,
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
  static getStorageFolderConfig() {
    try {
      const userFolderId = PropertiesService.getUserProperties().getProperty("AVATAR_STORAGE_FOLDER_ID");
      const activeFolderId = (userFolderId || CROPPER_CONFIG.STORAGE_FOLDER_ID || "").trim();
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

  static setStorageFolderConfig(folderInput) {
    try {
      const rawInput = String(folderInput || "").trim();
      const ss = SpreadsheetApp.getActiveSpreadsheet();

      // 1. Trường hợp xóa / reset folder về mặc định
      if (!rawInput) {
        PropertiesService.getUserProperties().deleteProperty("AVATAR_STORAGE_FOLDER_ID");

        // Xóa Developer Metadata cũ trên file nếu có
        const existingMeta = ss.getDeveloperMetadata().find((m) => m.getKey() === "AVATAR_STORAGE_FOLDER_ID");
        if (existingMeta) {
          existingMeta.remove();
        }

        const defaultConfig = this.getStorageFolderConfig();
        return {
          ok: true,
          message: "Đã đặt lại thư mục lưu về mặc định.",
          folderId: defaultConfig.folderId,
          folderName: defaultConfig.folderName,
        };
      }

      // 2. Validate và trích xuất Folder ID
      const cleanId = AvatarCropperUtil.extractDriveId(rawInput);
      if (!cleanId) {
        return {
          ok: false,
          message: "Đường link hoặc ID thư mục Google Drive không hợp lệ!",
        };
      }

      const folder = DriveApp.getFolderById(cleanId);

      // Lưu vào UserProperties cho tool local
      PropertiesService.getUserProperties().setProperty("AVATAR_STORAGE_FOLDER_ID", cleanId);

      // GHI VÀO DEVELOPER METADATA CỦA FILE SPREADSHEET (DOCUMENT scope để các script khác đọc được)
      const existingMeta = ss.getDeveloperMetadata().find((m) => m.getKey() === "AVATAR_STORAGE_FOLDER_ID");

      if (existingMeta) {
        existingMeta.setValue(cleanId);
      } else {
        ss.addDeveloperMetadata("AVATAR_STORAGE_FOLDER_ID", cleanId, SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT);
      }

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
   * Lưu ảnh Base64 vào Google Drive theo cấu trúc: Root -> Squad/Unit -> [<Code>] <FullName>
   * và ghi link vào Sheet UC06
   */
  static saveCroppedAvatar(rowIndex, base64Data, fileName) {
    if (!base64Data || !rowIndex) {
      throw new Error("Dữ liệu ảnh hoặc chỉ số dòng không hợp lệ!");
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const lastCol = sheet.getLastColumn();
    const headerValues = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const colMap = {};
    headerValues.forEach((h, idx) => {
      if (h && String(h).trim() !== "") {
        colMap[AvatarCropperUtil.normalizeHeader(h)] = idx + 1;
      }
    });

    const codeColIdx = colMap[AvatarCropperUtil.normalizeHeader(CROPPER_CONFIG.COLUMNS.CODE || "Code")] || -1;
    const nameColIdx = colMap[AvatarCropperUtil.normalizeHeader(CROPPER_CONFIG.COLUMNS.EMPLOYEE_NAME)] || -1;
    const emailColIdx = colMap[AvatarCropperUtil.normalizeHeader(CROPPER_CONFIG.COLUMNS.EMAIL || "Email Address")] || -1;

    const rowValues = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
    const employeeCode = codeColIdx !== -1 ? String(rowValues[codeColIdx - 1] || "").trim() : "";
    const fullName = nameColIdx !== -1 ? String(rowValues[nameColIdx - 1] || "").trim() : "";
    const email = emailColIdx !== -1 ? String(rowValues[emailColIdx - 1] || "").trim() : "";

    // 1. Tách chuỗi Base64
    let base64Content = base64Data;
    if (base64Data.indexOf(",") !== -1) {
      base64Content = base64Data.split(",")[1];
    }
    const decodedBytes = Utilities.base64Decode(base64Content);
    const mimeType = (CROPPER_CONFIG.OUTPUT_IMAGE && CROPPER_CONFIG.OUTPUT_IMAGE.MIME_TYPE) || "image/jpeg";
    const finalFileName = fileName || `avatar_${employeeCode || rowIndex}.jpg`;
    const blob = Utilities.newBlob(decodedBytes, mimeType, finalFileName);

    // 2. Xác định Thư mục gốc (Root Folder)
    let rootFolder;
    const userFolderId = PropertiesService.getUserProperties().getProperty("AVATAR_STORAGE_FOLDER_ID");
    const folderId = (userFolderId || CROPPER_CONFIG.STORAGE_FOLDER_ID || "").trim();

    if (folderId !== "") {
      try {
        rootFolder = DriveApp.getFolderById(folderId);
      } catch (err) {
        Logger.log(`⚠️ Không tìm thấy folder ID (${folderId}), chuyển sang Root My Drive: ${err.message}`);
        rootFolder = DriveApp.getRootFolder();
      }
    } else {
      rootFolder = DriveApp.getRootFolder();
    }

    // 3. TẠO HOẶC ĐIỀU HƯỚNG CÂY FOLDER:
    // Cấp 1: Tra cứu Squad/Unit từ sheet Hera theo Email cá nhân
    const squadName = AvatarCropperUtil.lookupSquadFromHera(email);
    const squadFolder = AvatarCropperUtil.getOrCreateFolder(rootFolder, squadName);

    // Cấp 2: Folder con đặt tên: [<Employee Code>] <Employee FullName>
    const candidateFolderName = employeeCode ? `[${employeeCode}] ${fullName}` : fullName || `Dòng_${rowIndex}`;
    const targetFolder = AvatarCropperUtil.getOrCreateFolder(squadFolder, candidateFolderName);

    // 4. Tạo file trong Folder đích
    const savedFile = targetFolder.createFile(blob);

    try {
      savedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareErr) {
      Logger.log("⚠️ Chính sách tổ chức chặn ANYONE_WITH_LINK: " + shareErr.message);
    }
    const fileUrl = savedFile.getUrl();

    // 5. Ghi link vào Sheet UC06 tại cột "Ảnh 1:1"
    const targetHeaderKey = AvatarCropperUtil.normalizeHeader(CROPPER_CONFIG.COLUMNS.SQUARE_PICTURE);
    let targetColIndex = -1;

    for (let i = 0; i < headerValues.length; i++) {
      if (AvatarCropperUtil.normalizeHeader(headerValues[i]) === targetHeaderKey) {
        targetColIndex = i + 1;
        break;
      }
    }

    if (targetColIndex === -1) {
      targetColIndex = lastCol + 1;
      sheet.getRange(1, targetColIndex).setValue(CROPPER_CONFIG.COLUMNS.SQUARE_PICTURE);
    }

    sheet.getRange(rowIndex, targetColIndex).setValue(fileUrl);

    return {
      ok: true,
      fileUrl: fileUrl,
      folderPath: `${squadName} / ${candidateFolderName}`,
    };
  }

  /**
   * Quét toàn bộ sheet UC06:
   * Nếu ảnh gốc (cột Picture) đã là 1:1 và cột "Ảnh 1:1" chưa có link,
   * tự động tạo bản sao vào đúng thư mục Drive và cập nhật link vào cột "Ảnh 1:1".
   */
  static syncSquareOriginalAvatars() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2) return { ok: true, count: 0, message: "Sheet chưa có dữ liệu." };

    // 1. Quét map tiêu đề cột
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const colMap = {};
    headers.forEach((h, idx) => {
      if (h && String(h).trim() !== "") {
        colMap[AvatarCropperUtil.normalizeHeader(h)] = idx + 1;
      }
    });

    const codeCol = colMap[AvatarCropperUtil.normalizeHeader(CROPPER_CONFIG.COLUMNS.CODE || "Code")] || -1;
    const nameCol = colMap[AvatarCropperUtil.normalizeHeader(CROPPER_CONFIG.COLUMNS.EMPLOYEE_NAME)] || -1;
    const emailCol = colMap[AvatarCropperUtil.normalizeHeader(CROPPER_CONFIG.COLUMNS.EMAIL || "Email Address")] || -1;
    const pictureCol = colMap[AvatarCropperUtil.normalizeHeader(CROPPER_CONFIG.COLUMNS.PICTURE)] || -1;
    const checkCol = colMap[AvatarCropperUtil.normalizeHeader(CROPPER_CONFIG.COLUMNS.CHECK_SQUARE || "Check 1:1")] || -1;
    let squareCol = colMap[AvatarCropperUtil.normalizeHeader(CROPPER_CONFIG.COLUMNS.SQUARE_PICTURE || "Ảnh 1:1")] || -1;

    if (pictureCol === -1) {
      throw new Error(`Không tìm thấy cột [${CROPPER_CONFIG.COLUMNS.PICTURE}] trên sheet!`);
    }

    // Tự động tạo cột "Ảnh 1:1" nếu trên sheet chưa có
    if (squareCol === -1) {
      squareCol = lastCol + 1;
      sheet.getRange(1, squareCol).setValue(CROPPER_CONFIG.COLUMNS.SQUARE_PICTURE);
    }

    // 2. Xác định Root Folder lưu trữ theo UserProperties hoặc Config
    let rootFolder;
    const userFolderId = PropertiesService.getUserProperties().getProperty("AVATAR_STORAGE_FOLDER_ID");
    const folderId = (userFolderId || CROPPER_CONFIG.STORAGE_FOLDER_ID || "").trim();
    if (folderId) {
      try {
        rootFolder = DriveApp.getFolderById(folderId);
      } catch (e) {
        rootFolder = DriveApp.getRootFolder();
      }
    } else {
      rootFolder = DriveApp.getRootFolder();
    }

    // 3. Đọc dữ liệu từ dòng 2
    const numRows = lastRow - 1;
    const totalCols = Math.max(lastCol, squareCol);
    const data = sheet.getRange(2, 1, numRows, totalCols).getValues();

    let processedCount = 0;

    data.forEach((row, i) => {
      const rowIndex = i + 2;
      const pictureUrl = pictureCol !== -1 ? String(row[pictureCol - 1] || "").trim() : "";
      const squareUrl = squareCol !== -1 ? String(row[squareCol - 1] || "").trim() : "";

      // Chỉ xử lý các ứng viên có ảnh gốc nhưng chưa có link ở cột "Ảnh 1:1"
      if (pictureUrl && !squareUrl) {
        const driveId = AvatarCropperUtil.extractDriveId(pictureUrl);
        if (!driveId) return;

        // Gọi hàm kiểm tra kích thước trực tiếp từ AvatarCropperUtil
        const dimensions = AvatarCropperUtil.getImageDimensions(driveId);
        if (dimensions && dimensions.isSquare) {
          const empCode = codeCol !== -1 ? String(row[codeCol - 1] || "").trim() : "";
          const fullName = nameCol !== -1 ? String(row[nameCol - 1] || "").trim() : "";
          const email = emailCol !== -1 ? String(row[emailCol - 1] || "").trim() : "";

          try {
            const sourceFile = DriveApp.getFileById(driveId);

            // Cây thư mục: Root -> Squad/Unit -> [<Code>] <FullName>
            const squadName = AvatarCropperUtil.lookupSquadFromHera(email);
            const squadFolder = AvatarCropperUtil.getOrCreateFolder(rootFolder, squadName);
            const candidateFolderName = empCode ? `[${empCode}] ${fullName}` : fullName || `Dòng_${rowIndex}`;
            const targetFolder = AvatarCropperUtil.getOrCreateFolder(squadFolder, candidateFolderName);

            // Nhân bản file ảnh sang folder đích
            const safeName = String(fullName || "Candidate")
              .replace(/[\\/:*?"<>|]/g, "")
              .trim();
            const newFileName = empCode ? `[${empCode}] ${safeName}_1x1.jpg` : `${safeName}_1x1.jpg`;
            const copiedFile = sourceFile.makeCopy(newFileName, targetFolder);

            try {
              copiedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            } catch (shareErr) {}

            const copiedUrl = copiedFile.getUrl();

            // Ghi link vào cột "Ảnh 1:1"
            sheet.getRange(rowIndex, squareCol).setValue(copiedUrl);

            // Đồng bộ luôn trạng thái cột "Check 1:1"
            if (checkCol !== -1) {
              sheet.getRange(rowIndex, checkCol).setValue("1:1 (Đã crop)");
            }

            processedCount++;
          } catch (err) {
            Logger.log(`⚠️ Lỗi sao lưu ảnh dòng #${rowIndex}: ${err.message}`);
          }
        }
      }
    });

    return {
      ok: true,
      count: processedCount,
      message: `Đã tự động sao chép ${processedCount} ảnh đạt chuẩn 1:1 vào thư mục lưu trữ và cập nhật vào Sheet UC06.`,
    };
  }
}

// Global bridges cho google.script.run
function cropperApi_getSelectedRowContext() {
  return AvatarCropperService.getSelectedRowContext();
}
function cropperApi_saveCroppedAvatar(rowIndex, base64Data, fileName) {
  return AvatarCropperService.saveCroppedAvatar(rowIndex, base64Data, fileName);
}
function cropperApi_getStorageFolder() {
  return AvatarCropperService.getStorageFolderConfig();
}
function cropperApi_setStorageFolder(folderInput) {
  return AvatarCropperService.setStorageFolderConfig(folderInput);
}
function cropperApi_getAvatarFileName(rowContext, suffix) {
  return AvatarCropperUtil.buildAvatarFileName(rowContext, suffix);
}
