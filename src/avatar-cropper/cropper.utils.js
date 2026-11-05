/**
 * src/avatar-cropper/cropper.utils.js
 */
class AvatarCropperUtil {
  /**
   * Chuẩn hóa tiêu đề cột khớp 100% với hàm normalizeHeaderKey của hệ thống Hera
   * @param {string} header
   * @returns {string}
   */
  static normalizeHeader(header) {
    if (!header) return "";
    return String(header)
      .toLowerCase()
      .replace(/[^a-z0-9]/g, "");
  }

  /**
   * Trích xuất Google Drive File/Folder ID từ URL hoặc chuỗi ID thuần
   * @param {string} url - Link Google Drive hoặc chuỗi ID
   * @returns {string|null} ID hợp lệ hoặc null
   */
  static extractDriveId(url) {
    if (!url) return null;
    const str = String(url).trim();
    const match = str.match(/(?:\/(?:d|folders)\/|[?&]id=)([-\w]{25,50})(?:[/?&#]|$)/);
    if (match) return match[1];
    if (/^[-\w]{25,50}$/.test(str)) {
      return str;
    }
    return null;
  }

  /**
   * Tạo tên file avatar vuông 1:1, giữ nguyên dấu tiếng Việt và loại bỏ ký tự cấm của OS
   * @param {Object} rowContext - Object thông tin dòng hiện tại { fullName, rowIndex }
   * @param {string} suffix - Hậu tố ('avatar_1x1' hoặc '1x1')
   * @returns {string} Tên file hoàn chỉnh
   */
  static buildAvatarFileName(rowContext, suffix = "avatar_1x1") {
    const rowIndex = rowContext ? rowContext.rowIndex : null;
    const rawFullName = rowContext && rowContext.fullName ? String(rowContext.fullName).trim() : "";
    const cleanName = rawFullName ? rawFullName.replace(/[\\/:*?"<>|]/g, "").replace(/\s+/g, " ") : `candidate_${rowIndex || "new"}`;
    return `${cleanName}_${suffix}.jpg`;
  }

  /**
   * Chuyển đổi số thứ tự cột dạng 1-indexed thành ký tự A1 chuẩn (1 -> "A", 27 -> "AA")
   * @param {number} colIndex
   * @returns {string}
   */
  static columnToLetter(colIndex) {
    let temp,
      letter = "";
    while (colIndex > 0) {
      temp = (colIndex - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      colIndex = Math.floor((colIndex - temp) / 26);
    }
    return letter;
  }

  /**
   * Lấy hoặc tạo mới một thư mục con bên trong thư mục cha (Get-or-Create)
   * @param {GoogleAppsScript.Drive.Folder} parentFolder
   * @param {string} folderName
   * @returns {GoogleAppsScript.Drive.Folder}
   */
  static getOrCreateFolder(parentFolder, folderName) {
    const safeName = String(folderName || "Others").trim();
    const folders = parentFolder.getFoldersByName(safeName);
    if (folders.hasNext()) {
      return folders.next();
    }
    return parentFolder.createFolder(safeName);
  }

  /**
   * Tra cứu tên Squad/Unit từ file Hera Onboarding List dựa theo Email cá nhân
   * @param {string} personalEmail
   * @returns {string} Tên Squad/Unit hoặc "Others" nếu không tìm thấy
   */
  static lookupSquadFromHera(personalEmail) {
    if (!personalEmail) return "Others";
    try {
      const heraId = CROPPER_CONFIG.HERA_SPREADSHEET_ID;
      if (!heraId) return "Others";
      const ss = SpreadsheetApp.openById(heraId);

      // 1. Trỏ chính xác vào tab theo GID 335034815 hoặc tên tab "2026"
      const sheet = ss.getSheets().find((s) => String(s.getSheetId()) === "335034815") || ss.getSheetByName("2026") || ss.getSheets()[0];
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow < 3) return "Others";

      // 2. Đọc dòng 1 để định vị vị trí các cột
      const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      let emailColIdx = -1;
      let squadColIdx = -1;
      const targetSquadKey = this.normalizeHeader((CROPPER_CONFIG.HERA_LOOKUP_COLS && CROPPER_CONFIG.HERA_LOOKUP_COLS.SQUAD) || "Squad/ Unit");
      const targetEmailKey = this.normalizeHeader((CROPPER_CONFIG.HERA_LOOKUP_COLS && CROPPER_CONFIG.HERA_LOOKUP_COLS.PERSONAL_EMAIL) || "Personal Email");

      for (let c = 0; c < lastCol; c++) {
        const val = this.normalizeHeader(headerRow[c]);
        if (val === targetEmailKey) emailColIdx = c;
        if (val === targetSquadKey) squadColIdx = c;
      }

      // Fallback theo index thực tế từ log: Cột I (index 8) và Cột AD (index 29)
      if (squadColIdx === -1) squadColIdx = 8;
      if (emailColIdx === -1) emailColIdx = 29;

      // 3. Đọc dữ liệu ứng viên từ dòng 3 trở đi
      const numRows = lastRow - 2;
      const data = sheet.getRange(3, 1, numRows, lastCol).getValues();
      const targetEmail = personalEmail.toLowerCase().trim();

      for (let i = 0; i < data.length; i++) {
        const rowEmail = String(data[i][emailColIdx] || "")
          .toLowerCase()
          .trim();
        if (rowEmail === targetEmail) {
          const squad = String(data[i][squadColIdx] || "").trim();
          return squad || "Others";
        }
      }
    } catch (err) {
      Logger.log("⚠️ Lỗi tra cứu Squad từ Hera: " + err.message);
    }
    return "Others";
  }

  /**
   * Trích xuất width & height của file ảnh từ Google Drive ID bằng binary header
   * @param {string} fileId - ID của file trên Google Drive
   * @returns {{ width: number, height: number, isSquare: boolean, ratio: number } | null}
   */
  static getImageDimensions(fileId) {
    try {
      const file = DriveApp.getFileById(fileId);
      const blob = file.getBlob();
      const bytes = blob.getBytes();
      const contentType = blob.getContentType().toLowerCase();
      let width = 0;
      let height = 0;

      // 1. Định dạng PNG
      if (contentType.includes("png")) {
        if (bytes.length >= 24) {
          width = ((bytes[16] & 0xff) << 24) | ((bytes[17] & 0xff) << 16) | ((bytes[18] & 0xff) << 8) | (bytes[19] & 0xff);
          height = ((bytes[20] & 0xff) << 24) | ((bytes[21] & 0xff) << 16) | ((bytes[22] & 0xff) << 8) | (bytes[23] & 0xff);
        }
      }
      // 2. Định dạng JPEG / JPG
      else if (contentType.includes("jpeg") || contentType.includes("jpg")) {
        let offset = 2;
        while (offset < bytes.length) {
          if ((bytes[offset] & 0xff) === 0xff) {
            const marker = bytes[offset + 1] & 0xff;
            // SOF0, SOF1, SOF2 markers chứa kích thước ảnh
            if ((marker >= 0xc0 && marker <= 0xc3) || (marker >= 0xc5 && marker <= 0xc7) || (marker >= 0xc9 && marker <= 0xcb) || (marker >= 0xcd && marker <= 0xcf)) {
              height = ((bytes[offset + 5] & 0xff) << 8) | (bytes[offset + 6] & 0xff);
              width = ((bytes[offset + 7] & 0xff) << 8) | (bytes[offset + 8] & 0xff);
              break;
            }
            const length = ((bytes[offset + 2] & 0xff) << 8) | (bytes[offset + 3] & 0xff);
            offset += 2 + length;
          } else {
            offset++;
          }
        }
      }
      // 3. Định dạng GIF
      else if (contentType.includes("gif")) {
        if (bytes.length >= 10) {
          width = (bytes[6] & 0xff) | ((bytes[7] & 0xff) << 8);
          height = (bytes[8] & 0xff) | ((bytes[9] & 0xff) << 8);
        }
      }
      // 4. Định dạng WEBP
      else if (contentType.includes("webp")) {
        if (bytes.length >= 30) {
          // Simple VP8
          if (bytes[12] === 0x56 && bytes[13] === 0x50 && bytes[14] === 0x38 && bytes[15] === 0x20) {
            width = ((bytes[26] & 0xff) | ((bytes[27] & 0xff) << 8)) & 0x3fff;
            height = ((bytes[28] & 0xff) | ((bytes[29] & 0xff) << 8)) & 0x3fff;
          }
          // Extended VP8X
          else if (bytes[12] === 0x56 && bytes[13] === 0x50 && bytes[14] === 0x38 && bytes[15] === 0x58) {
            width = 1 + ((bytes[24] & 0xff) | ((bytes[25] & 0xff) << 8) | ((bytes[26] & 0xff) << 16));
            height = 1 + ((bytes[27] & 0xff) | ((bytes[28] & 0xff) << 8) | ((bytes[29] & 0xff) << 16));
          }
        }
      }

      if (width > 0 && height > 0) {
        // 1. Tính toán trên số thực chính xác (không làm tròn trước)
        const rawRatio = width / height;
        const pixelDiff = Math.abs(width - height);
        const ratioDiff = Math.abs(rawRatio - 1.0);

        // 2. Tiêu chí: Lệch <= 4 pixel HOẶC lệch tỉ lệ <= 2% (0.02)
        const isSquare = pixelDiff <= 4 || ratioDiff <= 0.02;

        // 3. Làm tròn ratio để hiển thị/ghi log cho gọn
        const displayRatio = Number(rawRatio.toFixed(2));

        return {
          width: width,
          height: height,
          ratio: displayRatio,
          isSquare: isSquare,
        };
      }
    } catch (err) {
      Logger.log("❌ Lỗi đọc file ID: " + fileId + " - Chi tiết: " + err.message);
    }
    return null;
  }
}
