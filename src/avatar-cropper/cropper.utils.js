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
}
