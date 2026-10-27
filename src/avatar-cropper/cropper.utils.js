/**
 * src/avatar-cropper/cropper.utils.js
 */

class AvatarCropperUtil {
  /**
   * Trích xuất Google Drive File/Folder ID từ URL hoặc chuỗi ID thuần.
   * Áp dụng kiểm tra biên kết thúc [/?&#] và giới hạn độ dài 25-50 ký tự[cite: 5].
   * @param {string} url - Link Google Drive hoặc chuỗi ID
   * @returns {string|null} ID hợp lệ hoặc null
   */
  static extractDriveId(url) {
    if (!url) return null;
    const str = String(url).trim();

    // 1. Nhận diện cấu trúc URL: sau /d/ hoặc /folders/ hoặc ?id=
    const match = str.match(
      /(?:\/(?:d|folders)\/|[?&]id=)([-\w]{25,50})(?:[/?&#]|$)/,
    );
    if (match) return match[1];

    // 2. Fallback nếu dán thẳng chuỗi ID thô
    if (/^[-\w]{25,50}$/.test(str)) {
      return str;
    }

    return null;
  }

  /**
   * Tạo tên file avatar vuông 1:1, giữ nguyên dấu tiếng Việt và loại bỏ ký tự cấm của OS
   * @param {Object} rowContext - Object thông tin dòng hiện tại { fullName, rowIndex }
   * @param {string} suffix - Hậu tố ('avatar_1x1' hoặc '1x1')
   * @returns {string} Tên file hoàn chỉnh (VD: "Hoàng Thị Thùy Trang_avatar_1x1.jpg")
   */
  static buildAvatarFileName(rowContext, suffix = "avatar_1x1") {
    const rowIndex = rowContext ? rowContext.rowIndex : null;
    const rawFullName =
      rowContext && rowContext.fullName
        ? String(rowContext.fullName).trim()
        : "";

    // Giữ nguyên dấu tiếng Việt, loại bỏ ký tự cấm: \ / : * ? " < > | và chuẩn hóa khoảng trắng
    const cleanName = rawFullName
      ? rawFullName.replace(/[\\/:*?"<>|]/g, "").replace(/\s+/g, " ")
      : `candidate_${rowIndex || "new"}`;

    return `${cleanName}_${suffix}.jpg`;
  }

  /**
   * Chuẩn hóa tiêu đề cột (xóa dấu tiếng Việt, ký tự đặc biệt, khoảng trắng thừa)[cite: 6]
   * @param {string} header
   * @returns {string}
   */
  static normalizeHeader(header) {
    return String(header || "")
      .trim()
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-z0-9]/g, "");
  }

  /**
   * Chuyển đổi số thứ tự cột dạng 1-indexed thành ký tự A1 chuẩn (1 -> "A", 27 -> "AA")[cite: 1]
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
}
