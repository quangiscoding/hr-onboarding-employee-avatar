/** ==========================================
 * CROPPER-CONFIG.JS - CẤU HÌNH NỘI BỘ CỦA AVATAR CROPPER
 * ========================================== */
const CROPPER_CONFIG = {
  // Folder lưu ảnh gốc (nếu để trống sẽ lưu vào My Drive gốc)
  STORAGE_FOLDER_ID: "",

  // File Hera On-boarding List độc lập (dùng để tra cứu Squad/Unit qua Personal Email)
  HERA_SPREADSHEET_ID: "1rnowvvOuOJ_5i9zDrz_1-TigNq_EZMY02nnOGbGQgTE",

  // Chuẩn kích thước ảnh vuông 1:1
  OUTPUT_IMAGE: {
    WIDTH: 800,
    HEIGHT: 800,
    MIME_TYPE: "image/jpeg",
    QUALITY: 0.92,
  },

  // Bản đồ tên cột trên sheet UC06 hiện tại
  COLUMNS: {
    CODE: "Code", // Cột A: Mã nhân sự
    EMPLOYEE_NAME: "Employee Name", // Cột B: Họ và tên
    EMAIL: "Email Address", // Cột C: Email đối soát sang Hera
    PICTURE: "Picture", // Cột G: Ảnh gốc Drive
    CHECK_SQUARE: "Check 1:1", // Cột H: Trạng thái 1:1
    SQUARE_PICTURE: "Ảnh 1:1", // Cột I: Link ảnh sau khi crop
  },

  // Tên cột trên sheet Hera dùng để tra cứu
  HERA_LOOKUP_COLS: {
    SQUAD: "Squad/ Unit",
    PERSONAL_EMAIL: "Personal Email",
  },
};
