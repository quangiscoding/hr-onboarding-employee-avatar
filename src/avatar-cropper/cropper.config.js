/** ==========================================
 * CROPPER-CONFIG.JS - CẤU HÌNH NỘI BỘ CỦA AVATAR CROPPER
 * ========================================== */
const CROPPER_CONFIG = {
  // Folder lưu ảnh (nếu để trống sẽ lưu vào My Drive gốc)
  STORAGE_FOLDER_ID: "",

  // Chuẩn kích thước ảnh vuông 1:1
  OUTPUT_IMAGE: {
    WIDTH: 800,
    HEIGHT: 800,
    MIME_TYPE: "image/jpeg",
    QUALITY: 0.92,
  },

  // Bản đồ tên cột trên sheet UC06 mà feature này thao tác
  COLUMNS: {
    EMPLOYEE_NAME: "Employee Name", // Cột A
    EMAIL: "Email Address", // Cột B (hỗ trợ đối soát email nếu cần)
    PICTURE: "Picture", // Cột F
    CHECK_SQUARE: "Check 1:1", // Cột G
    SQUARE_PICTURE: "Ảnh 1:1", // Cột H
  },
};
