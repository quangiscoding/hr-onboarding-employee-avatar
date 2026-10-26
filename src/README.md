# UC06 - Employee Avatar Tools (Google Apps Script)

Dự án mở rộng Google Sheets (Container-bound Script) phục vụ chuẩn hóa, cắt chỉnh ảnh đại diện nhân sự (tỉ lệ 1:1) và đồng bộ Google Drive.

---

## 1. Cấu trúc thư mục (Directory Structure)

Dự án áp dụng kiến trúc **Feature-based Modular Structure**. Mọi mã nguồn nằm trong thư mục `src/`:

```text
.
├── .clasp.json               # Cấu hình clasp (scriptId, rootDir: ./src)
├── .claspignore             # Danh sách file loại trừ không đẩy lên Apps Script
├── package.json              # Khai báo dependencies (pnpm)
├── pnpm-lock.yaml            # File lock phiên bản gói
├── README.md                 # Tài liệu quy chuẩn dự án
└── src/
    ├── appsscript.json       # Manifest cấu hình OAuth scopes & runtime
    ├── main.js               # Controller điều phối: onOpen(), menu triggers
    ├── core/                 # (Tùy chọn) Hạ tầng dùng chung toàn dự án
    │   ├── config.js         # Cấu hình hằng số toàn cục
    │   └── utils.js          # Tiện ích chung hệ thống
    └── <feature-name>/       # Mỗi tính năng đóng gói trong 1 thư mục riêng
        ├── <feature>.config.js
        ├── <feature>.service.js
        ├── <feature>.utils.js
        └── <feature>-sidebar.html
```

---

## 2. Quy tắc lập trình bắt buộc (Engineering Guidelines)

### A. Phân tách Feature Folder (`src/<feature-name>`)

- Mỗi tính năng mới phải nằm trong một folder độc lập dưới `src/` (ví dụ: `src/avatar-cropper/`).
- Không đặt các file phục vụ riêng cho một tính năng ra ngoài thư mục gốc của `src/`.

### B. Quy tắc Service Pattern (`FeatureNameService`)

- Mọi nghiệp vụ (Business Logic, Drive API, Sheets API) phải được gom nhóm trong một class: `FeatureNameService`.
- Toàn bộ hàm thực thi được truy cập qua static methods hoặc instance của class này.
- **Quy ước:**

```javascript
class AvatarCropperService {
  /**
   * Lấy dữ liệu dòng đang chọn để truyền cho Sidebar
   */
  static getSelectedRowContext() {
    // Business logic...
  }

  /**
   * Nhận Base64, lưu Drive và gắn link Sheet
   */
  static saveCroppedAvatar(rowIndex, base64Data, fileName) {
    // Business logic...
  }
}
```

### C. Quy tắc Utility Pattern (`FeatureNameUtil`)

- Các hàm bổ trợ xử lý dữ liệu thuần (format chuỗi, regex, tính toán ma trận, xử lý Base64, v.v.) dành riêng cho tính năng phải được bọc trong class: `FeatureNameUtil`.
- Không tạo các hàm helper tự do ngoài global scope để chống xung đột tên hàm giữa các module.
- **Quy ước:**

```javascript
class AvatarCropperUtil {
  /**
   * Trích xuất Google Drive ID từ chuỗi URL
   */
  static extractDriveId(url) {
    if (!url) return null;
    const match = String(url).match(
      /(?:\/(?:d|folders)\/|[?&]id=)([-\w]{25,50})(?:[/?&#]|$)/,
    );
    return match ? match[1] : null;
  }

  /**
   * Chuẩn hóa text tiêu đề cột
   */
  static normalizeHeader(header) {
    return String(header || "")
      .trim()
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-z0-9]/g, "");
  }
}
```

### D. File `main.js` chỉ làm nhiệm vụ Controller

- `main.js` chỉ chứa các trigger hệ thống (`onOpen`, `onEdit`) và các hàm điều phối menu.
- Tuyệt đối không viết business logic, truy vấn dữ liệu trực tiếp trong `main.js`. Thay vào đó, gọi trực tiếp từ `FeatureNameService`:

```javascript
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🖼️ Avatar Tools")
    .addItem("1. Cắt ảnh nhân sự 1:1", "menuOpenAvatarCropper")
    .addToUi();
}

function menuOpenAvatarCropper() {
  AvatarCropperService.openCropperSidebar();
}
```

---

## 3. Quy trình phát triển (Local Workflow)

Dự án dùng **pnpm** và `@google/clasp`.

```bash
# Cài đặt dependencies
pnpm install

# Đẩy mã nguồn lên Google Apps Script
pnpm run push

# Bật chế độ tự động watch & push khi sửa code
pnpm run watch

# Mở Apps Script editor trên trình duyệt
pnpm run open
```
