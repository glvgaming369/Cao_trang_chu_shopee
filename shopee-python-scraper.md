# Shopee Scraper & Google Sheets Sync Python Tool - Implementation Plan (Multi-Worker Upgrade)

Bản kế hoạch chi tiết xây dựng công cụ Python cào dữ liệu Shopee kết hợp dữ liệu hoa hồng Addlivetag và đồng bộ lên Google Sheets, tích hợp tính năng đa luồng song song (Multi-workers) dựa trên danh sách Proxy.

## Overview
- **Project Type:** WEB (Python Desktop Tool)
- **Primary Agent:** `backend-specialist` (cho cơ chế Queue/Threading và điều phối Proxy) & `frontend-specialist` (cho GUI CustomTkinter)

## Success Criteria
1. Nâng cấp ô nhập Proxy trên UI thành Textbox nhập nhiều dòng (mỗi dòng một proxy).
2. Tự động nhận diện số lượng proxy đầu vào để tạo ra số lượng Worker tương ứng chạy song song.
3. Sử dụng `queue.Queue` để phân phối sản phẩm cần cào cho các worker an toàn.
4. Sử dụng `threading.Lock` để đồng bộ hóa việc ghi nhận kết quả và log giữa các luồng.
5. Mỗi worker giữ khoảng trễ tối thiểu 1.0 giây để bảo đảm không bị rate limit trên từng proxy.
6. Sau khi tất cả worker hoàn thành, dữ liệu sẽ được đẩy lên Google Sheets một lần duy nhất (hoặc xuất Excel dự phòng nếu lỗi).

## File Structure
```plaintext
Cao_trang_chu_shopee/
├── config/
│   ├── README.md               # Hướng dẫn cấu hình credentials
│   ├── config.json             # Cấu hình UI đã lưu (Tự động sinh)
│   └── token.json              # Google OAuth token (Tự động sinh sau khi auth)
├── requirements.txt            # Thư viện phụ thuộc
├── shopee_scraper_core.py      # Core logic cào dữ liệu, API, Google Sheets
├── shopee_scraper_gui.py       # Giao diện CustomTkinter chính (Chứa logic multi-worker)
└── shopee-python-scraper.md    # File kế hoạch này
```

## Task Breakdown

### Task 1: Nâng cấp Giao diện nhập Proxy
- **Agent:** `frontend-specialist`
- **INPUT:** Thay thế `CTkEntry` nhập proxy bằng `CTkTextbox` hỗ trợ nhiều dòng.
- **OUTPUT:** File `shopee_scraper_gui.py` được thay thế widget proxy và cập nhật các hàm load/save config.
- **VERIFY:** Khởi động GUI, giao diện hiển thị hộp textbox nhập proxy, lưu cấu hình thành công khi nhập nhiều proxy.

### Task 2: Triển khai luồng nền Multi-Worker
- **Agent:** `backend-specialist`
- **INPUT:** Danh sách sản phẩm, danh sách proxy, cờ dừng (`self.is_running`).
- **OUTPUT:** Triển khai cơ chế Multi-worker bằng `threading.Thread` và `queue.Queue` trong hàm `scraper_thread`:
  - Phân tích và làm sạch danh sách proxy.
  - Đưa tất cả ID sản phẩm vào Queue.
  - Tạo $N$ Worker, mỗi worker có tên định danh, được gán 1 proxy, và chạy độc lập trong khi Queue còn phần tử và cờ `is_running` là True.
  - Các worker cập nhật kết quả vào mảng dùng chung và ghi log an toàn qua thread Lock.
- **VERIFY:** Chạy kiểm thử đa luồng và quan sát log in ra đồng thời từ các worker khác nhau.

### Task 3: Tích hợp hoàn thiện và kiểm thử
- **Agent:** `test-engineer`
- **INPUT:** Hệ thống tích hợp hoàn chỉnh.
- **OUTPUT:** Bản chạy thử nghiệm không lỗi.
- **VERIFY:** 
  - Chạy thử với danh sách 10 link Shopee và 3 proxy.
  - Xác nhận tốc độ xử lý nhanh hơn đáng kể.
  - Kiểm tra xem dữ liệu ghi lên Google Sheets có đầy đủ và không bị trùng lặp hay không.
  - Bấm nút Stop giữa chừng để xác nhận tất cả worker dừng lại lập tức.
