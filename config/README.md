# Config Directory

Thư mục này dùng để lưu trữ các file cấu hình và file credentials của công cụ.

## Hướng dẫn:
1. Sao chép file JSON xác thực Google OAuth 2.0 (tải từ Google Cloud Console) vào thư mục này.
2. Đặt tên file là `client_secret.json` hoặc sử dụng tên mặc định và chọn đường dẫn từ giao diện công cụ.
3. Tool cũng sẽ lưu trữ file token đã xác thực (`token.json`) và file cấu hình ứng dụng (`config.json`) tại thư mục này để tránh phải thiết lập lại ở những lần chạy sau.
