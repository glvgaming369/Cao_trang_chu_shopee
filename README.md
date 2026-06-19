# Shopee Product Scraper & Sheet Sync Tool

Công cụ Tampermonkey Userscript hỗ trợ cào thông tin sản phẩm trên Shopee VN theo bộ lọc tùy chỉnh, tự động bổ sung dữ liệu hoa hồng chi tiết từ bên thứ ba (Addlivetag API), và đồng bộ hóa tự động lên **Google Sheets** thông qua OAuth 2.0 hoặc xuất file **Excel (.xlsx)** dự phòng khi có sự cố.

---

## ✨ Tính năng nổi bật

1. **Bộ lọc sản phẩm thông minh**:
   - Tự động lọc các sản phẩm có gắn nhãn **Xtra** (hỗ trợ cả mẫu nhãn ảnh cũ và mã nhãn ảnh mới trên CDN Shopee).
   - Thiết lập giá sàn tối thiểu (`Price Min`) và lượt bán tối thiểu (`Sold Min`).
   - Sắp xếp kết quả theo các tiêu chí của Shopee: Liên quan, Bán chạy, Mới nhất, Phổ biến.
2. **Cào dữ liệu tự động & mượt mà**:
   - Cơ chế tự động cuộn trang (`smoothScrollToBottom`) xuống đáy trang để kích hoạt lazy-load toàn bộ sản phẩm trước khi phân tích.
   - Hỗ trợ chuyển trang đệ quy tự động cho đến khi hết sản phẩm đạt chuẩn hoặc người dùng nhấn dừng.
   - Cơ chế loại trùng bằng bộ nhớ đệm cache (`crawled_items_cache`) lưu trên Tampermonkey để tránh cào trùng sản phẩm ở các phiên chạy sau.
   - **Đồng bộ bộ nhớ đệm từ Google Sheets**: Tích hợp tính năng tải toàn bộ danh sách Mã SP ở cột A (bỏ qua tiêu đề) từ Google Sheets về trình duyệt, tự động gộp vào bộ nhớ đệm và loại bỏ trùng lặp.
3. **Bổ sung thông tin Hoa hồng (Commission)**:
   - Tự động gọi API `data.addlivetag.com` để bổ sung 3 trường thông tin: **Hoa hồng chung** (`commission`), **Hoa hồng người bán** (`sellerComFinal`), và **Hoa hồng Shopee** (`shopeeComFinal`).
   - Cơ chế **Batching & Rate Limit Control**: Chia luồng gửi API tối đa 5 sản phẩm song song và delay 1 giây giữa các đợt để đảm bảo không vượt quá ngưỡng rate limit **~300 requests/phút** của API.
4. **Đồng bộ hóa dữ liệu linh hoạt**:
   - Mặc định đồng bộ hóa dữ liệu trực tiếp lên **Google Sheets** qua Google Sheets API v4.
   - Xác thực an toàn bằng OAuth 2.0 (hỗ trợ nhập Authorization Code trực tiếp trên giao diện để sinh Refresh Token lưu trữ lâu dài).
   - Nút **Test Connect** kiểm tra quyền ghi trực tiếp trên tab trang tính chỉ định.
   - Cơ chế **Fallback tự động**: Nếu xảy ra lỗi đẩy dữ liệu lên Google Sheets (như hết hạn token, lỗi mạng), công cụ sẽ tự động xuất và tải về file **Excel (.xlsx)** dự phòng.

---

## 🛠️ Hướng dẫn cài đặt

### Bước 1: Cài đặt Extension Tampermonkey
- Tải và cài đặt tiện ích mở rộng [Tampermonkey](https://www.tampermonkey.net/) trên Chrome, Edge hoặc trình duyệt bạn sử dụng.

### Bước 2: Cài đặt Userscript
1. Mở dashboard của Tampermonkey, chọn tab **Utilities**.
2. Chọn **Create a new script**.
3. Copy toàn bộ mã nguồn từ file [shopee_scraper.user.js](shopee_scraper.user.js) dán vào phần soạn thảo.
4. Nhấn `Ctrl + S` (hoặc `Cmd + S`) để lưu lại.

---

## ⚙️ Hướng dẫn cấu hình Google Sheets OAuth 2.0

Để sử dụng tính năng đồng bộ hóa Google Sheets, bạn cần tự tạo thông tin Client ID:

1. Truy cập [Google Cloud Console](https://console.cloud.google.com/).
2. Tạo một dự án mới và kích hoạt **Google Sheets API**.
3. Vào mục **OAuth consent screen**, cấu hình ứng dụng ở chế độ **External** hoặc **Internal** và thêm scope `.../auth/spreadsheets`.
4. Vào mục **Credentials** -> **Create Credentials** -> **OAuth client ID** -> Chọn Application type là **Desktop app**.
5. Nhận **Client ID** và **Client Secret**.
6. Mở trang Shopee, tại Panel của Tool:
   - Điền link **Spreadsheet URL** và **Tên Sheet Tab** cần ghi dữ liệu.
   - Điền **Google Client ID** và **Google Client Secret** rồi nhấn **Lưu cấu hình**.
   - Bấm **Cấp quyền Google**, cấp quyền và sao chép mã **Authorization Code** được cấp từ Google.
   - Dán mã code vào ô nhập liệu xuất hiện trên Panel và bấm **Xác nhận**.
   - Thử nghiệm kết nối bằng cách bấm nút **Test Connect**.

---

## 🚀 Cách sử dụng

1. Truy cập trang tìm kiếm hoặc trang danh mục bất kỳ trên [Shopee VN](https://shopee.vn).
2. Panel điều khiển sẽ hiển thị ở góc trên bên phải màn hình.
3. Thiết lập các thông số lọc:
   - **Sort by**: Tiêu chí sắp xếp.
   - **Price Min (₫)**: Giá sản phẩm tối thiểu (ví dụ: 10,000đ).
   - **Sold Min**: Lượt bán tối thiểu (ví dụ: 10).
4. Chọn đích lưu trữ mong muốn: **Chỉ xuất Excel**, **Đẩy trực tiếp Google Sheets** hoặc **Cả hai**.
5. Nhấn **Start** để bắt đầu quá trình cào và đồng bộ tự động. Nhấn **Stop** nếu muốn dừng khẩn cấp.
6. Nếu muốn xóa lịch sử các mã sản phẩm đã cào để cào lại từ đầu, hãy nhấn nút **Xóa Cache**.
7. Nếu muốn đồng bộ các sản phẩm đã cào từ Google Sheets về trình duyệt (tránh cào trùng những sản phẩm đã có trên Sheets ở các lượt cào tiếp theo), hãy nhấn nút **Đồng bộ Cache**.

---

## 📊 Cấu trúc cột dữ liệu đầu ra

Dữ liệu xuất ra Google Sheets và Excel bao gồm 9 cột thông tin:

| Cột | Tên trường dữ liệu | Mô tả | Nguồn dữ liệu |
| :--- | :--- | :--- | :--- |
| **A** | Mã SP (Item ID) | ID định danh sản phẩm trên Shopee | Shopee HTML |
| **B** | Mã Shop (Shop ID) | ID định danh cửa hàng | Shopee HTML |
| **C** | Tên sản phẩm | Tên đầy đủ của sản phẩm | Shopee HTML |
| **D** | Giá (đ) | Giá bán khuyến mãi hiện tại | Shopee HTML |
| **E** | Lượt bán | Số lượng sản phẩm đã bán | Shopee HTML |
| **F** | Đường dẫn sản phẩm | Link URL trực tiếp đến sản phẩm | Shopee HTML |
| **G** | Hoa hồng (đ) | Tổng số tiền hoa hồng của sản phẩm | Addlivetag API |
| **H** | Hoa hồng người bán (đ) | Tiền hoa hồng được trả bởi người bán | Addlivetag API |
| **I** | Hoa hồng Shopee (đ) | Tiền hoa hồng được tài trợ bởi Shopee | Addlivetag API |

---

## 🔒 Rate Limit & Performance Notes

- API Addlivetag giới hạn khoảng **~300 requests/phút**.
- Tool tự động điều phối hàng đợi yêu cầu (`batchSize: 5`, `delay: 1000ms`) để vừa bảo đảm tối ưu hóa thời gian xử lý dữ liệu, vừa tuân thủ chặt chẽ định mức của API nhằm tránh bị chặn IP/Device.
