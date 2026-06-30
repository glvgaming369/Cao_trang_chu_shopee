# Shopee Unified Market Scraper (VN + PH)

Tampermonkey Userscript cào thông tin sản phẩm trên **Shopee VN & PH** (tự nhận diện theo domain) theo **danh mục** hoặc **từ khóa**. Công cụ mở tab ngầm để bắt gói tin API affiliate nhằm bổ sung dữ liệu hoa hồng chi tiết mà không thao tác lộ liễu trên tài khoản, sau đó đồng bộ lên **Google Sheets**, **VideoAI API**, hoặc xuất **Excel (.xlsx)** dự phòng.

> Toàn bộ công cụ nằm gọn trong một file duy nhất: [`shopee_scraper_multi.user.js`](shopee_scraper_multi.user.js). Không phụ thuộc file ngoài (danh mục & blacklist đã nhúng sẵn trong script).

---

## ✨ Tính năng nổi bật

1. **Đa thị trường**: Tự nhận diện `shopee.vn` / `shopee.ph` (và cổng `affiliate.shopee.*`), áp dụng tiền tệ, danh mục và bộ lọc mặc định tương ứng.
2. **Ba chế độ cào**:
   - **Theo danh mục**: chạy trên trang tìm kiếm/danh mục Shopee, tự cuộn lazy-load và chuyển trang đệ quy.
   - **Theo từ khóa**: chạy trên cổng affiliate, lần lượt xử lý danh sách từ khóa.
   - **Theo link**: dán hàng loạt link sản phẩm (mỗi dòng 1 link, dạng `.../product/{shopId}/{itemId}`); tool cào đúng các sản phẩm đó (không lọc hoa hồng, luôn cào lại).
3. **Bắt gói tin API affiliate**: Mở tab ngầm tới trang chi tiết ưu đãi, hook `fetch`/`XHR` để lấy `/api/v3/offer/product`, từ đó trích **hoa hồng người bán / Shopee / tổng**, tồn kho, đánh giá, tên shop... Hỗ trợ lấy thêm tối đa **5 sản phẩm tương tự (L2)** từ mỗi sản phẩm gốc (L1).
4. **Chống ban & captcha**: Phát hiện trang captcha/verify, **tự động tạm dừng** và chờ giải, có banner cảnh báo khi tab affiliate dính captcha.
5. **Bộ nhớ đệm (cache)**: Lưu danh sách mã đã cào/đã bỏ qua trong Tampermonkey để tránh trùng lặp giữa các phiên. Hỗ trợ **đồng bộ cache từ Google Sheets** (cột A).
6. **Blacklist từ khóa**: Lọc bỏ các sản phẩm nhạy cảm (mặc định bật cho VN), có thể chỉnh sửa trên giao diện.
7. **Đồng bộ linh hoạt**:
   - **Google Sheets** qua OAuth 2.0 (Sheets API v4).
   - **VideoAI API** — đẩy sản phẩm lên `shopee-cache/batch` (thủ công hoặc tự động).
   - **Excel (.xlsx)** dự phòng khi đồng bộ lỗi.

---

## 🛠️ Cài đặt

1. Cài tiện ích [Tampermonkey](https://www.tampermonkey.net/) trên Chrome/Edge.
2. Mở dashboard Tampermonkey → **Create a new script**.
3. Dán toàn bộ nội dung [`shopee_scraper_multi.user.js`](shopee_scraper_multi.user.js) vào và lưu (`Ctrl + S`).
4. Mở `shopee.vn` hoặc `shopee.ph` — panel điều khiển hiện ở góc trên bên phải.

---

## 🚀 Sử dụng

1. Chọn **chế độ cào** (danh mục / từ khóa) và thiết lập bộ lọc (giá tối thiểu, lượt bán tối thiểu, hoa hồng người bán tối thiểu).
2. Nhấn **Start Crawl**. Dữ liệu được đẩy theo từng trang/từng từ khóa. Nhấn **Stop** để dừng.
3. **Xóa Cache** để cào lại từ đầu; **Đồng bộ Cache** để nạp mã đã có trên Google Sheets vào bộ nhớ đệm; **Xuất Excel** để tải file.

### Cấu hình Google Sheets (OAuth 2.0)
1. Tạo project trên [Google Cloud Console](https://console.cloud.google.com/), bật **Google Sheets API**, thêm scope `.../auth/spreadsheets`.
2. Tạo **OAuth client ID** loại **Desktop app** để lấy **Client ID** + **Client Secret**.
3. Trong panel: điền Spreadsheet URL, tên Tab, Client ID/Secret → **Lưu cấu hình** → **Kết nối Google** → dán **Authorization Code** → **Xác nhận**. Dùng **Test Sheet** để kiểm tra quyền ghi.

### Cấu hình VideoAI API
1. Mở khối **🚀 Cấu hình VideoAI API** trong panel.
2. Dán **API Key** (dạng `api_xxx`), giữ nguyên Endpoint mặc định, **Lưu cấu hình** rồi **Test kết nối**.
3. Bật **Tự động đẩy** để đồng bộ realtime, hoặc bấm **Đẩy lên VideoAI** để gửi thủ công.

Yêu cầu của API: mỗi sản phẩm cần `url, title, price, rating, soldCount, stock, shopName, images (≥3), reviewCount, ratingCount`. Tool tự **lọc bỏ sản phẩm < 3 ảnh** trước khi gửi, chia lô tối đa **200 sản phẩm/request**, và cùng `url` + cùng key sẽ **ghi đè**.

---

## 📊 Cột dữ liệu đầu ra

Dữ liệu xuất ra Google Sheets / Excel gồm 18 cột: Mã SP, Mã Shop, Tên sản phẩm, Giá, Lượt bán 30 ngày, Đường dẫn, Hoa hồng, Hoa hồng người bán, Hoa hồng Shopee, Nhóm sản phẩm, Danh sách ảnh, Category, Tồn kho, Đánh giá sao, Lượt đánh giá, Ngày đăng bán, Phân loại link (L1/L2), Link gộp.

---

## 🔒 Ghi chú

- Công cụ phục vụ mục đích nghiên cứu/tổng hợp dữ liệu công khai; hãy tuân thủ điều khoản của Shopee và pháp luật hiện hành.
- API Key và thông tin OAuth được lưu cục bộ trong storage của Tampermonkey, không nhúng cứng trong mã nguồn.
