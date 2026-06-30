# Shopee PDP → VideoAI (Chrome Extension)

Bản Chrome Extension (Manifest V3) của tool cào dữ liệu sản phẩm Shopee qua API `get_pc` rồi đẩy lên **VideoAI**. Chạy độc lập với userscript Tampermonkey (`../shopee_pdp_to_videoai.user.js`).

## Tính năng (v1.0)
- Dán hàng loạt link sản phẩm Shopee, tự **loại link trùng**.
- Mở **tab ngầm**, bắt API `get_pc` (event-driven, không poll) rồi đẩy VideoAI theo **lô ≤200**, **chu kỳ đẩy mỗi 30 SP**.
- **Cache itemID đã xử lý** (đẩy OK hoặc bị loại) — không làm lại; nút **Xóa** cache.
- **Đa tab song song** 1/2/3.
- **Resume** khi service worker bị ngủ/khởi động lại (nhờ session + cache).
- **Báo Telegram** khi gặp captcha (token / chat id / tên tab + cooldown 60s).

## Kiến trúc
| File | Vai trò |
|---|---|
| `background.js` | Service worker: state, mở tab, gọi VideoAI/Telegram, cache, resume |
| `content/pageHook.js` | MAIN world: hook `fetch`/XHR bắt `get_pc` → `postMessage` |
| `content/bridge.js` | Isolated: chuyển dữ liệu về background; phát hiện captcha + banner |
| `popup/` | Giao diện cấu hình + Start/Stop + log |
| `lib/shared.js` | Helper thuần + mapper dùng chung |

## Cài đặt (load unpacked)
1. Mở `chrome://extensions`.
2. Bật **Developer mode** (góc trên phải).
3. Bấm **Load unpacked** → chọn thư mục:
   - `shopee-pdp-extension` (bản source, để phát triển), HOẶC
   - `shopee-pdp-extension/dist` (bản **minify** để phát hành — xem mục Build).
4. Ghim icon extension, bấm để mở popup.

> Yêu cầu Chrome 111+ (dùng content script `world: "MAIN"`).

## Build bản minify (bảo vệ mã)
Source giữ dễ đọc; bản phân phối được **minify** (gộp `lib/shared.js`, đổi tên biến, bỏ comment) ra thư mục `dist/`:

```bash
cd shopee-pdp-extension
npm install      # cài esbuild (1 lần)
npm run build    # sinh dist/
```

Sau đó **Load unpacked** thư mục `dist/` và phát hành thư mục đó.

> Lưu ý: minify chỉ **làm khó đọc**, không phải bảo mật tuyệt đối — code client luôn có thể dịch ngược. API key vẫn nên để người dùng tự nhập (không hardcode). `dist/` và `node_modules/` không commit vào git (đã `.gitignore`).

## Sử dụng
1. Mở popup → nhập **VideoAI API Key** (key **production** cho host `videoai-api.devappnow.com`).
2. Dán link (mỗi dòng 1), chọn số tab song song.
3. (Tuỳ chọn) mở mục **Telegram**, nhập token/chat id/tên tab → **Test Telegram**.
4. Bấm **Bắt đầu**. Có thể đóng popup — background vẫn chạy; mở lại để xem tiến độ/log.

## Lưu ý
- Endpoint mặc định: `https://videoai-api.devappnow.com/api/products/shopee-cache/batch`. Host này **dùng API key khác** với bản dev — hãy nhập đúng key production (key dev sẽ trả `401`).
- Để cào được, phải **đăng nhập Shopee** trong trình duyệt.
- Phần "lấy link từ jobs server" sẽ bổ sung ở bản sau.
