// ==UserScript==
// @name         Shopee Product Scraper
// @namespace    http://tampermonkey.net/
// @version      1.1
// @description  Cào dữ liệu sản phẩm Shopee theo bộ lọc tùy chỉnh và xuất file Excel / Push Google Sheets
// @author       Antigravity
// @match        https://shopee.vn/*
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_registerMenuCommand
// @grant        GM_xmlhttpRequest
// @require      https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js
// ==/UserScript==

(function () {
    'use strict';

    // Thư viện XLSX đã được load qua @require
    // Khởi tạo bộ nhớ cache cho các item đã cào
    let crawledCache = GM_getValue('crawled_items_cache', []);
    let crawledData = []; // Dữ liệu của lượt cào hiện tại
    let isRunning = false;
    let autoScrollInterval = null;

    // Load các cấu hình Google Sheets đã lưu
    let sheetConfig = GM_getValue('sheet_config', {
        sheetUrl: '',
        tabName: 'Sheet1',
        clientId: '',
        clientSecret: ''
    });

    // Tạo giao diện Panel điều khiển nổi (Floating Panel)
    const panel = document.createElement('div');
    panel.id = 'shopee-scraper-panel';
    panel.style.position = 'fixed';
    panel.style.top = '80px';
    panel.style.right = '20px';
    panel.style.zIndex = '999999';
    panel.style.backgroundColor = '#1a1a1a';
    panel.style.color = '#ffffff';
    panel.style.padding = '15px';
    panel.style.borderRadius = '8px';
    panel.style.boxShadow = '0 4px 15px rgba(0,0,0,0.5)';
    panel.style.fontFamily = 'Arial, sans-serif';
    panel.style.fontSize = '13px';
    panel.style.width = '280px';
    panel.style.border = '1px solid #ee4d2d';

    // Cập nhật số lượng trong cache hiển thị lên giao diện
    function updateCacheCount() {
        const cacheEl = document.getElementById('scraper-cache-count');
        if (cacheEl) {
            cacheEl.innerText = crawledCache ? crawledCache.length : 0;
        }
    }

    // Ghi log ra Panel
    function addLog(msg, type = 'info') {
        const logBox = document.getElementById('scraper-logs');
        if (!logBox) {
            console.log(`[Shopee Scraper] ${msg}`);
            return;
        }
        const time = new Date().toLocaleTimeString();
        let color = '#ccc';
        if (type === 'success') color = '#4caf50';
        else if (type === 'warn') color = '#ffeb3b';
        else if (type === 'error') color = '#f44336';
        else if (type === 'info') color = '#2196f3';

        const line = document.createElement('div');
        line.style.marginBottom = '4px';
        line.style.color = color;
        line.innerHTML = `<span style="color: #777;">[${time}]</span> ${msg}`;
        logBox.appendChild(line);
        logBox.scrollTop = logBox.scrollHeight;
    }

    function initPanel() {
        if (document.getElementById('shopee-scraper-panel')) return;

        panel.innerHTML = `
            <div style="font-weight: bold; font-size: 15px; margin-bottom: 12px; color: #ee4d2d; border-bottom: 1px solid #333; padding-bottom: 5px; display: flex; justify-content: space-between; align-items: center;">
                <span>Shopee Scraper Pro v1.1</span>
                <span id="scraper-status" style="font-size: 10px; padding: 2px 6px; background-color: #333; border-radius: 4px; color: #aaa;">Idle</span>
            </div>
            
            <!-- Bộ lọc cơ bản -->
            <div style="margin-bottom: 10px;">
                <label style="display: block; margin-bottom: 4px; font-weight: bold;">Sort by:</label>
                <select id="scraper-sort" style="width: 100%; padding: 6px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white;">
                    <option value="relevance">Relevance (Liên quan)</option>
                    <option value="sales">Top Sales (Bán chạy)</option>
                    <option value="ctime">Latest (Mới nhất)</option>
                    <option value="pop">Popular (Phổ biến)</option>
                </select>
            </div>
            <div style="margin-bottom: 10px; display: flex; gap: 8px;">
                <div style="flex: 1;">
                    <label style="display: block; margin-bottom: 4px; font-weight: bold;">Price Min (₫):</label>
                    <input type="number" id="scraper-price-min" value="80000" style="width: 90%; padding: 6px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white;">
                </div>
                <div style="flex: 1;">
                    <label style="display: block; margin-bottom: 4px; font-weight: bold;">Sold Min:</label>
                    <input type="number" id="scraper-sold-min" value="10" style="width: 90%; padding: 6px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white;">
                </div>
            </div>
            
            <!-- Accordion cấu hình Google Sheets -->
            <div style="margin-bottom: 10px; border: 1px solid #333; border-radius: 4px;">
                <div id="scraper-sheet-toggle" style="background-color: #2b2b2b; padding: 6px 10px; font-weight: bold; cursor: pointer; display: flex; justify-content: space-between; align-items: center; border-radius: 4px 4px 0 0;">
                    <span>⚙️ Cấu hình Google Sheets</span>
                    <span id="scraper-toggle-arrow">▲</span>
                </div>
                <div id="scraper-sheet-fields" style="padding: 10px; display: block; background-color: #1f1f1f;">
                    <div style="margin-bottom: 8px;">
                        <label style="display: block; margin-bottom: 3px; font-size: 11px;">Chọn đích lưu trữ:</label>
                        <select id="scraper-save-target" style="width: 100%; padding: 4px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white; font-size: 12px;">
                            <option value="gsheet" selected>Đẩy trực tiếp Google Sheets</option>
                            <option value="excel">Chỉ tải file Excel (.xlsx)</option>
                            <option value="both">Cả hai phương thức trên</option>
                        </select>
                    </div>
                    <div style="margin-bottom: 8px;">
                        <label style="display: block; margin-bottom: 3px; font-size: 11px;">Đường dẫn Spreadsheet:</label>
                        <input type="text" id="scraper-sheet-url" placeholder="https://docs.google.com/spreadsheets/d/..." value="${sheetConfig.sheetUrl}" style="width: 95%; padding: 5px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white; font-size: 11px;">
                    </div>
                    <div style="margin-bottom: 8px;">
                        <label style="display: block; margin-bottom: 3px; font-size: 11px;">Tên Sheet Tab:</label>
                        <input type="text" id="scraper-sheet-tab" placeholder="Sheet1" value="${sheetConfig.tabName}" style="width: 95%; padding: 5px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white; font-size: 11px;">
                    </div>
                    <div style="margin-bottom: 8px;">
                        <label style="display: block; margin-bottom: 3px; font-size: 11px;">Google Client ID:</label>
                        <input type="password" id="scraper-client-id" placeholder="Client ID của bạn" value="${sheetConfig.clientId}" style="width: 95%; padding: 5px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white; font-size: 11px;">
                    </div>
                    <div style="margin-bottom: 8px;">
                        <label style="display: block; margin-bottom: 3px; font-size: 11px;">Google Client Secret:</label>
                        <input type="password" id="scraper-client-secret" placeholder="Client Secret của bạn" value="${sheetConfig.clientSecret}" style="width: 95%; padding: 5px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white; font-size: 11px;">
                    </div>
                    <div style="display: flex; gap: 5px; justify-content: flex-end; margin-bottom: 8px; flex-wrap: wrap;">
                        <button id="scraper-btn-save-sheet" style="padding: 4px 6px; border: none; border-radius: 3px; background-color: #0288d1; color: white; cursor: pointer; font-size: 11px;">Lưu cấu hình</button>
                        <button id="scraper-btn-connect-google" style="padding: 4px 6px; border: none; border-radius: 3px; background-color: #e65100; color: white; cursor: pointer; font-size: 11px; font-weight: bold;">Cấp quyền Google</button>
                        <button id="scraper-btn-test-sheet" style="padding: 4px 6px; border: none; border-radius: 3px; background-color: #6a1b9a; color: white; cursor: pointer; font-size: 11px; font-weight: bold;">Test Connect</button>
                    </div>
                    
                    <div id="scraper-auth-code-wrapper" style="border-top: 1px dashed #444; padding-top: 8px; display: none;">
                        <label style="display: block; margin-bottom: 3px; font-size: 11px; color: #ffeb3b;">Nhập mã Authorization Code:</label>
                        <div style="display: flex; gap: 5px;">
                            <input type="text" id="scraper-auth-code-input" placeholder="Dán mã Google cung cấp tại đây" style="flex: 1; padding: 4px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white; font-size: 11px;">
                            <button id="scraper-btn-confirm-code" style="padding: 4px 8px; border: none; border-radius: 3px; background-color: #4caf50; color: white; cursor: pointer; font-size: 11px; font-weight: bold;">Xác nhận</button>
                        </div>
                    </div>
                </div>
            </div>

            <div style="margin-bottom: 12px; display: flex; gap: 5px; justify-content: space-between; align-items: center; background-color: #2b2b2b; padding: 6px; border-radius: 4px;">
                <span>Đã cào: <strong id="scraper-count" style="color: #ee4d2d;">0</strong> SP</span>
                <span>Bộ nhớ đệm: <strong id="scraper-cache-count" style="color: #4caf50;">0</strong></span>
            </div>
            
            <div style="display: flex; gap: 5px; margin-bottom: 8px;">
                <button id="scraper-btn-start" style="flex: 1; padding: 8px; border: none; border-radius: 4px; background-color: #ee4d2d; color: white; font-weight: bold; cursor: pointer;">Start</button>
                <button id="scraper-btn-stop" style="flex: 1; padding: 8px; border: none; border-radius: 4px; background-color: #555; color: white; font-weight: bold; cursor: pointer;" disabled>Stop</button>
            </div>
            <div style="display: flex; gap: 5px; margin-bottom: 8px;">
                <button id="scraper-btn-clear-cache" style="width: 33%; padding: 6px; border: none; border-radius: 4px; background-color: #444; color: #ccc; cursor: pointer; font-size: 11px;">Xóa Cache</button>
                <button id="scraper-btn-sync-cache" style="width: 34%; padding: 6px; border: none; border-radius: 4px; background-color: #0288d1; color: white; cursor: pointer; font-size: 11px; font-weight: bold;">Đồng bộ Cache</button>
                <button id="scraper-btn-export" style="width: 33%; padding: 6px; border: none; border-radius: 4px; background-color: #2e7d32; color: white; cursor: pointer; font-size: 11px; font-weight: bold;">Xuất Excel</button>
            </div>
            <div style="margin-top: 10px;">
                <div style="font-weight: bold; margin-bottom: 4px;">System Logs:</div>
                <div id="scraper-logs" style="height: 110px; overflow-y: auto; background-color: #0c0c0c; border: 1px solid #333; padding: 6px; font-family: monospace; font-size: 11px; border-radius: 4px; word-break: break-all;">
                    <div style="color: #888;">Chờ lệnh Start...</div>
                </div>
            </div>
        `;
        document.body.appendChild(panel);

        // Accordion toggle logic
        document.getElementById('scraper-sheet-toggle').addEventListener('click', () => {
            const fields = document.getElementById('scraper-sheet-fields');
            const arrow = document.getElementById('scraper-toggle-arrow');
            if (fields.style.display === 'none') {
                fields.style.display = 'block';
                arrow.innerText = '▲';
            } else {
                fields.style.display = 'none';
                arrow.innerText = '▼';
            }
        });

        // Kết nối OAuth thủ công
        document.getElementById('scraper-btn-connect-google').addEventListener('click', () => {
            // Tự động lưu cấu hình trước khi mở flow để có Client ID
            sheetConfig.sheetUrl = document.getElementById('scraper-sheet-url').value.trim();
            sheetConfig.tabName = document.getElementById('scraper-sheet-tab').value.trim();
            sheetConfig.clientId = document.getElementById('scraper-client-id').value.trim();
            sheetConfig.clientSecret = document.getElementById('scraper-client-secret').value.trim();
            GM_setValue('sheet_config', sheetConfig);

            // Mở luồng lấy code và hiện ô input nhập code
            startGoogleAuthFlow();
            document.getElementById('scraper-auth-code-wrapper').style.display = 'block';
        });

        // Nút Xác nhận trao đổi Token
        document.getElementById('scraper-btn-confirm-code').addEventListener('click', () => {
            const code = document.getElementById('scraper-auth-code-input').value.trim();
            if (!code) {
                alert('Vui lòng dán mã Authorization Code trước khi xác nhận!');
                return;
            }
            exchangeAuthCodeForTokens(code);
            document.getElementById('scraper-auth-code-input').value = '';
            document.getElementById('scraper-auth-code-wrapper').style.display = 'none';
        });

        // Click nút Test Connect GSheet
        document.getElementById('scraper-btn-test-sheet').addEventListener('click', async () => {
            // Tự động lưu cấu hình trước khi chạy test
            sheetConfig.sheetUrl = document.getElementById('scraper-sheet-url').value.trim();
            sheetConfig.tabName = document.getElementById('scraper-sheet-tab').value.trim();
            sheetConfig.clientId = document.getElementById('scraper-client-id').value.trim();
            sheetConfig.clientSecret = document.getElementById('scraper-client-secret').value.trim();
            GM_setValue('sheet_config', sheetConfig);

            await testGoogleSheetsConnection();
        });

        // Click lưu cấu hình Google Sheets
        document.getElementById('scraper-btn-save-sheet').addEventListener('click', () => {
            sheetConfig.sheetUrl = document.getElementById('scraper-sheet-url').value.trim();
            sheetConfig.tabName = document.getElementById('scraper-sheet-tab').value.trim();
            sheetConfig.clientId = document.getElementById('scraper-client-id').value.trim();
            sheetConfig.clientSecret = document.getElementById('scraper-client-secret').value.trim();

            GM_setValue('sheet_config', sheetConfig);
            addLog('Đã lưu cấu hình Google Sheets thành công!', 'success');
            alert('Đã lưu thông tin cấu hình Google Sheets!');
        });

        // Lắng nghe sự kiện click các nút trên giao diện
        document.getElementById('scraper-btn-start').addEventListener('click', startScraper);
        document.getElementById('scraper-btn-stop').addEventListener('click', stopScraper);

        document.getElementById('scraper-btn-clear-cache').addEventListener('click', () => {
            if (confirm('Bạn có chắc chắn muốn xóa bộ nhớ đệm cache các sản phẩm đã cào không?')) {
                crawledCache = [];
                GM_setValue('crawled_items_cache', crawledCache);
                updateCacheCount();
                addLog('Đã xóa sạch bộ nhớ đệm cache', 'warn');
                alert('Đã xóa cache thành công!');
            }
        });

        document.getElementById('scraper-btn-export').addEventListener('click', exportToExcel);
        document.getElementById('scraper-btn-sync-cache').addEventListener('click', syncCacheFromGoogleSheets);
        updateCacheCount();
    }

    // Khởi chạy panel khi trang đã sẵn sàng
    if (document.readyState === 'complete' || document.readyState === 'interactive') {
        initPanel();
    } else {
        window.addEventListener('DOMContentLoaded', initPanel);
    }
    // Hỗ trợ cả trường hợp SPA của Shopee chuyển trang nhưng script đã chạy sẵn
    const observer = new MutationObserver(() => {
        if (!document.getElementById('shopee-scraper-panel') && document.body) {
            initPanel();
        }
    });
    observer.observe(document.documentElement, { childList: true, subtree: true });


    // Các hàm phụ trợ
    function sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    function cleanNumber(str) {
        if (!str) return 0;

        let s = str.trim().toLowerCase();

        // 1. Xử lý Lượt bán (ví dụ: "4k+ sold", "đã bán 4,2k", "đã bán 12.3k", "đã bán 12")
        // Tìm cụm số + chữ k (ví dụ: 4k, 4.2k, 4,2k)
        const kMatch = s.match(/([\d.,]+)\s*k/);
        if (kMatch) {
            let numStr = kMatch[1].replace(/,/g, '.'); // Chuyển dấu phẩy thành dấu chấm cho parseFloat
            let num = parseFloat(numStr);
            return isNaN(num) ? 0 : Math.round(num * 1000);
        }

        // 2. Xử lý số thông thường (ví dụ: "đã bán 150", hoặc giá tiền "669.000₫")
        // Loại bỏ ký hiệu tiền tệ ₫ và text
        s = s.replace(/₫/g, '').replace(/đ/g, '');

        // Lấy tất cả các chữ số và dấu chấm/phẩy phân tách hàng nghìn/phần thập phân
        // Shopee VN: "669.000", "1.130.000", "45.000"
        const numMatch = s.match(/[\d.,]+/g);
        if (!numMatch) return 0;

        // Lấy chuỗi số dài nhất tìm được (tránh lấy nhầm các số linh tinh khác)
        let longestNumStr = numMatch.reduce((a, b) => a.length > b.length ? a : b, '');

        // Nếu chuỗi chứa cả dấu chấm/dấu phẩy phân tách
        // Đối với giá tiền VN (ví dụ: 669.000 hoặc 1.250.000) ta bỏ hết dấu phân tách hàng nghìn
        if (longestNumStr.includes('.') || longestNumStr.includes(',')) {
            // Đếm số lượng dấu chấm và phẩy
            const dots = (longestNumStr.match(/\./g) || []).length;
            const commas = (longestNumStr.match(/,/g) || []).length;

            if (dots > 0) {
                // Ví dụ: 669.000 -> bỏ dấu chấm
                longestNumStr = longestNumStr.replace(/\./g, '');
            } else if (commas > 0) {
                // Ví dụ: 669,000 -> bỏ dấu phẩy
                longestNumStr = longestNumStr.replace(/,/g, '');
            }
        }

        let parsed = parseInt(longestNumStr);
        return isNaN(parsed) ? 0 : parsed;
    }

    // Lọc lấy ShopId và ItemId từ Url
    function parseProductUrl(href) {
        if (!href) return null;

        // Giải mã URL-encode nếu có (ví dụ %C3%81o -> Áo) để tránh lỗi Regex khi khớp chuỗi
        let decodedHref = href;
        try {
            decodedHref = decodeURIComponent(href);
        } catch (e) {
            // Giữ nguyên nếu không giải mã được
        }

        // Regex hỗ trợ cả link tương đối hoặc tuyệt đối có chứa định dạng i.shopId.itemId
        // Ví dụ: /Bộ-Quần-Áo-...-i.302865535.54005267037?extraParams=...
        const match = decodedHref.match(/i\.(\d+)\.(\d+)/);
        if (match) {
            return {
                shopId: match[1],
                itemId: match[2],
                productUrl: `https://shopee.vn/product/${match[1]}/${match[2]}`
            };
        }
        return null;
    }

    // Scroll từ từ xuống chân trang theo số lần chỉ định để load Lazy Items
    async function smoothScrollToBottom() {
        return new Promise((resolve) => {
            let currentPosition = window.scrollY || window.pageYOffset || 0;
            const distance = 400; // Mỗi lần cuộn 400px để đi được nhiều hơn trong 10 lần
            const interval = 200;  // Cuộn mỗi 200ms
            let scrollCount = 0;
            const maxScrolls = 10; // Giới hạn cuộn 10 lần

            // Đưa scroll về đỉnh trước khi bắt đầu cuộn xuống (đảm bảo trigger load từ đầu)
            window.scrollTo(0, 0);
            currentPosition = 0;

            let timer = setInterval(() => {
                // Cuộn xuống một đoạn
                currentPosition += distance;
                window.scrollTo(0, currentPosition);
                scrollCount++;

                // Điều kiện dừng: Khi đã cuộn đủ 10 lần hoặc vượt quá chiều cao trang thực tế
                let scrollHeight = document.documentElement.scrollHeight;
                if (scrollCount >= maxScrolls || currentPosition >= scrollHeight) {
                    clearInterval(timer);
                    resolve();
                }
            }, interval);
        });
    }

    // Hàm thực thi tìm kiếm và click Sort
    async function selectSortOption(sortVal) {
        const sortButtons = document.querySelectorAll('.shopee-sort-by-options__option, .shopee-sort-bar__label, .shopee-sort-by-options button');
        // Tìm button chứa giá trị tương ứng
        let targetText = '';
        if (sortVal === 'relevance') targetText = 'Liên quan';
        else if (sortVal === 'sales') targetText = 'Bán chạy';
        else if (sortVal === 'ctime') targetText = 'Mới nhất';
        else if (sortVal === 'pop') targetText = 'Phổ biến';

        for (let btn of sortButtons) {
            if (btn.textContent.trim().toLowerCase().includes(targetText.toLowerCase()) || btn.textContent.trim().toLowerCase().includes(sortVal)) {
                btn.click();
                await sleep(2000); // Chờ load trang mới sau khi click
                break;
            }
        }
    }

    // Hàm gọi API lấy thêm thông tin hoa hồng cho 1 sản phẩm (hỗ trợ thử lại tối đa 3 lần)
    function fetchCommissionData(itemId, retries = 3) {
        return new Promise((resolve) => {
            function attemptFetch(attempt) {
                GM_xmlhttpRequest({
                    method: 'GET',
                    url: `https://data.addlivetag.com/product-data/product-data.php?item_id=${itemId}`,
                    onload: function (response) {
                        if (response.status === 200) {
                            try {
                                const res = JSON.parse(response.responseText);
                                if (res.status === 'success' && res.productInfo) {
                                    resolve({
                                        commission: res.productInfo.commission || 0,
                                        sellerComFinal: res.productInfo.sellerComFinal || 0,
                                        shopeeComFinal: res.productInfo.shopeeComFinal || 0
                                    });
                                    return;
                                } else {
                                    handleError(`API trả về trạng thái thất bại hoặc thiếu productInfo`);
                                }
                            } catch (e) {
                                handleError(`Lỗi phân tích JSON: ${e.message}`);
                            }
                        } else {
                            handleError(`HTTP Status ${response.status}`);
                        }
                    },
                    onerror: function (err) {
                        handleError(err.message || 'Lỗi kết nối mạng');
                    }
                });

                async function handleError(errorMsg) {
                    if (attempt < retries) {
                        addLog(`[Thử lại ${attempt}/${retries}] Lỗi lấy hoa hồng SP ${itemId}: ${errorMsg}. Đang thử lại sau 1.5 giây...`, 'warn');
                        await sleep(1500);
                        attemptFetch(attempt + 1);
                    } else {
                        addLog(`[Lỗi] Thất bại khi lấy hoa hồng cho SP ${itemId} sau ${retries} lần thử: ${errorMsg}. Trả về mặc định 0.`, 'error');
                        resolve({ commission: 0, sellerComFinal: 0, shopeeComFinal: 0 });
                    }
                }
            }

            attemptFetch(1);
        });
    }

    // Hàm bổ sung thông tin hoa hồng cho toàn bộ sản phẩm đã cào trong crawledData
    async function enrichCrawledDataWithCommission() {
        if (crawledData.length === 0) return;

        addLog(`Bắt đầu lấy thông tin hoa hồng cho ${crawledData.length} sản phẩm từ API...`, 'info');
        setStatus('Enriching API...', '#009688');

        const batchSize = 5; // Chạy song song tối đa 5 requests một lúc
        const delayBetweenBatches = 1000; // Khoảng trễ 1 giây giữa các đợt để bảo đảm rate limit ~300 req/phút

        for (let i = 0; i < crawledData.length; i += batchSize) {
            if (!isRunning) break;

            const batch = crawledData.slice(i, i + batchSize);
            addLog(`Đang gọi API hoa hồng cho nhóm SP từ ${i + 1} đến ${Math.min(i + batchSize, crawledData.length)}...`, 'info');

            const promises = batch.map(async (item) => {
                const itemId = item['Mã SP (Item ID)'];
                if (itemId) {
                    const info = await fetchCommissionData(itemId);
                    item['Hoa hồng (đ)'] = info.commission;
                    item['Hoa hồng người bán (đ)'] = info.sellerComFinal;
                    item['Hoa hồng Shopee (đ)'] = info.shopeeComFinal;
                }
            });

            await Promise.all(promises);

            if (i + batchSize < crawledData.length) {
                await sleep(delayBetweenBatches);
            }
        }
        addLog('Hoàn thành bổ sung thông tin hoa hồng!', 'success');
    }

    // Hàm cào trang hiện tại
    async function crawlCurrentPage() {
        if (!isRunning) return;

        addLog('Đang cuộn chuột xuống đáy trang để tải thêm sản phẩm...', 'info');
        setStatus('Scrolling...', '#ff9800');
        // Cuộn xuống từ từ để kích hoạt lazy load
        await smoothScrollToBottom();
        await sleep(1500); // Chờ thêm một chút cho hình ảnh và giá load hẳn

        setStatus('Analyzing...', '#2196f3');
        const items = document.querySelectorAll('[data-sqe="item"]');
        addLog(`Tìm thấy tổng cộng ${items.length} thẻ sản phẩm trên trang này`, 'info');

        if (items.length === 0) {
            addLog('Không tìm thấy sản phẩm nào! Có thể bạn đang ở sai trang Shopee hoặc cấu trúc web đã đổi.', 'error');
        }

        let countThisPage = 0;

        const priceMin = parseInt(document.getElementById('scraper-price-min').value) || 0;
        const soldMin = parseInt(document.getElementById('scraper-sold-min').value) || 0;

        for (let i = 0; i < items.length; i++) {
            const item = items[i];
            if (!isRunning) return;

            try {
                // 1. Kiểm tra ảnh Xtra đặc trưng
                // Hỗ trợ cả 2 mẫu nhãn Xtra:
                // Mẫu 1: 11134258-7r98o-lym830xtkwipb8 (dạng ảnh cũ hoặc resize)
                // Mẫu 2: 3a4abff6345de81bad16.png (dạng ảnh mới lưu trên CDN Shopee)
                const xtraImg = item.querySelector(
                    'img[src*="11134258-7r98o-lym830xtkwipb8"], ' +
                    'img[srcset*="11134258-7r98o-lym830xtkwipb8"], ' +
                    'img[src*="3a4abff6345de81bad16.png"], ' +
                    'img[srcset*="3a4abff6345de81bad16.png"]'
                );
                if (!xtraImg) {
                    continue; // Bỏ qua nếu không có nhãn Xtra này
                }

                // 2. Lấy liên kết & kiểm tra trùng lặp
                // Thêm lựa chọn tìm thẻ a.contents hoặc a thông thường
                const aTag = item.querySelector('a.contents, a');
                if (!aTag) {
                    addLog(`[SP #${i + 1}] Thẻ <a> không tồn tại`, 'warn');
                    continue;
                }
                const hrefAttr = aTag.getAttribute('href');
                const parsed = parseProductUrl(hrefAttr);
                if (!parsed) {
                    addLog(`[SP #${i + 1}] Không phân tách được URL: ${hrefAttr}`, 'warn');
                    continue;
                }

                // Lấy tên sản phẩm (Title) trước để log cho rõ
                let title = "";
                // Ưu tiên 1: Lấy từ thuộc tính aria-label của thẻ cha 'Product card: [Tên sản phẩm]'
                const cardWrapper = item.querySelector('[aria-label^="Product card:"]');
                if (cardWrapper) {
                    const fullAria = cardWrapper.getAttribute('aria-label');
                    title = fullAria.replace("Product card:", "").trim();
                }
                // Ưu tiên 2: Tìm thẻ chứa class line-clamp-2
                if (!title) {
                    const titleNode = item.querySelector('div.whitespace-normal.line-clamp-2');
                    if (titleNode) {
                        title = titleNode.textContent.trim();
                    }
                }
                // Ưu tiên 3: Lấy từ thuộc tính alt của ảnh chính
                if (!title) {
                    const mainImg = item.querySelector('img[alt]:not([alt="promotion-label"]):not([alt="custom-overlay"]):not([alt="flag-label"])');
                    if (mainImg) {
                        title = mainImg.getAttribute('alt') || "";
                    }
                }
                const shortTitle = title.length > 25 ? title.substring(0, 25) + '...' : title;

                // Kiểm tra cache trùng lặp
                if (crawledCache.includes(parsed.itemId)) {
                    addLog(`Bỏ qua: "${shortTitle}" (Đã cào trước đó - Trùng Cache)`, 'warn');
                    continue;
                }

                // 3. Kiểm tra Giá (Price)
                let priceVal = 0;

                // Tìm thẻ chứa giá bán gần chính xác nhất (thường là span có class chứa text-base)
                const priceNode = item.querySelector('[aria-label="promotion price"] ~ div span.truncate, span.truncate.text-base\\/5, .text-base\\/5');
                if (priceNode) {
                    priceVal = cleanNumber(priceNode.textContent);
                } else {
                    const promotionPriceNode = item.querySelector('[aria-label="promotion price"]');
                    if (promotionPriceNode) {
                        const parent = promotionPriceNode.parentElement;
                        if (parent) {
                            priceVal = cleanNumber(parent.textContent);
                        }
                    }
                }

                if (priceVal < priceMin) {
                    addLog(`Bỏ qua: "${shortTitle}" (Giá ${priceVal}₫ < ${priceMin}₫)`, 'warn');
                    continue;
                }

                // 4. Kiểm tra Lượt bán (Sold)
                let soldVal = 0;
                // Tìm thẻ cụ thể chứa lượt bán (Thường nằm ở góc phải bên dưới hoặc trong khối chứa text "sold" hoặc "đã bán" độc lập)
                const textNodes = Array.from(item.querySelectorAll('*'));
                const soldNode = textNodes.find(node => {
                    const text = node.textContent.trim().toLowerCase();
                    // Nhận diện "sold", "đã bán", "sold/month", "đã bán/tháng"
                    const hasKeyword = text.includes('sold') || text.includes('đã bán');
                    // Đảm bảo node có chứa số, độ dài text ngắn và là node lá (không có thẻ con bên trong)
                    // để tránh lấy nhầm node cha bọc cả Rating (4.9) dẫn đến việc ghép chữ thành 4.97k
                    return hasKeyword && /\d/.test(text) && text.length < 30 && node.children.length === 0;
                });

                if (soldNode) {
                    soldVal = cleanNumber(soldNode.textContent);
                } else {
                    // Dự phòng nếu cấu trúc thay đổi, tìm theo class thông dụng hoặc thuộc tính text có số bán
                    const fallbackSold = item.querySelector('div[class*="sold"], div[class*="black87"]');
                    if (fallbackSold) {
                        soldVal = cleanNumber(fallbackSold.textContent);
                    }
                }

                if (soldVal < soldMin) {
                    addLog(`Bỏ qua: "${shortTitle}" (Đã bán ${soldVal} < ${soldMin})`, 'warn');
                    continue;
                }

                // Thỏa mãn toàn bộ điều kiện -> Lưu kết quả
                crawledData.push({
                    'Mã SP (Item ID)': parsed.itemId,
                    'Mã Shop (Shop ID)': parsed.shopId,
                    'Tên sản phẩm': title,
                    'Giá (đ)': priceVal,
                    'Lượt bán': soldVal,
                    'Đường dẫn sản phẩm': parsed.productUrl
                });

                // Đưa vào cache đã cào để không cào lại nữa
                crawledCache.push(parsed.itemId);
                countThisPage++;
                addLog(`ĐÃ CÀO THÀNH CÔNG SP: "${shortTitle}" (Giá: ${priceVal}đ | Bán: ${soldVal})`, 'success');

                // Cập nhật UI ngay lập tức
                document.getElementById('scraper-count').innerText = crawledData.length;
                GM_setValue('crawled_items_cache', crawledCache);
                updateCacheCount();
            } catch (err) {
                addLog(`Lỗi khi xử lý SP #${i + 1}: ${err.message}`, 'error');
            }
        }

        addLog(`Hoàn thành cào trang này. Thu thập thêm được ${countThisPage} SP đạt chuẩn.`, 'success');

        // Tự động tìm và bấm chuyển trang tiếp theo
        const nextPageBtn = document.querySelector('.shopee-icon-button--right:not([aria-disabled="true"]):not([disabled])');
        if (nextPageBtn && isRunning) {
            addLog('Phát hiện nút chuyển trang. Đang chuyển sang trang tiếp theo sau 3 giây...', 'info');
            setStatus('Next Page...', '#9c27b0');
            nextPageBtn.click();
            await sleep(3500); // Đợi trang tiếp theo load
            await crawlCurrentPage(); // Tiếp tục đệ quy cào trang tiếp
        } else {
            addLog('Đã cào xong tất cả các trang hoặc người dùng dừng!', 'success');
            setStatus('Processing...', '#2196f3');

            // Lấy thêm thông tin hoa hồng từ API
            await enrichCrawledDataWithCommission();

            addLog('Mặc định: Đang đẩy dữ liệu lên Google Sheets...', 'info');
            try {
                await pushToGoogleSheets();
            } catch (err) {
                addLog(`Lỗi đẩy dữ liệu lên Google Sheets: ${err.message}`, 'error');
                addLog('Kích hoạt chế độ dự phòng (Fallback): Tự động tải file Excel...', 'warn');
                exportToExcel();
            }

            setStatus('Done!', '#4caf50');
            stopScraper();
            showDoneOverlay();
        }
    }

    // Hàm xuất dữ liệu Excel dùng XLSX.js
    function exportToExcel() {
        if (crawledData.length === 0) {
            addLog('Không có dữ liệu để xuất Excel!', 'warn');
            return;
        }

        const worksheet = XLSX.utils.json_to_sheet(crawledData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Sản phẩm");

        // Tạo tên file tự động theo định dạng ngày_tháng_năm_giờ_phut
        const now = new Date();
        const dateStr = `${now.getDate()}_${now.getMonth() + 1}_${now.getFullYear()}_${now.getHours()}h${now.getMinutes()}m`;
        const fileName = `product_${dateStr}.xlsx`;

        XLSX.writeFile(workbook, fileName);
        addLog(`Đã xuất và tải về file Excel dự phòng: ${fileName}`, 'success');
        alert(`Gặp lỗi kết nối Google Sheets! Đã tự động tải file Excel dự phòng thay thế: ${fileName}`);
    }

    // Trích xuất Spreadsheet ID từ URL
    function getSpreadsheetId(url) {
        if (!url) return null;
        const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
        return match ? match[1] : null;
    }

    // Hàm mở luồng OAuth lấy code
    function startGoogleAuthFlow() {
        const client_id = sheetConfig.clientId;
        if (!client_id) {
            addLog('Lỗi: Hãy điền Client ID trước khi Kết nối Google!', 'error');
            alert('Vui lòng điền Google Client ID trước!');
            return;
        }

        // Sử dụng redirect_uri chuẩn hỗ trợ copy paste là: urn:ietf:wg:oauth:2.0:oob
        const redirect_uri = 'urn:ietf:wg:oauth:2.0:oob';
        const authUrl = `https://accounts.google.com/o/oauth2/v2/auth?` +
            `client_id=${encodeURIComponent(client_id)}` +
            `&redirect_uri=${encodeURIComponent(redirect_uri)}` +
            `&response_type=code` +
            `&scope=${encodeURIComponent('https://www.googleapis.com/auth/spreadsheets')}` +
            `&access_type=offline` +
            `&prompt=consent`;

        addLog('Đang mở tab Google cấp quyền. Vui lòng làm theo hướng dẫn...', 'info');
        window.open(authUrl, '_blank');
        addLog('Vui lòng dán mã Authorization Code nhận được vào ô nhập liệu bên dưới và nhấn "Xác nhận".', 'warn');
    }

    // Gửi yêu cầu đổi mã code sang Token
    function exchangeAuthCodeForTokens(authCode) {
        const client_id = sheetConfig.clientId;
        const client_secret = sheetConfig.clientSecret;
        const redirect_uri = 'urn:ietf:wg:oauth:2.0:oob';

        addLog('Đang trao đổi mã xác thực lấy Token...', 'info');
        GM_xmlhttpRequest({
            method: 'POST',
            url: 'https://oauth2.googleapis.com/token',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            data: `code=${encodeURIComponent(authCode)}` +
                `&client_id=${encodeURIComponent(client_id)}` +
                `&client_secret=${encodeURIComponent(client_secret)}` +
                `&redirect_uri=${encodeURIComponent(redirect_uri)}` +
                `&grant_type=authorization_code`,
            onload: function (response) {
                try {
                    const res = JSON.parse(response.responseText);
                    if (res.access_token) {
                        GM_setValue('google_oauth_token', res.access_token);
                        GM_setValue('google_token_expiry', Date.now() + (res.expires_in * 1000));
                        if (res.refresh_token) {
                            GM_setValue('google_refresh_token', res.refresh_token);
                        }
                        addLog('KẾT NỐI GOOGLE SPREADSHEETS THÀNH CÔNG!', 'success');
                        alert('Đăng nhập và Kết nối Google Sheets thành công!');
                    } else {
                        addLog(`Lỗi phản hồi Google: ${response.responseText}`, 'error');
                        alert(`Lỗi: ${res.error_description || 'Không thể lấy token. Vui lòng kiểm tra lại Client Secret / ID'}`);
                    }
                } catch (e) {
                    addLog(`Lỗi xử lý JSON: ${e.message}`, 'error');
                }
            },
            onerror: (err) => {
                addLog(`Lỗi kết nối mạng: ${err.message}`, 'error');
            }
        });
    }

    // Hàm lấy Access Token từ Google OAuth 2.0 Auth Flow (Dành cho ứng dụng của người dùng)
    async function getGoogleAccessToken() {
        return new Promise((resolve, reject) => {
            const storedToken = GM_getValue('google_oauth_token', null);
            const tokenExpiry = GM_getValue('google_token_expiry', 0);

            // Nếu Token vẫn còn hạn (trước 5 phút hết hạn để an toàn)
            if (storedToken && Date.now() < tokenExpiry - 300000) {
                resolve(storedToken);
                return;
            }

            const client_id = sheetConfig.clientId;
            const client_secret = sheetConfig.clientSecret;
            const refresh_token = GM_getValue('google_refresh_token', null);

            if (!client_id || !client_secret) {
                addLog('Lỗi: Hãy điền Client ID và Client Secret, sau đó nhấn "Lưu cấu hình" & "Kết nối Google"', 'error');
                reject(new Error('Missing OAuth credentials'));
                return;
            }

            if (!refresh_token) {
                addLog('Chưa có Refresh Token. Vui lòng bấm nút "Kết nối Google" trên Panel trước!', 'error');
                reject(new Error('No refresh token available'));
                return;
            }

            // Đã có Refresh Token, gia hạn Access Token mới
            addLog('Gia hạn Access Token tự động...', 'info');
            GM_xmlhttpRequest({
                method: 'POST',
                url: 'https://oauth2.googleapis.com/token',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                data: `refresh_token=${encodeURIComponent(refresh_token)}` +
                    `&client_id=${encodeURIComponent(client_id)}` +
                    `&client_secret=${encodeURIComponent(client_secret)}` +
                    `&grant_type=refresh_token`,
                onload: function (response) {
                    try {
                        const res = JSON.parse(response.responseText);
                        if (res.access_token) {
                            GM_setValue('google_oauth_token', res.access_token);
                            GM_setValue('google_token_expiry', Date.now() + (res.expires_in * 1000));
                            addLog('Tự động gia hạn Access Token thành công!', 'success');
                            resolve(res.access_token);
                        } else {
                            addLog(`Lỗi gia hạn Token (hết hạn hoặc thu hồi): ${response.responseText}`, 'error');
                            GM_setValue('google_refresh_token', null); // Xoá refresh token lỗi để yêu cầu kết nối lại
                            reject(new Error('Refresh token invalid'));
                        }
                    } catch (e) {
                        reject(e);
                    }
                },
                onerror: (err) => reject(err)
            });
        });
    }

    // Hàm chính đẩy dữ liệu lên Google Sheets
    async function pushToGoogleSheets() {
        return new Promise(async (resolve, reject) => {
            if (crawledData.length === 0) {
                addLog('Không có dữ liệu sản phẩm nào để đẩy lên Google Sheets!', 'warn');
                resolve();
                return;
            }

            const spreadsheetId = getSpreadsheetId(sheetConfig.sheetUrl);
            if (!spreadsheetId) {
                addLog('Lỗi: URL Google Sheets không hợp lệ!', 'error');
                reject(new Error('URL Google Sheets không hợp lệ'));
                return;
            }

            try {
                addLog('Đang kết nối xác thực tài khoản Google...', 'info');
                const token = await getGoogleAccessToken();
                addLog('Xác thực thành công. Đang ghi dữ liệu lên Google Sheet...', 'info');

                // Định dạng dữ liệu dạng mảng hai chiều để ghi API Sheets
                // Tiêu đề cột
                const headers = ['Mã SP (Item ID)', 'Mã Shop (Shop ID)', 'Tên sản phẩm', 'Giá (đ)', 'Lượt bán', 'Đường dẫn sản phẩm', 'Hoa hồng (đ)', 'Hoa hồng người bán (đ)', 'Hoa hồng Shopee (đ)'];
                const rows = crawledData.map(item => [
                    item['Mã SP (Item ID)'],
                    item['Mã Shop (Shop ID)'],
                    item['Tên sản phẩm'],
                    item['Giá (đ)'],
                    item['Lượt bán'],
                    item['Đường dẫn sản phẩm'],
                    item['Hoa hồng (đ)'] || 0,
                    item['Hoa hồng người bán (đ)'] || 0,
                    item['Hoa hồng Shopee (đ)'] || 0
                ]);

                // Append dữ liệu (Viết nối tiếp vào bảng)
                const range = `${sheetConfig.tabName}!A:I`;
                const valueInputOption = 'USER_ENTERED';

                GM_xmlhttpRequest({
                    method: 'POST',
                    url: `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}:append?valueInputOption=${valueInputOption}`,
                    headers: {
                        'Authorization': `Bearer ${token}`,
                        'Content-Type': 'application/json'
                    },
                    data: JSON.stringify({
                        values: [headers, ...rows]
                    }),
                    onload: function (response) {
                        if (response.status === 200) {
                            addLog(`GHI ĐỦ ${crawledData.length} SP LÊN GOOGLE SHEETS THÀNH CÔNG!`, 'success');
                            resolve();
                        } else {
                            addLog(`Lỗi ghi Sheets (Mã lỗi ${response.status}): ${response.responseText}`, 'error');
                            // Nếu do Token lỗi, clear cache token để login lại
                            if (response.status === 401) {
                                GM_setValue('google_oauth_token', null);
                            }
                            reject(new Error(`Google API respond with status ${response.status}`));
                        }
                    },
                    onerror: function (err) {
                        addLog(`Lỗi kết nối API Google: ${err.message}`, 'error');
                        reject(err);
                    }
                });

            } catch (e) {
                reject(e);
            }
        });
    }

    // Hàm đồng bộ bộ nhớ đệm từ cột A của Google Sheets
    async function syncCacheFromGoogleSheets() {
        const spreadsheetId = getSpreadsheetId(sheetConfig.sheetUrl);
        if (!spreadsheetId) {
            addLog('Lỗi đồng bộ: URL Google Sheets không đúng định dạng!', 'error');
            alert('Vui lòng cấu hình URL Spreadsheet hợp lệ trước khi đồng bộ!');
            return;
        }

        try {
            addLog('Đang kết nối xác thực Google để đồng bộ bộ nhớ đệm...', 'info');
            const token = await getGoogleAccessToken();
            addLog('Đang tải dữ liệu cột A từ Google Sheet...', 'info');

            // Đọc cột A từ dòng thứ 2 (bỏ dòng tiêu đề A1)
            const range = `${sheetConfig.tabName}!A2:A`;

            GM_xmlhttpRequest({
                method: 'GET',
                url: `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}`,
                headers: {
                    'Authorization': `Bearer ${token}`
                },
                onload: function (response) {
                    if (response.status === 200) {
                        try {
                            const data = JSON.parse(response.responseText);
                            const values = data.values; // Mảng hai chiều: [[id1], [id2], ...]
                            
                            if (!values || values.length === 0) {
                                addLog('Đồng bộ hoàn tất: Không tìm thấy dữ liệu nào ở cột A hoặc tab trống. Bộ nhớ đệm cũ được giữ nguyên.', 'warn');
                                alert('Không tìm thấy dữ liệu sản phẩm nào trên Google Sheets để đồng bộ!');
                                return;
                            }

                            // Trích xuất các ID sản phẩm hợp lệ từ cột A
                            const sheetIds = values
                                .map(row => row[0]) // Lấy giá trị đầu tiên của mỗi dòng
                                .filter(val => val !== undefined && val !== null)
                                .map(val => String(val).trim())
                                .filter(val => /^\d+$/.test(val)); // Lọc chỉ giữ lại các chuỗi số nguyên hợp lệ (Mã SP)

                            if (sheetIds.length === 0) {
                                addLog('Đồng bộ hoàn tất: Cột A không chứa mã sản phẩm hợp lệ nào. Bộ nhớ đệm cũ được giữ nguyên.', 'warn');
                                alert('Không tìm thấy mã sản phẩm hợp lệ nào ở cột A của Google Sheets!');
                                return;
                            }

                            // Gộp với bộ nhớ đệm hiện có và loại bỏ trùng lặp
                            const beforeCount = crawledCache.length;
                            crawledCache = Array.from(new Set([...crawledCache, ...sheetIds]));
                            const addedCount = crawledCache.length - beforeCount;

                            GM_setValue('crawled_items_cache', crawledCache);
                            updateCacheCount();

                            addLog(`ĐỒNG BỘ CACHE THÀNH CÔNG!`, 'success');
                            addLog(`- Đã quét từ Sheet: ${sheetIds.length} mã sản phẩm.`, 'success');
                            addLog(`- Đã thêm mới vào Cache: ${addedCount} mã sản phẩm.`, 'success');
                            addLog(`- Tổng số cache hiện tại: ${crawledCache.length} mã sản phẩm.`, 'success');

                            alert(`Đồng bộ bộ nhớ đệm thành công!\nĐã thêm mới: ${addedCount} mã sản phẩm.\nTổng số trong Cache: ${crawledCache.length}.`);
                        } catch (err) {
                            addLog(`Lỗi xử lý dữ liệu đồng bộ: ${err.message}`, 'error');
                            alert(`Lỗi xử lý dữ liệu: ${err.message}`);
                        }
                    } else {
                        addLog(`Đồng bộ thất bại (Mã lỗi ${response.status}): ${response.responseText}`, 'error');
                        alert(`Không thể kết nối lấy dữ liệu Google Sheet để đồng bộ!\nMã lỗi: ${response.status}`);
                        if (response.status === 401) {
                            GM_setValue('google_oauth_token', null);
                        }
                    }
                },
                onerror: function (err) {
                    addLog(`Lỗi mạng khi kết nối đồng bộ: ${err.message}`, 'error');
                    alert(`Lỗi kết nối mạng khi đồng bộ: ${err.message}`);
                }
            });
        } catch (e) {
            addLog(`Lỗi quá trình đồng bộ: ${e.message}`, 'error');
            alert(`Lỗi xác thực đồng bộ: ${e.message}`);
        }
    }

    // Hàm kiểm tra kết nối Google Sheets (Test Connect)
    async function testGoogleSheetsConnection() {
        const spreadsheetId = getSpreadsheetId(sheetConfig.sheetUrl);
        if (!spreadsheetId) {
            addLog('Lỗi Test Connect: URL Google Sheets không đúng định dạng!', 'error');
            alert('Vui lòng nhập link Spreadsheet hợp lệ!');
            return;
        }

        try {
            addLog('Đang chuẩn bị kiểm tra kết nối với Google API...', 'info');
            const token = await getGoogleAccessToken();
            addLog('Đang lấy Metadata của Spreadsheet từ Google...', 'info');

            GM_xmlhttpRequest({
                method: 'GET',
                url: `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?fields=properties.title,sheets.properties.title`,
                headers: {
                    'Authorization': `Bearer ${token}`
                },
                onload: function (response) {
                    if (response.status === 200) {
                        try {
                            const data = JSON.parse(response.responseText);
                            const title = data.properties.title;
                            const sheets = data.sheets.map(s => s.properties.title);

                            addLog(`TEST KẾT NỐI THÀNH CÔNG!`, 'success');
                            addLog(`- Tên Spreadsheet: "${title}"`, 'success');
                            addLog(`- Các Sheet Tab tìm thấy: [${sheets.join(', ')}]`, 'info');

                            // Kiểm tra xem tab chỉ định có tồn tại không
                            if (sheets.includes(sheetConfig.tabName)) {
                                addLog(`- Tab chỉ định "${sheetConfig.tabName}": HỢP LỆ!`, 'success');
                                alert(`Kết nối thành công!\nTên bảng tính: ${title}\nTab "${sheetConfig.tabName}" sẵn sàng ghi dữ liệu.`);
                            } else {
                                addLog(`- Tab chỉ định "${sheetConfig.tabName}": KHÔNG TÌM THẤY!`, 'warn');
                                addLog(`Lưu ý: API sẽ tự động tạo tab mới nếu bạn tiến hành chạy tool.`, 'warn');
                                alert(`Kết nối thành công!\nTên bảng tính: ${title}\nCảnh báo: Tab "${sheetConfig.tabName}" không tồn tại. API sẽ tự động tạo tab mới này khi đẩy dữ liệu.`);
                            }
                        } catch (err) {
                            addLog(`Lỗi parse JSON kết quả: ${err.message}`, 'error');
                        }
                    } else {
                        addLog(`Test thất bại (Mã lỗi ${response.status}): ${response.responseText}`, 'error');
                        alert(`Kết nối thất bại!\nMã lỗi: ${response.status}\nHãy kiểm tra lại Client Secret, ID hoặc quyền chia sẻ của Google Sheet.`);
                        if (response.status === 401) {
                            GM_setValue('google_oauth_token', null);
                        }
                    }
                },
                onerror: function (err) {
                    addLog(`Lỗi kết nối mạng khi Test: ${err.message}`, 'error');
                    alert(`Không thể kết nối tới Google Server: ${err.message}`);
                }
            });
        } catch (e) {
            addLog(`Lỗi quá trình kiểm tra: ${e.message}`, 'error');
            alert(`Lỗi xác thực: ${e.message}`);
        }
    }

    // Cập nhật trạng thái Panel
    function setStatus(text, color = '#aaa') {
        const statusEl = document.getElementById('scraper-status');
        if (statusEl) {
            statusEl.innerText = text;
            statusEl.style.backgroundColor = color;
            statusEl.style.color = '#fff';
        }
    }

    // Bắt đầu cào
    async function startScraper() {
        if (isRunning) return;
        isRunning = true;
        crawledData = []; // Reset dữ liệu của lượt chạy này

        document.getElementById('scraper-btn-start').disabled = true;
        document.getElementById('scraper-btn-stop').disabled = false;
        document.getElementById('scraper-btn-start').style.backgroundColor = '#555';
        document.getElementById('scraper-btn-stop').style.backgroundColor = '#ee4d2d';

        // Áp dụng tùy chọn Sort
        const sortVal = document.getElementById('scraper-sort').value;
        setStatus('Sorting...', '#ff9800');
        await selectSortOption(sortVal);

        // Bắt đầu tiến trình cào đệ quy
        await crawlCurrentPage();
    }

    // Dừng cào
    function stopScraper() {
        isRunning = false;
        const startBtn = document.getElementById('scraper-btn-start');
        const stopBtn = document.getElementById('scraper-btn-stop');
        if (startBtn) {
            startBtn.disabled = false;
            startBtn.style.backgroundColor = '#ee4d2d';
        }
        if (stopBtn) {
            stopBtn.disabled = true;
            stopBtn.style.backgroundColor = '#555';
        }
        setStatus('Stopped', '#d32f2f');
    }

    // Hiển thị lớp phủ màu xanh lá cây nhạt kèm thông báo Done!
    function showDoneOverlay() {
        // Xóa lớp phủ cũ nếu có
        const existing = document.getElementById('scraper-done-overlay');
        if (existing) existing.remove();

        const overlay = document.createElement('div');
        overlay.id = 'scraper-done-overlay';
        overlay.style.position = 'fixed';
        overlay.style.top = '0';
        overlay.style.left = '0';
        overlay.style.width = '100vw';
        overlay.style.height = '100vh';
        overlay.style.backgroundColor = 'rgba(46, 125, 50, 0.15)'; // Xanh lá cây nhạt
        overlay.style.backdropFilter = 'blur(6px)';
        overlay.style.webkitBackdropFilter = 'blur(6px)';
        overlay.style.zIndex = '9999999'; // Nằm trên tất cả các phần tử
        overlay.style.display = 'flex';
        overlay.style.justifyContent = 'center';
        overlay.style.alignItems = 'center';
        overlay.style.transition = 'opacity 0.4s ease';
        overlay.style.opacity = '0';

        const card = document.createElement('div');
        card.style.backgroundColor = '#1a1a1a';
        card.style.border = '2px solid #4caf50';
        card.style.borderRadius = '12px';
        card.style.padding = '35px 60px';
        card.style.boxShadow = '0 10px 30px rgba(0,0,0,0.6)';
        card.style.textAlign = 'center';
        card.style.color = '#ffffff';
        card.style.fontFamily = 'Arial, sans-serif';
        card.style.transform = 'scale(0.8)';
        card.style.transition = 'transform 0.3s cubic-bezier(0.175, 0.885, 0.32, 1.275)';

        card.innerHTML = `
            <div style="font-size: 50px; color: #4caf50; margin-bottom: 12px; animation: scraper-bounce 0.8s infinite alternate;">🎉</div>
            <div style="font-size: 32px; font-weight: bold; margin-bottom: 8px; color: #4caf50; letter-spacing: 1px;">Done!</div>
            <div style="font-size: 14px; color: #ccc; margin-bottom: 25px;">Đã cào dữ liệu và đồng bộ hoàn tất.</div>
            <button id="scraper-btn-close-overlay" style="padding: 10px 30px; font-size: 14px; font-weight: bold; color: white; background-color: #4caf50; border: none; border-radius: 6px; cursor: pointer; transition: background-color 0.2s; box-shadow: 0 4px 10px rgba(76, 175, 80, 0.3);">OK</button>
        `;

        // Tạo style cho hiệu ứng bounce và hover của button OK
        const style = document.createElement('style');
        style.id = 'scraper-overlay-style';
        style.innerHTML = `
            @keyframes scraper-bounce {
                from { transform: translateY(0); }
                to { transform: translateY(-10px); }
            }
            #scraper-btn-close-overlay:hover {
                background-color: #388e3c !important;
            }
        `;
        document.head.appendChild(style);

        overlay.appendChild(card);
        document.body.appendChild(overlay);

        // Hiệu ứng Fade In
        setTimeout(() => {
            overlay.style.opacity = '1';
            card.style.transform = 'scale(1)';
        }, 50);

        // Event đóng lớp phủ khi click nút OK
        const closeBtn = card.querySelector('#scraper-btn-close-overlay');
        closeBtn.addEventListener('click', () => {
            overlay.style.opacity = '0';
            card.style.transform = 'scale(0.8)';
            setTimeout(() => {
                overlay.remove();
                const styleEl = document.getElementById('scraper-overlay-style');
                if (styleEl) styleEl.remove();
            }, 400);
        });
    }

})();
