// ==UserScript==
// @name         Shopee Vietnam Product Scraper
// @namespace    http://tampermonkey.net/
// @version      1.6
// @description  Cào dữ liệu sản phẩm Shopee VN theo bộ lọc tùy chỉnh và đồng bộ Google Sheets / Xuất Excel
// @author       Antigravity
// @match        https://shopee.vn/*
// @match        https://affiliate.shopee.vn/*
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_registerMenuCommand
// @grant        GM_xmlhttpRequest
// @grant        GM_addValueChangeListener
// @grant        GM_removeValueChangeListener
// @connect      affiliate.shopee.vn
// @require      https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js
// @run-at       document-start
// ==/UserScript==

(function () {
    'use strict';

    const CATEGORIES_DATA = {
        "Balo & Túi Ví Nam": [
            "Ba Lô Nam",
            "Ba Lô Laptop Nam",
            "Túi & Cặp Đựng Laptop",
            "Túi Chống Sốc Laptop Nam",
            "Túi Tote Nam",
            "Cặp Xách Công Sở Nam",
            "Ví Cầm Tay Nam",
            "Túi Đeo Hông & Túi Đeo Ngực Nam",
            "Túi Đeo Chéo Nam",
            "Bóp/Ví Nam",
            "Khác"
        ],
        "Bách Hóa Online": [
            "Đồ ăn vặt",
            "Đồ chế biến sẵn",
            "Nhu yếu phẩm",
            "Nguyên liệu nấu ăn",
            "Đồ làm bánh",
            "Sữa - trứng",
            "Đồ uống",
            "Ngũ cốc & mứt",
            "Các loại bánh",
            "Đồ uống có cồn",
            "Bộ quà tặng",
            "Thực phẩm tươi sống và thực phẩm đông lạnh",
            "Khác"
        ],
        "Chăm Sóc Thú Cưng": [
            "Thức ăn cho thú cưng",
            "Phụ kiện cho thú cưng",
            "Vệ sinh cho thú cưng",
            "Quần áo thú cưng",
            "Chăm sóc sức khỏe",
            "Làm đẹp cho thú cưng",
            "Khác"
        ],
        "Dụng cụ và thiết bị tiện ích": [
            "Dụng cụ cầm tay",
            "Dụng cụ điện và thiết bị lớn",
            "Thiết bị mạch điện",
            "Vật liệu xây dựng",
            "Thiết bị và phụ kiện xây dựng"
        ],
        "Điện Thoại & Phụ Kiện": [
            "Điện thoại",
            "Máy tính bảng",
            "Pin Dự Phòng",
            "Pin Gắn Trong, Cáp và Bộ Sạc",
            "Ốp lưng, bao da, Miếng dán điện thoại",
            "Bảo vệ màn hình",
            "Đế giữ điện thoại",
            "Thẻ nhớ",
            "Sim",
            "Phụ kiện khác",
            "Thiết bị khác"
        ],
        "Đồ Chơi": [
            "Sở thích & Sưu tầm",
            "Đồ chơi giải trí",
            "Đồ chơi giáo dục",
            "Đồ chơi cho trẻ sơ sinh & trẻ nhỏ",
            "Đồ chơi vận động & ngoài trời",
            "Búp bê & Đồ chơi nhồi bông"
        ],
        "Đồng Hồ": [
            "Đồng Hồ Nam",
            "Đồng Hồ Nữ",
            "Bộ Đồng Hồ & Đồng Hồ Cặp",
            "Đồng Hồ Trẻ Em",
            "Phụ Kiện Đồng Hồ",
            "Khác"
        ],
        "Giày Dép Nam": [
            "Bốt",
            "Giày Thể Thao/ Sneakers",
            "Giày Sục",
            "Giày Tây Lười",
            "Giày Oxfords & Giày Buộc Dây",
            "Xăng-đan và Dép",
            "Phụ kiện giày dép",
            "Khác"
        ],
        "Giày Dép Nữ": [
            "Bốt",
            "Giày Thể Thao/ Sneaker",
            "Giày Đế Bằng",
            "Giày Cao Gót",
            "Giày Đế Xuồng",
            "Xăng-đan Và Dép",
            "Phụ Kiện Giày",
            "Giày Khác"
        ],
        "Giặt Giũ & Chăm Sóc Nhà Cửa": [
            "Giặt giũ & Chăm sóc nhà cửa",
            "Giấy vệ sinh, khăn giấy",
            "Vệ sinh nhà cửa",
            "Vệ sinh bát đĩa",
            "Dụng cụ vệ sinh",
            "Chất khử mùi, làm thơm",
            "Thuốc diệt côn trùng",
            "Túi, màng bọc thực phẩm",
            "Bao bì, túi đựng rác"
        ],
        "Máy Tính & Laptop": [
            "Máy Tính Bàn",
            "Màn Hình",
            "Linh Kiện Máy Tính",
            "Thiết Bị Lưu Trữ",
            "Thiết Bị Mạng",
            "Máy In, Máy Scan & Máy Chiếu",
            "Phụ Kiện Máy Tính",
            "Laptop",
            "Khác",
            "Gaming"
        ],
        "Máy Ảnh & Máy Quay Phim": [
            "Máy ảnh - Máy quay phim",
            "Camera giám sát & Camera hệ thống",
            "Thẻ nhớ",
            "Ống kính",
            "Phụ kiện máy ảnh",
            "Máy bay camera & Phụ kiện"
        ],
        "Mẹ & Bé": [
            "Đồ dùng du lịch cho bé",
            "Đồ dùng ăn dặm cho bé",
            "Phụ kiện cho mẹ",
            "Chăm sóc sức khỏe mẹ",
            "Đồ dùng phòng tắm & Chăm sóc cơ thể bé",
            "Đồ dùng phòng ngủ cho bé",
            "An toàn cho bé",
            "Thực phẩm cho bé",
            "Chăm sóc sức khỏe bé",
            "Tã & bô em bé",
            "Đồ chơi",
            "Bộ & Gói quà tặng",
            "Khác",
            "Sữa công thức trên 24 tháng",
            "Sữa công thức 0-24 tháng tuổi"
        ],
        "Nhà Cửa & Đời Sống": [
            "Chăn, Ga, Gối & Nệm",
            "Đồ nội thất",
            "Trang trí nhà cửa",
            "Dụng cụ & Thiết bị tiện ích",
            "Đồ dùng nhà bếp và hộp đựng thực phẩm",
            "Đèn",
            "Ngoài trời & Sân vườn",
            "Đồ dùng phòng tắm",
            "Vật phẩm thờ cúng",
            "Đồ trang trí tiệc",
            "Chăm sóc nhà cửa và giặt ủi",
            "Sắp xếp nhà cửa",
            "Dụng cụ pha chế",
            "Tinh dầu thơm phòng",
            "Đồ dùng phòng ăn"
        ],
        "Nhà Sách Online": [
            "Sách Tiếng Việt",
            "Sách ngoại văn",
            "Gói Quà",
            "Bút viết",
            "Dụng cụ học sinh & văn phòng",
            "Màu, Họa Cụ và Đồ Thủ Công",
            "Sổ và Giấy Các Loại",
            "Quà Lưu Niệm",
            "Nhạc cụ và phụ kiện âm nhạc"
        ],
        "Ô Tô & Xe Máy & Xe Đạp": [
            "Xe đạp, xe điện",
            "Mô tô, xe máy",
            "Xe Ô tô",
            "Mũ bảo hiểm",
            "Phụ kiện xe máy",
            "Phụ kiện xe đạp",
            "Phụ kiện bên trong ô tô",
            "Dầu nhớt & dầu nhờn",
            "Phụ tùng ô tô",
            "Phụ tùng xe máy",
            "Phụ kiện bên ngoài ô tô",
            "Chăm sóc ô tô",
            "Dịch vụ cho xe"
        ],
        "Phụ Kiện & Trang Sức Nữ": [
            "Nhẫn",
            "Bông tai",
            "Khăn choàng",
            "Găng tay",
            "Phụ kiện tóc",
            "Vòng tay & Lắc tay",
            "Lắc chân",
            "Mũ",
            "Dây chuyền",
            "Kính mắt",
            "Kim loại quý",
            "Thắt lưng",
            "Cà vạt & Nơ cổ",
            "Phụ kiện thêm",
            "Bộ phụ kiện",
            "Khác",
            "Vớ/ Tất",
            "Ô/Dù"
        ],
        "Sắc Đẹp": [
            "Chăm sóc da mặt",
            "Tắm & chăm sóc cơ thể",
            "Trang điểm",
            "Chăm sóc tóc",
            "Dụng cụ & Phụ kiện Làm đẹp",
            "Vệ sinh răng miệng",
            "Nước hoa",
            "Chăm sóc nam giới",
            "Khác",
            "Chăm sóc phụ nữ",
            "Bộ sản phẩm làm đẹp"
        ],
        "Sức Khỏe": [
            "Vật tư y tế",
            "Chống muỗi & xua đuổi côn trùng",
            "Thực phẩm chức năng",
            "Tã người lớn",
            "Hỗ trợ làm đẹp",
            "Dụng cụ massage và trị liệu",
            "Khác"
        ],
        "Thiết Bị Điện Gia Dụng": [
            "Đồ gia dụng nhà bếp",
            "Đồ gia dụng lớn",
            "Máy hút bụi & Thiết bị làm sạch",
            "Quạt & Máy nóng lạnh",
            "Thiết bị chăm sóc quần áo",
            "Khác",
            "Máy xay, ép, máy đánh trứng trộn bột, máy xay thực phẩm",
            "Bếp điện"
        ],
        "Thiết Bị Điện Tử": [
            "Phụ kiện tivi",
            "Máy Game Console",
            "Phụ kiện Console",
            "Đĩa game",
            "Linh phụ kiện",
            "Tai nghe nhét tai",
            "Loa",
            "Tivi",
            "Tivi Box",
            "Headphones"
        ],
        "Thể Thao & Du Lịch": [
            "Vali",
            "Túi du lịch",
            "Phụ kiện du lịch",
            "Dụng Cụ Thể Thao & Dã Ngoại",
            "Giày Thể Thao",
            "Thời Trang Thể Thao & Dã Ngoại",
            "Phụ Kiện Thể Thao & Dã Ngoại",
            "Khác"
        ],
        "Thời Trang Nam": [
            "Áo Khoác",
            "Áo Vest và Blazer",
            "Áo Hoodie, Áo Len & Áo Nỉ",
            "Quần Jeans",
            "Quần Dài/Quần Âu",
            "Quần Short",
            "Áo",
            "Áo Ba Lỗ",
            "Đồ Lót",
            "Đồ Ngủ",
            "Đồ Bộ",
            "Vớ/Tất",
            "Trang Phục Truyền Thống",
            "Đồ Hóa Trang",
            "Trang Phục Ngành Nghề",
            "Khác",
            "Trang Sức Nam",
            "Kính Mắt Nam",
            "Thắt Lưng Nam",
            "Cà vạt & Nơ cổ",
            "Phụ Kiện Nam"
        ],
        "Thời Trang Nữ": [
            "Quần",
            "Quần đùi",
            "Chân váy",
            "Quần jeans",
            "Đầm/Váy",
            "Váy cưới",
            "Đồ liền thân",
            "Áo khoác, Áo choàng & Vest",
            "Áo len & Cardigan",
            "Hoodie và Áo nỉ",
            "Bộ",
            "Đồ lót",
            "Đồ ngủ",
            "Áo",
            "Đồ tập",
            "Đồ Bầu",
            "Đồ truyền thống",
            "Đồ hóa trang",
            "Vải",
            "Vớ/ Tất",
            "Khác"
        ],
        "Thời Trang Trẻ Em": [
            "Trang phục bé trai",
            "Trang phục bé gái",
            "Giày dép bé trai",
            "Giày dép bé gái",
            "Khác",
            "Quần áo em bé",
            "Giày tập đi & Tất sơ sinh",
            "Phụ kiện trẻ em"
        ],
        "Túi Ví Nữ": [
            "Ba Lô Nữ",
            "Cặp Laptop",
            "Ví Dự Tiệc & Ví Cầm Tay",
            "Túi Đeo Hông & Túi Đeo Ngực",
            "Túi Tote",
            "Túi Quai Xách",
            "Túi Đeo Chéo & Túi Đeo Vai",
            "Ví/Bóp Nữ",
            "Phụ Kiện Túi",
            "Khác"
        ],
        "Voucher & Dịch Vụ": [
            "Nhà hàng & Ăn uống",
            "Sự kiện & Giải trí",
            "Nạp tiền tài khoản",
            "Sức khỏe & Làm đẹp",
            "Gọi xe",
            "Khóa học",
            "Du lịch & Khách sạn",
            "Mua sắm",
            "Mã quà tặng Shopee",
            "Thanh toán hóa đơn",
            "Dịch vụ khác"
        ]
    };

    function getActiveSubCategoryText() {
        const xpath = '//*[@class="shopee-category-list__sub-category shopee-category-list__sub-category--active"]';
        let el = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
        if (!el) {
            el = document.querySelector('.shopee-category-list__sub-category--active');
        }
        return el ? el.textContent.trim() : null;
    }

    function clickSubCategory(subName) {
        let el = null;
        const allElements = document.querySelectorAll('.shopee-category-list__sub-category, a, span');
        for (const item of allElements) {
            if (item.textContent.trim() === subName) {
                el = item;
                break;
            }
        }
        if (el) {
            const aTag = el.closest('a') || (el.tagName.toLowerCase() === 'a' ? el : null);
            if (aTag && aTag.href) {
                addLog(`Tìm thấy liên kết chuyển danh mục: ${aTag.href}. Tiến hành chuyển hướng...`, 'info');
                window.location.href = aTag.href;
                return true;
            }
            el.click();
            return true;
        }
        return false;
    }

    // Phân tách logic chạy: affiliate.shopee.vn (Auto Hook) vs shopee.vn (Scraper)
    const isAffiliate = window.location.hostname.includes('affiliate.shopee.vn');

    if (isAffiliate) {
        try {
            runAffiliateAutoHook();
        } catch (err) {
            console.error('[Shopee VN Scraper] Lỗi khởi chạy Auto Hook:', err);
        }
        return;
    }

    // --- LOGIC CHẠY TRÊN TRANG SHOPEE.VN (CÀO SẢN PHẨM) ---

    // Khởi tạo bộ nhớ cache cho các item đã cào an toàn
    let crawledCache = GM_getValue('crawled_items_cache_vn');
    if (!Array.isArray(crawledCache)) {
        crawledCache = [];
    }
    let crawledData = []; // Dữ liệu của lượt cào hiện tại
    let isRunning = false;

    // Load các cấu hình Google Sheets và bộ lọc đã lưu (phòng thủ fallback tránh null/undefined)
    let sheetConfigRaw = GM_getValue('sheet_config_vn') || {};
    let sheetConfig = {
        sheetUrl: sheetConfigRaw.sheetUrl || '',
        tabName: sheetConfigRaw.tabName || 'Sheet1',
        clientId: sheetConfigRaw.clientId || '',
        clientSecret: sheetConfigRaw.clientSecret || '',
        priceMin: sheetConfigRaw.priceMin !== undefined ? sheetConfigRaw.priceMin : 80000,
        soldMin: sheetConfigRaw.soldMin !== undefined ? sheetConfigRaw.soldMin : 10,
        sellerCommissionMin: sheetConfigRaw.sellerCommissionMin !== undefined ? sheetConfigRaw.sellerCommissionMin : 5000,
        selectedCategory: sheetConfigRaw.selectedCategory || 'no-category'
    };

    function saveRunningState(step) {
        const state = {
            isRunning: isRunning,
            crawledData: crawledData,
            currentStep: step,
            selectedCategory: sheetConfig.selectedCategory
        };
        GM_setValue('scraper_running_state_vn', state);
    }

    function checkAndResumeState() {
        try {
            const savedState = GM_getValue('scraper_running_state_vn');
            if (savedState && savedState.isRunning) {
                addLog('Phát hiện phiên cào trước đó đang chạy dở dang. Đang tự động khôi phục...', 'warn');
                isRunning = true;
                crawledData = savedState.crawledData || [];
                
                // Đồng bộ trạng thái UI
                const startBtn = document.getElementById('scraper-btn-start');
                const stopBtn = document.getElementById('scraper-btn-stop');
                if (startBtn) {
                    startBtn.disabled = true;
                    startBtn.style.backgroundColor = '#555';
                }
                if (stopBtn) {
                    stopBtn.disabled = false;
                    stopBtn.style.backgroundColor = '#ee4d2d';
                }
                setStatus('Resuming...', '#ff9800');

                // Đợi 3 giây để trang Shopee VN tải ổn định rồi tiếp tục cào
                setTimeout(async () => {
                    if (savedState.currentStep === 'waiting_topsales') {
                        addLog('Tiếp tục bước: Chuyển sang bộ lọc Bán chạy...', 'info');
                        await transitionToTopSalesAndCrawl();
                    } else {
                        addLog('Tiếp tục cào dữ liệu sản phẩm...', 'info');
                        await crawlCurrentPage();
                    }
                }, 3000);
            }
        } catch (e) {
            console.error('Lỗi khi khôi phục trạng thái chạy:', e);
        }
    }

    // Tạo giao diện Panel điều khiển nổi (Floating Panel)
    const panel = document.createElement('div');
    panel.id = 'shopee-scraper-panel-vn';
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
    panel.style.width = '290px';
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
            console.log(`[Shopee VN Scraper] ${msg}`);
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
        try {
            if (document.getElementById('shopee-scraper-panel-vn')) return;

            // Đọc lại cấu hình Security Headers an toàn
            let apiHeadersConfigRaw = GM_getValue('api_headers_config_vn') || {};
            let tokenVal = apiHeadersConfigRaw.afAcEncSzToken || '';
            let datVal = apiHeadersConfigRaw.afAcEncDat || '';
            let secVal = apiHeadersConfigRaw.xSapSec || '';
            let cookieVal = apiHeadersConfigRaw.cookie || '';
            let sdkVal = apiHeadersConfigRaw.xSzSdkVersion || '1.12.21';

            panel.innerHTML = `
                <div style="font-weight: bold; font-size: 15px; margin-bottom: 12px; color: #ee4d2d; border-bottom: 1px solid #333; padding-bottom: 5px; display: flex; justify-content: space-between; align-items: center;">
                    <span>Shopee VN Scraper Pro v1.6</span>
                    <span id="scraper-status" style="font-size: 10px; padding: 2px 6px; background-color: #333; border-radius: 4px; color: #aaa;">Idle</span>
                </div>
                
                <!-- Bộ lọc cơ bản -->
                <div style="margin-bottom: 10px; display: flex; gap: 5px; flex-wrap: wrap;">
                    <div style="flex: 1; min-width: 80px;">
                        <label style="display: block; margin-bottom: 4px; font-weight: bold; font-size: 11px;">Price Min (₫):</label>
                        <input type="number" id="scraper-price-min" value="${sheetConfig.priceMin}" style="width: 85%; padding: 5px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white; font-size: 12px;">
                    </div>
                    <div style="flex: 1; min-width: 80px;">
                        <label style="display: block; margin-bottom: 4px; font-weight: bold; font-size: 11px;">Sold Min:</label>
                        <input type="number" id="scraper-sold-min" value="${sheetConfig.soldMin}" style="width: 85%; padding: 5px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white; font-size: 12px;">
                    </div>
                    <div style="flex: 1; min-width: 90px;">
                        <label style="display: block; margin-bottom: 4px; font-weight: bold; font-size: 11px;">Seller Com Min (₫):</label>
                        <input type="number" id="scraper-com-min" value="${sheetConfig.sellerCommissionMin}" style="width: 85%; padding: 5px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white; font-size: 12px;">
                    </div>
                </div>

                <!-- Chọn Danh Mục Cào Tuần Tự -->
                <div style="margin-bottom: 10px;">
                    <label style="display: block; margin-bottom: 4px; font-weight: bold; font-size: 11px;">Chế độ cào danh mục:</label>
                    <select id="scraper-category-select" style="width: 98%; padding: 5px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white; font-size: 12px;">
                        <option value="no-category" ${sheetConfig.selectedCategory === 'no-category' ? 'selected' : ''}>No Category (Chạy trang hiện tại)</option>
                        ${Object.keys(CATEGORIES_DATA).map(cat => `<option value="${cat}" ${sheetConfig.selectedCategory === cat ? 'selected' : ''}>${cat}</option>`).join('')}
                    </select>
                </div>
                
                <!-- Accordion cấu hình Google Sheets -->
                <div style="margin-bottom: 10px; border: 1px solid #333; border-radius: 4px;">
                    <div id="scraper-sheet-toggle" style="background-color: #2b2b2b; padding: 6px 10px; font-weight: bold; cursor: pointer; display: flex; justify-content: space-between; align-items: center; border-radius: 4px 4px 0 0;">
                        <span>⚙️ Cấu hình Google Sheets</span>
                        <span id="scraper-toggle-arrow">▲</span>
                    </div>
                    <div id="scraper-sheet-fields" style="padding: 10px; display: block; background-color: #1f1f1f;">
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

                <!-- Accordion xem cấu hình API Security Headers (Bắt tự động) -->
                <div style="margin-bottom: 10px; border: 1px solid #333; border-radius: 4px;">
                    <div id="scraper-api-toggle" style="background-color: #2b2b2b; padding: 6px 10px; font-weight: bold; cursor: pointer; display: flex; justify-content: space-between; align-items: center; border-radius: 4px 4px 0 0;">
                        <span>🛡️ API Security Headers (Auto Caught)</span>
                        <span id="scraper-api-toggle-arrow">▼</span>
                    </div>
                    <div id="scraper-api-fields" style="padding: 10px; display: none; background-color: #1f1f1f;">
                        <div style="font-size: 11px; color: #4caf50; font-weight: bold; margin-bottom: 8px; border-bottom: 1px dashed #333; padding-bottom: 4px;">
                            Tự động cập nhật khi bạn truy cập affiliate.shopee.vn
                        </div>
                        <div style="margin-bottom: 8px;">
                            <label style="display: block; margin-bottom: 3px; font-size: 10px; color: #888;">af-ac-enc-sz-token:</label>
                            <input type="text" id="scraper-api-token" readonly value="${tokenVal ? '✓ Đã lưu (Độ dài: ' + tokenVal.length + ' ký tự)' : 'Chưa có'}" style="width: 95%; padding: 4px; border-radius: 4px; border: 1px solid #333; background-color: #1a1a1a; color: #4caf50; font-size: 11px;">
                        </div>
                        <div style="margin-bottom: 8px;">
                            <label style="display: block; margin-bottom: 3px; font-size: 10px; color: #888;">af-ac-enc-dat:</label>
                            <input type="text" id="scraper-api-dat" readonly value="${datVal ? '✓ Đã lưu: ' + datVal : 'Chưa có'}" style="width: 95%; padding: 4px; border-radius: 4px; border: 1px solid #333; background-color: #1a1a1a; color: #4caf50; font-size: 11px;">
                        </div>
                        <div style="margin-bottom: 8px;">
                            <label style="display: block; margin-bottom: 3px; font-size: 10px; color: #888;">x-sap-sec:</label>
                            <input type="text" id="scraper-api-sec" readonly value="${secVal ? '✓ Đã lưu' : 'Chưa có'}" style="width: 95%; padding: 4px; border-radius: 4px; border: 1px solid #333; background-color: #1a1a1a; color: #4caf50; font-size: 11px;">
                        </div>
                        <div style="display: flex; justify-content: space-between; align-items: center; font-size: 11px; margin-bottom: 5px;">
                            <span style="color: #aaa;">Đồng bộ Cookie:</span>
                            <span style="color: ${cookieVal ? '#4caf50' : '#f44336'}; font-weight: bold;">${cookieVal ? 'ĐÃ ĐỒNG BỘ' : 'CHƯA CÓ'}</span>
                        </div>
                        <div style="display: flex; justify-content: space-between; align-items: center; font-size: 11px;">
                            <span style="color: #aaa;">SDK Version:</span>
                            <span style="color: #4caf50; font-weight: bold;">${sdkVal}</span>
                        </div>
                    </div>
                </div>

                <div style="margin-bottom: 12px; display: flex; gap: 5px; justify-content: space-between; align-items: center; background-color: #2b2b2b; padding: 6px; border-radius: 4px;">
                    <span>Đã cào: <strong id="scraper-count" style="color: #ee4d2d;">0</strong> SP</span>
                    <span>Bộ nhớ đệm: <strong id="scraper-cache-count" style="color: #4caf50;">0</strong></span>
                </div>
                
                <div style="display: flex; gap: 5px; margin-bottom: 8px;">
                    <button id="scraper-btn-start" style="flex: 1; padding: 8px; border: none; border-radius: 4px; background-color: #ee4d2d; color: white; font-weight: bold; cursor: pointer;">Start Crawl</button>
                    <button id="scraper-btn-stop" style="flex: 1; padding: 8px; border: none; border-radius: 4px; background-color: #555; color: white; font-weight: bold; cursor: pointer;" disabled>Stop</button>
                </div>
                <div style="display: flex; gap: 5px; margin-bottom: 8px;">
                    <button id="scraper-btn-clear-cache" style="width: 33%; padding: 6px; border: none; border-radius: 4px; background-color: #444; color: #ccc; cursor: pointer; font-size: 11px;">Xóa Cache</button>
                    <button id="scraper-btn-sync-cache" style="width: 34%; padding: 6px; border: none; border-radius: 4px; background-color: #0288d1; color: white; cursor: pointer; font-size: 11px; font-weight: bold;">Đồng bộ Cache</button>
                    <button id="scraper-btn-export" style="width: 33%; padding: 6px; border: none; border-radius: 4px; background-color: #2e7d32; color: white; cursor: pointer; font-size: 11px; font-weight: bold;">Xuất Excel</button>
                </div>
                <div style="margin-top: 10px;">
                    <div style="font-weight: bold; margin-bottom: 4px;">System Logs:</div>
                    <div id="scraper-logs" style="height: 120px; overflow-y: auto; background-color: #0c0c0c; border: 1px solid #333; padding: 6px; font-family: monospace; font-size: 11px; border-radius: 4px; word-break: break-all;">
                        <div style="color: #888;">Chờ lệnh Start...</div>
                    </div>
                </div>
            `;
            document.body.appendChild(panel);

            // Accordion toggle logic Google Sheets
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

            // Accordion toggle logic API Headers
            document.getElementById('scraper-api-toggle').addEventListener('click', () => {
                const fields = document.getElementById('scraper-api-fields');
                const arrow = document.getElementById('scraper-api-toggle-arrow');
                if (fields.style.display === 'none') {
                    fields.style.display = 'block';
                    arrow.innerText = '▲';
                } else {
                    fields.style.display = 'none';
                    arrow.innerText = '▼';
                }
            });

            // Nút kết nối OAuth Google
            document.getElementById('scraper-btn-connect-google').addEventListener('click', () => {
                sheetConfig.sheetUrl = document.getElementById('scraper-sheet-url').value.trim();
                sheetConfig.tabName = document.getElementById('scraper-sheet-tab').value.trim();
                sheetConfig.clientId = document.getElementById('scraper-client-id').value.trim();
                sheetConfig.clientSecret = document.getElementById('scraper-client-secret').value.trim();
                GM_setValue('sheet_config_vn', sheetConfig);

                startGoogleAuthFlow();
                document.getElementById('scraper-auth-code-wrapper').style.display = 'block';
            });

            // Xác nhận mã auth code
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

            // Test Connect
            document.getElementById('scraper-btn-test-sheet').addEventListener('click', async () => {
                sheetConfig.sheetUrl = document.getElementById('scraper-sheet-url').value.trim();
                sheetConfig.tabName = document.getElementById('scraper-sheet-tab').value.trim();
                sheetConfig.clientId = document.getElementById('scraper-client-id').value.trim();
                sheetConfig.clientSecret = document.getElementById('scraper-client-secret').value.trim();
                GM_setValue('sheet_config_vn', sheetConfig);

                await testGoogleSheetsConnection();
            });

            // Lưu cấu hình Google Sheets & Bộ lọc
            document.getElementById('scraper-btn-save-sheet').addEventListener('click', () => {
                sheetConfig.sheetUrl = document.getElementById('scraper-sheet-url').value.trim();
                sheetConfig.tabName = document.getElementById('scraper-sheet-tab').value.trim();
                sheetConfig.clientId = document.getElementById('scraper-client-id').value.trim();
                sheetConfig.clientSecret = document.getElementById('scraper-client-secret').value.trim();
                sheetConfig.priceMin = parseInt(document.getElementById('scraper-price-min').value) || 0;
                sheetConfig.soldMin = parseInt(document.getElementById('scraper-sold-min').value) || 0;
                sheetConfig.sellerCommissionMin = parseFloat(document.getElementById('scraper-com-min').value) || 0;
                sheetConfig.selectedCategory = document.getElementById('scraper-category-select').value;

                GM_setValue('sheet_config_vn', sheetConfig);
                addLog('Đã lưu cấu hình Google Sheets & Bộ lọc thành công!', 'success');
                alert('Đã lưu thông tin cấu hình và bộ lọc!');
            });

            // Lắng nghe sự kiện click các nút trên giao diện
            document.getElementById('scraper-btn-start').addEventListener('click', startScraper);
            document.getElementById('scraper-btn-stop').addEventListener('click', stopScraper);
            document.getElementById('scraper-btn-sync-cache').addEventListener('click', syncCacheFromGoogleSheets);

            document.getElementById('scraper-btn-clear-cache').addEventListener('click', () => {
                if (confirm('Bạn có chắc chắn muốn xóa bộ nhớ đệm cache các sản phẩm đã cào ở Việt Nam không?')) {
                    crawledCache = [];
                    GM_setValue('crawled_items_cache_vn', crawledCache);
                    updateCacheCount();
                    addLog('Đã xóa sạch bộ nhớ đệm cache VN', 'warn');
                    alert('Đã xóa cache thành công!');
                }
            });

            document.getElementById('scraper-btn-export').addEventListener('click', exportToExcel);

            // Tự động lưu dropdown danh mục khi thay đổi để chống reset khi reload trang
            const categorySelect = document.getElementById('scraper-category-select');
            if (categorySelect) {
                categorySelect.addEventListener('change', (e) => {
                    sheetConfig.selectedCategory = e.target.value;
                    GM_setValue('sheet_config_vn', sheetConfig);
                    addLog(`Đã chuyển chế độ cào danh mục sang: ${e.target.value === 'no-category' ? 'No Category' : e.target.value}`, 'info');
                });
            }

            updateCacheCount();

            // Tự động kiểm tra và khôi phục trạng thái chạy dở dang
            checkAndResumeState();
        } catch (err) {
            alert("Lỗi khởi tạo Panel: " + err.stack);
            console.error(err);
        }
    }

    // Khởi chạy panel khi trang đã sẵn sàng
    if (document.readyState === 'complete' || document.readyState === 'interactive') {
        initPanel();
    } else {
        window.addEventListener('DOMContentLoaded', initPanel);
    }

    // Hỗ trợ SPA chuyển trang
    const observer = new MutationObserver(() => {
        if (!document.getElementById('shopee-scraper-panel-vn') && document.body) {
            initPanel();
        }
    });
    observer.observe(document.documentElement, { childList: true, subtree: true });

    // Tiện ích
    function sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    // Tối ưu hóa xử lý số và lượt bán tại Việt Nam
    function cleanNumber(str) {
        if (!str) return 0;
        let s = str.trim().toLowerCase();

        // 1. Xử lý Lượt bán (ví dụ: "4k+ sold", "đã bán 4,2k", "đã bán 12.3k", "đã bán 12")
        const kMatch = s.match(/([\d.,]+)\s*k/);
        if (kMatch) {
            let numStr = kMatch[1].replace(/,/g, '.'); // Chuyển dấu phẩy thành dấu chấm cho parseFloat
            let num = parseFloat(numStr);
            return isNaN(num) ? 0 : Math.round(num * 1000);
        }

        // 2. Xử lý số thông thường (ví dụ: "đã bán 150", hoặc giá tiền "669.000₫")
        s = s.replace(/[₫đđ]/g, '');

        const numMatch = s.match(/[\d.,]+/g);
        if (!numMatch) return 0;

        let longestNumStr = numMatch.reduce((a, b) => a.length > b.length ? a : b, '');

        if (longestNumStr.includes('.') || longestNumStr.includes(',')) {
            const dots = (longestNumStr.match(/\./g) || []).length;
            const commas = (longestNumStr.match(/,/g) || []).length;

            if (dots > 0) {
                longestNumStr = longestNumStr.replace(/\./g, '');
            } else if (commas > 0) {
                longestNumStr = longestNumStr.replace(/,/g, '');
            }
        }

        let parsed = parseInt(longestNumStr);
        return isNaN(parsed) ? 0 : parsed;
    }

    // Lọc lấy ShopId và ItemId từ Url
    function parseProductUrl(href) {
        if (!href) return null;
        let decodedHref = href;
        try {
            decodedHref = decodeURIComponent(href);
        } catch (e) { }

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

    // Lấy Category từ Url hiện tại
    function getCategoryFromUrl() {
        try {
            const url = window.location.href;
            const match = url.match(/\/([a-zA-Z0-9-_]+)-cat\.\d+/);
            if (match && match[1]) {
                return match[1].replace(/-/g, ' ');
            }
        } catch (e) {
            console.error('Lỗi phân tích category từ URL:', e);
        }
        return '';
    }

    // Định dạng timestamp (giây) thành DD/MM/YYYY
    function formatDate(timestamp) {
        if (!timestamp) return '';
        try {
            const date = new Date(timestamp * 1000);
            const day = String(date.getDate()).padStart(2, '0');
            const month = String(date.getMonth() + 1).padStart(2, '0');
            const year = date.getFullYear();
            return `${day}/${month}/${year}`;
        } catch (e) {
            return '';
        }
    }

    // Cuộn trang tự động để load Lazy Items
    async function smoothScrollToBottom() {
        return new Promise((resolve) => {
            let currentPosition = 0;
            const distance = 400;
            const interval = 200;
            let scrollCount = 0;
            const maxScrolls = 10;

            window.scrollTo(0, 0);

            let timer = setInterval(() => {
                currentPosition += distance;
                window.scrollTo(0, currentPosition);
                scrollCount++;

                let scrollHeight = document.documentElement.scrollHeight;
                if (scrollCount >= maxScrolls || currentPosition >= scrollHeight) {
                    clearInterval(timer);
                    resolve();
                }
            }, interval);
        });
    }

    // Gọi API Shopee Affiliate lấy dữ liệu hoa hồng thông qua Bridge tab affiliate.shopee.vn
    function fetchCommissionData(itemId, retries = 3) {
        return new Promise((resolve) => {
            const requestId = Math.random().toString(36).substring(2, 15);
            let checkTimeout = null;
            let listenerId = null;

            function attemptFetch(attempt) {
                listenerId = GM_addValueChangeListener('scraper_response_vn', (name, oldValue, newValue, remote) => {
                    if (newValue && newValue.requestId === requestId) {
                        if (listenerId) {
                            GM_removeValueChangeListener(listenerId);
                            listenerId = null;
                        }
                        clearTimeout(checkTimeout);

                        if (newValue.status === 200) {
                            try {
                                const res = JSON.parse(newValue.responseText);
                                if (res.code === 0 && res.data) {
                                    const commissionStr = res.data.commission || "₫0";
                                    const sellerComStr = (res.data.commission_rate && res.data.commission_rate.seller_commission) || "₫0";
                                    const shopeeComStr = (res.data.commission_rate && res.data.commission_rate.shopee_commission) || "₫0";
                                    const similarOffers = (res.data.similar_product_offers && res.data.similar_product_offers.list) || [];
                                    const rootImages = (res.data.batch_item_for_item_card_full && res.data.batch_item_for_item_card_full.images) || [];

                                    const card = res.data.batch_item_for_item_card_full;
                                    const stock = card ? card.stock : 0;
                                    const ratingStar = (card && card.item_rating) ? card.item_rating.rating_star : 0;
                                    const ratingCountArr = (card && card.item_rating) ? card.item_rating.rating_count : [];
                                    const ratingCount = (Array.isArray(ratingCountArr) && ratingCountArr.length > 0) ? Math.max(...ratingCountArr) : 0;
                                    const ctime = card ? card.ctime : 0;

                                    resolve({
                                        commission: cleanNumber(commissionStr),
                                        sellerComFinal: cleanNumber(sellerComStr),
                                        shopeeComFinal: cleanNumber(shopeeComStr),
                                        similarOffers: similarOffers,
                                        images: rootImages,
                                        stock: stock,
                                        ratingStar: Math.round(ratingStar * 100) / 100,
                                        ratingCount: ratingCount,
                                        ctime: ctime,
                                        success: true
                                    });
                                    return;
                                } else {
                                    const errCode = res.code !== undefined ? res.code : (res.error !== undefined ? res.error : 'unknown');
                                    const errMsg = res.msg || res.message || JSON.stringify(res);
                                    handleError(`API trả về lỗi (code ${errCode}): ${errMsg}`);
                                }
                            } catch (e) {
                                handleError(`Lỗi phân tích JSON: ${e.message}`);
                            }
                        } else {
                            handleError(`HTTP Status ${newValue.status} | Lỗi: ${newValue.error || 'Yêu cầu qua Bridge thất bại'}`);
                        }
                    }
                });

                // Gửi yêu cầu qua Bridge
                GM_setValue('scraper_request_vn', { itemId, requestId, timestamp: Date.now() });

                // Hạn giờ phản hồi sau 8 giây
                checkTimeout = setTimeout(() => {
                    if (listenerId) {
                        GM_removeValueChangeListener(listenerId);
                        listenerId = null;
                    }
                    handleError('Không nhận được phản hồi từ tab Bridge (Timeout)');
                }, 8000);

                async function handleError(errorMsg) {
                    if (attempt < retries) {
                        addLog(`[Thử lại ${attempt}/${retries}] Lỗi lấy hoa hồng SP ${itemId}: ${errorMsg}. Đang thử lại sau 2 giây...`, 'warn');
                        await sleep(2000);
                        attemptFetch(attempt + 1);
                    } else {
                        addLog(`[Lỗi] Thất bại khi lấy hoa hồng cho SP ${itemId} sau ${retries} lần thử: ${errorMsg}. Điền mặc định 0.`, 'error');
                        resolve({ commission: 0, sellerComFinal: 0, shopeeComFinal: 0, similarOffers: [], images: [], stock: 0, ratingStar: 0, ratingCount: 0, ctime: 0, success: false });
                    }
                }
            }

            attemptFetch(1);
        });
    }

    // Bổ sung thông tin hoa hồng cho các sản phẩm đã cào
    async function enrichCrawledDataWithCommission() {
        try {
            if (crawledData.length === 0) return;

            addLog(`Đang kết nối Bridge tới tab affiliate.shopee.vn...`, 'info');
            setStatus('Checking Bridge...', '#ff9800');

            // Gửi tín hiệu ping xem Bridge đã hoạt động chưa
            let bridgeActive = false;
            const pingVal = Math.random().toString();
            const pongListener = GM_addValueChangeListener('scraper_pong_vn', (name, oldValue, newValue, remote) => {
                bridgeActive = true;
            });

            GM_setValue('scraper_ping_vn', pingVal);
            await sleep(1200); // Đợi phản hồi trong 1.2s
            GM_removeValueChangeListener(pongListener);

            if (!bridgeActive) {
                addLog('Không tìm thấy tab affiliate.shopee.vn đang hoạt động. Tự động mở tab mới làm Bridge...', 'warn');
                window.open('https://affiliate.shopee.vn/', '_blank');
                addLog('Vui lòng chờ 6 giây để tab Bridge tải và sẵn sàng...', 'info');
                await sleep(6000);
            } else {
                addLog('Kết nối Bridge thành công! Bắt đầu gọi API thông qua Bridge...', 'success');
            }

            addLog(`Bắt đầu gọi API Affiliate lấy hoa hồng cho ${crawledData.length} sản phẩm...`, 'info');
            setStatus('Enriching API...', '#009688');

            const batchSize = 3;
            const delayBetweenBatches = 2000;

            const finalFilteredData = [];
            const sellerComMin = parseFloat(document.getElementById('scraper-com-min').value) || 0;

            for (let i = 0; i < crawledData.length; i += batchSize) {
                if (!isRunning) break;

                const batch = crawledData.slice(i, i + batchSize);
                addLog(`Đang gọi API hoa hồng cho nhóm SP từ ${i + 1} đến ${Math.min(i + batchSize, crawledData.length)}...`, 'info');

                const promises = batch.map(async (item) => {
                    const itemId = item['Mã SP (Item ID)'];
                    if (itemId) {
                        const info = await fetchCommissionData(itemId);
                        item['Hoa hồng (₫)'] = info.commission;
                        item['Hoa hồng người bán (₫)'] = info.sellerComFinal;
                        item['Hoa hồng Shopee (₫)'] = info.shopeeComFinal;
                        item['Nhóm sản phẩm (ID-Group)'] = `gr-${itemId}`;

                        const rootImages = (info.images || []).map(imgId => `https://down-sg.img.susercontent.com/file/${imgId}`).join('|');
                        item['Danh sách ảnh'] = rootImages;

                        item['Category'] = getCategoryFromUrl();
                        item['Tồn kho'] = info.stock;
                        item['Đánh giá sao'] = info.ratingStar;
                        item['Lượt đánh giá'] = info.ratingCount;
                        item['Ngày đăng bán'] = formatDate(info.ctime);

                        if (info.sellerComFinal >= sellerComMin) {
                            finalFilteredData.push(item);

                            // Lấy sản phẩm tương tự
                            if (info.similarOffers && info.similarOffers.length > 0) {
                                const parsedSimilar = info.similarOffers.map(sim => {
                                    const card = sim.batch_item_for_item_card_full;
                                    if (!card) return null;

                                    const priceVal = parseFloat(card.price) / 100000;
                                    const sellerRate = parseFloat(sim.seller_commission_rate || "0%") || 0;
                                    const defaultRate = parseFloat(sim.default_commission_rate || "0%") || 0;

                                    const sellerCom = Math.round((priceVal * (sellerRate / 100)) * 100) / 100;
                                    const totalCom = Math.round((priceVal * (defaultRate / 100)) * 100) / 100;
                                    const shopeeCom = Math.max(0, Math.round((totalCom - sellerCom) * 100) / 100);

                                    const simItemId = sim.item_id || card.itemid;
                                    const simShopId = card.shopid;
                                    const simUrl = sim.product_link || `https://shopee.vn/product/${simShopId}/${simItemId}`;
                                    const simImages = (card.images || []).map(imgId => `https://down-sg.img.susercontent.com/file/${imgId}`).join('|');

                                    const simRatingStar = (card && card.item_rating) ? card.item_rating.rating_star : 0;
                                    const simRatingCountArr = (card && card.item_rating) ? card.item_rating.rating_count : [];
                                    const simRatingCount = (Array.isArray(simRatingCountArr) && simRatingCountArr.length > 0) ? Math.max(...simRatingCountArr) : 0;
                                    const simCtime = card ? card.ctime : 0;

                                    return {
                                        'Mã SP (Item ID)': simItemId,
                                        'Mã Shop (Shop ID)': simShopId,
                                        'Tên sản phẩm': card.name || "",
                                        'Giá (₫)': priceVal,
                                        'Lượt bán trong 30 ngày': card.sold || 0,
                                        'Đường dẫn sản phẩm': simUrl,
                                        'Hoa hồng (₫)': totalCom,
                                        'Hoa hồng người bán (₫)': sellerCom,
                                        'Hoa hồng Shopee (₫)': shopeeCom,
                                        'Nhóm sản phẩm (ID-Group)': `gr-${itemId}`,
                                        'Danh sách ảnh': simImages,
                                        'Category': getCategoryFromUrl(),
                                        'Tồn kho': card.stock || 0,
                                        'Đánh giá sao': Math.round(simRatingStar * 100) / 100,
                                        'Lượt đánh giá': simRatingCount,
                                        'Ngày đăng bán': formatDate(simCtime)
                                    };
                                }).filter(x => x !== null);

                                // Lọc theo sellerComMin và kiểm tra cache tránh cào lại sản phẩm tương tự đã có
                                const filteredSimilar = parsedSimilar.filter(sim => 
                                    sim['Hoa hồng người bán (₫)'] >= sellerComMin && 
                                    !crawledCache.includes(sim['Mã SP (Item ID)'])
                                );

                                // Sắp xếp giảm dần theo hoa hồng người bán
                                filteredSimilar.sort((a, b) => b['Hoa hồng người bán (₫)'] - a['Hoa hồng người bán (₫)']);

                                // Lấy tối đa 5 sản phẩm tốt nhất
                                const top5Similar = filteredSimilar.slice(0, 5);

                                // Thêm vào finalFilteredData
                                finalFilteredData.push(...top5Similar);

                                // Lưu tạm thời vào cache để tránh trùng lặp trong các lô xử lý sau
                                top5Similar.forEach(sim => {
                                    const simId = sim['Mã SP (Item ID)'];
                                    if (simId && !crawledCache.includes(simId)) {
                                        crawledCache.push(simId);
                                    }
                                });

                                if (top5Similar.length > 0) {
                                    addLog(`Tìm thấy và thêm ${top5Similar.length} SP tương tự đạt chuẩn hoa hồng cho SP gốc ${itemId}`, 'success');
                                }
                            }
                        } else {
                            addLog(`Lọc bỏ SP "${item['Tên sản phẩm'].substring(0, 20)}..." do hoa hồng người bán (₫${info.sellerComFinal} < ₫${sellerComMin})`, 'warn');
                        }
                    }
                });

                await Promise.all(promises);

                if (i + batchSize < crawledData.length) {
                    await sleep(delayBetweenBatches);
                }
            }

            // Tiến hành lọc trùng hoàn toàn trên toàn bộ danh sách kết quả phiên cào hiện tại
            const uniqueFilteredData = [];
            const seenIds = new Set();
            
            for (const item of finalFilteredData) {
                const itemId = item['Mã SP (Item ID)'];
                if (!seenIds.has(itemId)) {
                    seenIds.add(itemId);
                    uniqueFilteredData.push(item);

                    if (!crawledCache.includes(itemId)) {
                        crawledCache.push(itemId);
                    }
                } else {
                    addLog(`[Lọc trùng] Bỏ qua sản phẩm trùng lặp trong phiên cào này: ${item['Tên sản phẩm'].substring(0, 20)}... (ID: ${itemId})`, 'warn');
                }
            }

            crawledData = uniqueFilteredData;
            GM_setValue('crawled_items_cache_vn', crawledCache);
            updateCacheCount();

            const countEl = document.getElementById('scraper-count');
            if (countEl) countEl.innerText = crawledData.length;
            addLog(`Hoàn thành lấy dữ liệu hoa hồng. Giữ lại ${crawledData.length} SP đạt chuẩn hoa hồng và không trùng lặp!`, 'success');
        } catch (err) {
            alert("Lỗi bổ sung thông tin hoa hồng: " + err.stack);
        }
    }

    // Hàm cào trang hiện tại
    async function crawlCurrentPage() {
        try {
            if (!isRunning) return;

            addLog('Đang cuộn chuột xuống đáy trang để tải thêm sản phẩm...', 'info');
            setStatus('Scrolling...', '#ff9800');
            await smoothScrollToBottom();
            await sleep(1500);

            setStatus('Analyzing...', '#2196f3');
            const items = document.querySelectorAll('[data-sqe="item"]');
            addLog(`Tìm thấy tổng cộng ${items.length} thẻ sản phẩm trên trang này`, 'info');

            if (items.length === 0) {
                addLog('Không tìm thấy sản phẩm nào! Hãy đảm bảo bạn đang ở trang tìm kiếm hoặc danh mục của shopee.vn.', 'error');
                
                // TỰ ĐỘNG RELOAD TRANG NẾU KHÔNG CÓ SẢN PHẨM NÀO Ở BƯỚC TOP SALES (BÁN CHẠY)
                const savedState = GM_getValue('scraper_running_state_vn');
                if (savedState && savedState.currentStep === 'waiting_topsales') {
                    addLog('Phát hiện trang không load được sản phẩm ở bộ lọc Bán chạy. Đang tự động reload trang để tải lại...', 'warn');
                    await sleep(1500);
                    window.location.reload();
                    return;
                }
            }

            let countThisPage = 0;
            const priceMin = parseInt(document.getElementById('scraper-price-min').value) || 0;
            const soldMin = parseInt(document.getElementById('scraper-sold-min').value) || 0;

            for (let i = 0; i < items.length; i++) {
                if (!isRunning) return;
                const item = items[i];

                try {
                    const aTag = item.querySelector('a.contents, a');
                    if (!aTag) continue;

                    const hrefAttr = aTag.getAttribute('href');
                    const parsed = parseProductUrl(hrefAttr);
                    if (!parsed) continue;

                    let title = "";
                    const cardWrapper = item.querySelector('[aria-label^="Product card:"]');
                    if (cardWrapper) {
                        title = cardWrapper.getAttribute('aria-label').replace("Product card:", "").trim();
                    }
                    if (!title) {
                        const titleNode = item.querySelector('div.whitespace-normal.line-clamp-2');
                        if (titleNode) title = titleNode.textContent.trim();
                    }
                    if (!title) {
                        const mainImg = item.querySelector('img[alt]:not([alt="promotion-label"]):not([alt="custom-overlay"]):not([alt="flag-label"])');
                        if (mainImg) title = mainImg.getAttribute('alt') || "";
                    }

                    const shortTitle = title.length > 25 ? title.substring(0, 25) + '...' : title;

                    if (crawledCache.includes(parsed.itemId)) {
                        addLog(`Bỏ qua: "${shortTitle}" (Đã cào trước đó - Trùng Cache)`, 'warn');
                        continue;
                    }

                    let priceVal = 0;
                    const promotionPriceNode = item.querySelector('[aria-label="promotion price"]');
                    if (promotionPriceNode) {
                        const parent = promotionPriceNode.parentElement;
                        if (parent) {
                            priceVal = cleanNumber(parent.textContent);
                        }
                    } else {
                        const textNodes = Array.from(item.querySelectorAll('*'));
                        const priceNode = textNodes.find(node => node.textContent.includes('₫') && node.children.length === 0);
                        if (priceNode) {
                            priceVal = cleanNumber(priceNode.textContent);
                        }
                    }

                    if (priceVal < priceMin) {
                        addLog(`Bỏ qua: "${shortTitle}" (Giá ${priceVal}₫ < ${priceMin}₫)`, 'warn');
                        continue;
                    }

                    let soldVal = 0;
                    const textNodes = Array.from(item.querySelectorAll('*'));
                    const soldNode = textNodes.find(node => {
                        const text = node.textContent.trim().toLowerCase();
                        const hasKeyword = text.includes('sold') || text.includes('đã bán');
                        return hasKeyword && /\d/.test(text) && text.length < 30 && node.children.length === 0;
                    });

                    if (soldNode) {
                        soldVal = cleanNumber(soldNode.textContent);
                    }

                    if (soldVal < soldMin) {
                        addLog(`Bỏ qua: "${shortTitle}" (Đã bán ${soldVal} < ${soldMin})`, 'warn');
                        continue;
                    }

                    crawledData.push({
                        'Mã SP (Item ID)': parsed.itemId,
                        'Mã Shop (Shop ID)': parsed.shopId,
                        'Tên sản phẩm': title,
                        'Giá (₫)': priceVal,
                        'Lượt bán trong 30 ngày': soldVal,
                        'Đường dẫn sản phẩm': parsed.productUrl,
                        'Hoa hồng (₫)': 0,
                        'Hoa hồng người bán (₫)': 0,
                        'Hoa hồng Shopee (₫)': 0,
                        'Nhóm sản phẩm (ID-Group)': `gr-${parsed.itemId}`,
                        'Danh sách ảnh': '',
                        'Category': '',
                        'Tồn kho': 0,
                        'Đánh giá sao': 0,
                        'Lượt đánh giá': 0,
                        'Ngày đăng bán': ''
                    });

                    crawledCache.push(parsed.itemId);
                    countThisPage++;
                    addLog(`ĐÃ CÀO THÀNH CÔNG: "${shortTitle}" (Giá: ${priceVal}₫ | Bán: ${soldVal})`, 'success');

                    const countEl = document.getElementById('scraper-count');
                    if (countEl) countEl.innerText = crawledData.length;
                    GM_setValue('crawled_items_cache_vn', crawledCache);
                    updateCacheCount();

                    // Lưu trạng thái chạy sau mỗi lần thêm sản phẩm thành công
                    const savedState = GM_getValue('scraper_running_state_vn');
                    const step = (savedState && savedState.currentStep) || 'crawling_popular';
                    saveRunningState(step);
                } catch (err) {
                    addLog(`Lỗi xử lý sản phẩm: ${err.message}`, 'error');
                }
            }

            addLog(`Hoàn thành cào trang này. Thu thập thêm được ${countThisPage} SP đạt chuẩn.`, 'success');

            const nextPageBtn = document.querySelector('.shopee-icon-button--right:not([aria-disabled="true"]):not([disabled])');
            if (nextPageBtn && isRunning) {
                addLog('Phát hiện nút chuyển trang. Đang chuyển sang trang tiếp theo sau 3 giây...', 'info');
                setStatus('Next Page...', '#9c27b0');
                nextPageBtn.click();
                await sleep(3500);
                await crawlCurrentPage();
            } else {
                addLog('Đã cào xong tất cả các trang của bộ lọc hiện tại.', 'success');
                
                const selectedXPath = '//*[@class="shopee-sort-by-options__option shopee-sort-by-options__option--selected"]//span[text()="Bán chạy"]';
                function getElementByXPath(xpath) {
                    return document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                }

                let isTopSalesSelected = getElementByXPath(selectedXPath);
                if (!isTopSalesSelected) {
                    // Nếu chưa cào Bán chạy, chuyển sang cào Bán chạy
                    await transitionToTopSalesAndCrawl();
                } else {
                    // Nếu đã cào xong cả Bán chạy, kết thúc cào và đồng bộ dữ liệu
                    await finishCrawlingSession();
                }
            }
        } catch (err) {
            alert("Lỗi cào sản phẩm: " + err.stack);
            stopScraper();
        }
    }

    // Hàm chuyển sang cào Top Sales (Bán chạy)
    async function transitionToTopSalesAndCrawl() {
        if (!isRunning) return;

        const selectedXPath = '//*[@class="shopee-sort-by-options__option shopee-sort-by-options__option--selected"]//span[text()="Bán chạy"]';
        const targetXPath = '//span[text()="Bán chạy"]';

        function getElementByXPath(xpath) {
            return document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
        }

        let isTopSalesSelected = getElementByXPath(selectedXPath);
        if (isTopSalesSelected) {
            await finishCrawlingSession();
            return;
        }

        addLog('Bắt đầu chuyển sang bộ lọc "Bán chạy" để tiếp tục cào...', 'info');
        const targetBtn = getElementByXPath(targetXPath);
        if (targetBtn) {
            // Lưu trạng thái trước khi click để nếu click kích hoạt reload trang, tool vẫn nhớ
            saveRunningState('waiting_topsales');
            targetBtn.click();
            addLog('Đang chờ trang Bán chạy load xong...', 'info');
            setStatus('Loading Top Sales...', '#2196f3');

            let activated = false;
            for (let attempt = 1; attempt <= 10; attempt++) {
                await sleep(1000);
                if (!isRunning) return;
                if (getElementByXPath(selectedXPath)) {
                    activated = true;
                    break;
                }
            }

            if (activated) {
                addLog('Chuyển sang bộ lọc Bán chạy thành công! Bắt đầu cào các trang của Bán chạy...', 'success');
                await sleep(1500);
                await crawlCurrentPage();
            } else {
                addLog('Không thể xác nhận bộ lọc Bán chạy đã kích hoạt (Timeout). Tiến hành gọi API với dữ liệu hiện có...', 'warn');
                await finishCrawlingSession();
            }
        } else {
            addLog('Không tìm thấy nút Bán chạy trên giao diện. Tiến hành gọi API với dữ liệu hiện có...', 'warn');
            await finishCrawlingSession();
        }
    }

    // Hàm kết thúc phiên cào: Gọi API, push Google Sheets và kiểm tra chuyển danh mục con tiếp theo
    async function finishCrawlingSession() {
        addLog('Hoàn thành cào tất cả các trang cần thiết. Bắt đầu gọi API hoa hồng...', 'info');

        await enrichCrawledDataWithCommission();

        addLog('Bắt đầu đồng bộ dữ liệu lên Google Sheets...', 'info');
        try {
            await pushToGoogleSheets();
        } catch (err) {
            addLog(`Lỗi đẩy dữ liệu lên Google Sheets: ${err.message}`, 'error');
            addLog('Tự động xuất file Excel dự phòng...', 'warn');
            exportToExcel(true);
        }

        // Kiểm tra chế độ chạy sau khi lần đầu up dữ liệu hoàn tất
        if (isRunning) {
            const selectedCat = sheetConfig.selectedCategory;
            if (selectedCat && selectedCat !== 'no-category') {
                const subCats = CATEGORIES_DATA[selectedCat];
                if (subCats && subCats.length > 0) {
                    const activeSub = getActiveSubCategoryText();
                    let nextIndex = 0;

                    if (activeSub) {
                        const idx = subCats.indexOf(activeSub);
                        if (idx !== -1) {
                            nextIndex = idx + 1;
                        } else {
                            addLog(`[Cảnh báo] Danh mục con active trên web là "${activeSub}" nhưng không khớp chính xác với bất kỳ danh mục con nào trong cấu hình của nhóm "${selectedCat}"!`, 'warn');
                            nextIndex = 0;
                        }
                    } else {
                        addLog(`Không tìm thấy danh mục con nào đang active khi bắt đầu. Mặc định cào từ danh mục con đầu tiên của "${selectedCat}"...`, 'info');
                        nextIndex = 0;
                    }

                    if (nextIndex < subCats.length) {
                        const nextSubName = subCats[nextIndex];
                        addLog(`Chuyển sang danh mục con tiếp theo: "${nextSubName}"...`, 'info');
                        setStatus(`Moving to: ${nextSubName}`, '#9c27b0');

                        // Lưu trạng thái trước khi chuyển danh mục mới
                        crawledData = [];
                        saveRunningState('crawling_popular');

                        const clicked = clickSubCategory(nextSubName);
                        if (clicked) {
                            addLog(`Click thành công danh mục con "${nextSubName}" hoặc đang chuyển hướng trang. Đợi 4 giây cho trang tải dữ liệu...`, 'success');
                            await sleep(4000);
                            await crawlCurrentPage();
                            return;
                        } else {
                            addLog(`[Lỗi] Không thể click vào danh mục con "${nextSubName}" trên giao diện! Dừng cào tự động.`, 'error');
                            setStatus('Error!', '#d32f2f');
                            stopScraper();
                            return;
                        }
                    } else {
                        addLog(`Đã hoàn thành cào tất cả các danh mục con trong "${selectedCat}"!`, 'success');
                    }
                } else {
                    addLog(`[Lỗi] Cấu hình của danh mục "${selectedCat}" rỗng hoặc không hợp lệ!`, 'error');
                }
            }
        }

        setStatus('Done!', '#4caf50');
        stopScraper();
        showDoneOverlay();
    }

    // Google Sheets OAuth
    function startGoogleAuthFlow() {
        try {
            const client_id = sheetConfig.clientId;
            if (!client_id) {
                addLog('Lỗi: Hãy điền Client ID trước khi Kết nối Google!', 'error');
                alert('Vui lòng điền Google Client ID trước!');
                return;
            }

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
        } catch (err) {
            alert("Lỗi khởi chạy Google Auth: " + err.stack);
        }
    }

    function exchangeAuthCodeForTokens(authCode) {
        try {
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
                            GM_setValue('google_oauth_token_vn', res.access_token);
                            GM_setValue('google_token_expiry_vn', Date.now() + (res.expires_in * 1000));
                            if (res.refresh_token) {
                                GM_setValue('google_refresh_token_vn', res.refresh_token);
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
        } catch (err) {
            alert("Lỗi trao đổi Authorization Code: " + err.stack);
        }
    }

    async function getGoogleAccessToken() {
        return new Promise((resolve, reject) => {
            try {
                const storedToken = GM_getValue('google_oauth_token_vn', null);
                const tokenExpiry = GM_getValue('google_token_expiry_vn', 0);

                if (storedToken && Date.now() < tokenExpiry - 300000) {
                    resolve(storedToken);
                    return;
                }

                const client_id = sheetConfig.clientId;
                const client_secret = sheetConfig.clientSecret;
                const refresh_token = GM_getValue('google_refresh_token_vn', null);

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
                                GM_setValue('google_oauth_token_vn', res.access_token);
                                GM_setValue('google_token_expiry_vn', Date.now() + (res.expires_in * 1000));
                                addLog('Tự động gia hạn Access Token thành công!', 'success');
                                resolve(res.access_token);
                            } else {
                                addLog(`Lỗi gia hạn Token (hết hạn hoặc thu hồi): ${response.responseText}`, 'error');
                                GM_setValue('google_refresh_token_vn', null);
                                reject(new Error('Refresh token invalid'));
                            }
                        } catch (e) {
                            reject(e);
                        }
                    },
                    onerror: (err) => reject(err)
                });
            } catch (err) {
                alert("Lỗi gia hạn Access Token: " + err.stack);
                reject(err);
            }
        });
    }

    async function pushToGoogleSheets() {
        return new Promise(async (resolve, reject) => {
            try {
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

                addLog('Đang kết nối xác thực tài khoản Google...', 'info');
                const token = await getGoogleAccessToken();
                addLog('Xác thực thành công. Đang ghi dữ liệu lên Google Sheet...', 'info');

                const headers = ['Mã SP (Item ID)', 'Mã Shop (Shop ID)', 'Tên sản phẩm', 'Giá (₫)', 'Lượt bán trong 30 ngày', 'Đường dẫn sản phẩm', 'Hoa hồng (₫)', 'Hoa hồng người bán (₫)', 'Hoa hồng Shopee (₫)', 'Nhóm sản phẩm (ID-Group)', 'Danh sách ảnh', 'Category', 'Tồn kho', 'Đánh giá sao', 'Lượt đánh giá', 'Ngày đăng bán'];
                const rows = crawledData.map(item => [
                    item['Mã SP (Item ID)'],
                    item['Mã Shop (Shop ID)'],
                    item['Tên sản phẩm'],
                    item['Giá (₫)'],
                    item['Lượt bán trong 30 ngày'],
                    item['Đường dẫn sản phẩm'],
                    item['Hoa hồng (₫)'] || 0,
                    item['Hoa hồng người bán (₫)'] || 0,
                    item['Hoa hồng Shopee (₫)'] || 0,
                    item['Nhóm sản phẩm (ID-Group)'] || '',
                    item['Danh sách ảnh'] || '',
                    item['Category'] || '',
                    item['Tồn kho'] || 0,
                    item['Đánh giá sao'] || 0,
                    item['Lượt đánh giá'] || 0,
                    item['Ngày đăng bán'] || ''
                ]);

                const range = `${sheetConfig.tabName}!A:P`;
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
                            if (response.status === 401) {
                                GM_setValue('google_oauth_token_vn', null);
                            }
                            reject(new Error(`Google API respond with status ${response.status}`));
                        }
                    },
                    onerror: function (err) {
                        addLog(`Lỗi kết nối API Google: ${err.message}`, 'error');
                        reject(err);
                    }
                });
            } catch (err) {
                alert("Lỗi đẩy dữ liệu Sheets: " + err.stack);
                reject(err);
            }
        });
    }

    async function syncCacheFromGoogleSheets() {
        try {
            const spreadsheetId = getSpreadsheetId(sheetConfig.sheetUrl);
            if (!spreadsheetId) {
                addLog('Lỗi đồng bộ: URL Google Sheets không đúng định dạng!', 'error');
                alert('Vui lòng cấu hình URL Spreadsheet hợp lệ trước khi đồng bộ!');
                return;
            }

            addLog('Đang kết nối xác thực Google để đồng bộ bộ nhớ đệm...', 'info');
            const token = await getGoogleAccessToken();
            addLog('Đang tải dữ liệu cột A từ Google Sheet...', 'info');

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
                            const values = data.values;

                            if (!values || values.length === 0) {
                                addLog('Đồng bộ hoàn tất: Không tìm thấy dữ liệu nào ở cột A. Cache cũ giữ nguyên.', 'warn');
                                alert('Không tìm thấy dữ liệu sản phẩm nào để đồng bộ!');
                                return;
                            }

                            const sheetIds = values
                                .map(row => row[0])
                                .filter(val => val !== undefined && val !== null)
                                .map(val => String(val).trim())
                                .filter(val => /^\d+$/.test(val));

                            if (sheetIds.length === 0) {
                                addLog('Đồng bộ hoàn tất: Không có mã sản phẩm hợp lệ ở cột A.', 'warn');
                                return;
                            }

                            const beforeCount = crawledCache.length;
                            crawledCache = Array.from(new Set([...crawledCache, ...sheetIds]));
                            const addedCount = crawledCache.length - beforeCount;

                            GM_setValue('crawled_items_cache_vn', crawledCache);
                            updateCacheCount();

                            addLog(`ĐỒNG BỘ CACHE THÀNH CÔNG!`, 'success');
                            addLog(`- Quét từ Sheet: ${sheetIds.length} mã sản phẩm.`, 'success');
                            addLog(`- Thêm mới vào Cache: ${addedCount} mã sản phẩm.`, 'success');

                            alert(`Đồng bộ bộ nhớ đệm thành công!\nĐã thêm mới: ${addedCount} mã sản phẩm.\nTổng số trong Cache: ${crawledCache.length}.`);
                        } catch (err) {
                            addLog(`Lỗi xử lý dữ liệu đồng bộ: ${err.message}`, 'error');
                        }
                    } else {
                        addLog(`Đồng bộ thất bại (Mã lỗi ${response.status}): ${response.responseText}`, 'error');
                        if (response.status === 401) {
                            GM_setValue('google_oauth_token_vn', null);
                        }
                    }
                },
                onerror: function (err) {
                    addLog(`Lỗi mạng khi kết nối đồng bộ: ${err.message}`, 'error');
                }
            });
        } catch (e) {
            alert("Lỗi đồng bộ cache từ Sheets: " + e.stack);
        }
    }

    async function testGoogleSheetsConnection() {
        try {
            const spreadsheetId = getSpreadsheetId(sheetConfig.sheetUrl);
            if (!spreadsheetId) {
                addLog('Lỗi Test Connect: URL Google Sheets không đúng định dạng!', 'error');
                alert('Vui lòng nhập link Spreadsheet hợp lệ!');
                return;
            }

            addLog('Đang chuẩn bị kiểm tra kết nối với Google API...', 'info');
            const token = await getGoogleAccessToken();
            addLog('Đang lấy Metadata của Spreadsheet...', 'info');

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
                            addLog(`- Các Sheet Tab: [${sheets.join(', ')}]`, 'info');

                            if (sheets.includes(sheetConfig.tabName)) {
                                addLog(`- Tab chỉ định "${sheetConfig.tabName}": HỢP LỆ!`, 'success');
                                alert(`Kết nối thành công!\nTên bảng tính: ${title}\nTab "${sheetConfig.tabName}" sẵn sàng.`);
                            } else {
                                addLog(`- Tab chỉ định "${sheetConfig.tabName}": KHÔNG TÌM THẤY!`, 'warn');
                                alert(`Kết nối thành công!\nTên bảng tính: ${title}\nCảnh báo: Tab "${sheetConfig.tabName}" chưa tồn tại (sẽ tự tạo khi ghi dữ liệu).`);
                            }
                        } catch (err) {
                            addLog(`Lỗi parse JSON: ${err.message}`, 'error');
                        }
                    } else {
                        addLog(`Test thất bại (Mã lỗi ${response.status}): ${response.responseText}`, 'error');
                        alert(`Kết nối thất bại!\nMã lỗi: ${response.status}`);
                        if (response.status === 401) {
                            GM_setValue('google_oauth_token_vn', null);
                        }
                    }
                },
                onerror: function (err) {
                    addLog(`Lỗi kết nối mạng khi Test: ${err.message}`, 'error');
                }
            });
        } catch (e) {
            alert("Lỗi Test Connection: " + e.stack);
        }
    }

    // Bắt đầu cào
    async function startScraper() {
        try {
            if (isRunning) return;
            isRunning = true;
            crawledData = [];

            const startBtn = document.getElementById('scraper-btn-start');
            const stopBtn = document.getElementById('scraper-btn-stop');
            if (startBtn) {
                startBtn.disabled = true;
                startBtn.style.backgroundColor = '#555';
            }
            if (stopBtn) {
                stopBtn.disabled = false;
                stopBtn.style.backgroundColor = '#ee4d2d';
            }

            // Lưu trạng thái bắt đầu cào
            saveRunningState('crawling_popular');

            await crawlCurrentPage();
        } catch (err) {
            alert("Lỗi khi bắt đầu Scraper: " + err.stack);
            console.error(err);
            stopScraper();
        }
    }

    // Dừng cào
    function stopScraper() {
        try {
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
            
            // Xóa trạng thái chạy lưu trong bộ nhớ
            GM_setValue('scraper_running_state_vn', null);
        } catch (err) {
            console.error("Lỗi khi dừng Scraper:", err);
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

    // Hàm xuất dữ liệu Excel dùng XLSX.js
    function exportToExcel(isFallback = false) {
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
        const fileName = `product_vn_${dateStr}.xlsx`;

        XLSX.writeFile(workbook, fileName);
        addLog(`Đã xuất và tải về file Excel: ${fileName}`, 'success');
        if (isFallback === true) {
            alert(`Gặp lỗi kết nối Google Sheets! Đã tự động tải file Excel dự phòng thay thế: ${fileName}`);
        } else {
            alert(`Đã xuất và tải file Excel thành công: ${fileName}`);
        }
    }

    // Trích xuất Spreadsheet ID từ URL
    function getSpreadsheetId(url) {
        if (!url) return null;
        const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
        return match ? match[1] : null;
    }

    // Hiển thị lớp phủ màu xanh lá cây nhạt kèm thông báo Done!
    function showDoneOverlay() {
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
        overlay.style.zIndex = '9999999';
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

    // --- LOGIC CHẠY TRÊN TRANG AFFILIATE.SHOPEE.VN (AUTO HOOK TOKEN BẢO MẬT) ---

    function runAffiliateAutoHook() {
        console.log('[Shopee VN Scraper] Khởi chạy chế độ tự động bắt Token bảo mật...');

        const notify = document.createElement('div');
        notify.id = 'scraper-vn-hook-notify';
        notify.style.position = 'fixed';
        notify.style.bottom = '20px';
        notify.style.left = '20px';
        notify.style.zIndex = '9999999';
        notify.style.backgroundColor = '#1e1e1e';
        notify.style.color = '#ffffff';
        notify.style.padding = '10px 15px';
        notify.style.borderRadius = '5px';
        notify.style.border = '1px solid #ee4d2d';
        notify.style.boxShadow = '0 4px 12px rgba(0,0,0,0.5)';
        notify.style.fontFamily = 'Arial, sans-serif';
        notify.style.fontSize = '12px';
        notify.innerHTML = `
            <div style="font-weight: bold; color: #ee4d2d; margin-bottom: 3px;">🛡️ Shopee VN Scraper Active</div>
            <div id="scraper-hook-status" style="color: #aaa;">Đang lắng nghe Security Headers...</div>
        `;
        document.body.appendChild(notify);

        function setHookStatus(text, success = true) {
            const statusEl = document.getElementById('scraper-hook-status');
            if (statusEl) {
                statusEl.innerText = text;
                statusEl.style.color = success ? '#4caf50' : '#ff9800';
            }
        }

        let headersConfig = GM_getValue('api_headers_config_vn') || {};
        if (typeof headersConfig !== 'object' || headersConfig === null) {
            headersConfig = {
                cookie: '',
                afAcEncSzToken: '',
                afAcEncDat: '',
                xSapSec: '',
                xSapRi: '',
                xSzSdkVersion: '1.12.21'
            };
        }

        function updateConfigIfChanged(key, value) {
            let changed = false;
            if (headersConfig[key] !== value) {
                headersConfig[key] = value;
                changed = true;
            }

            const currentCookie = document.cookie;
            if (headersConfig.cookie !== currentCookie) {
                headersConfig.cookie = currentCookie;
                changed = true;
            }

            if (changed) {
                GM_setValue('api_headers_config_vn', headersConfig);
                setHookStatus('✅ Đã tự động bắt & cập nhật Token bảo mật!');
                console.log(`[Shopee VN Scraper] Đã tự động cập nhật: ${key}`);
            }
        }

        const win = (typeof unsafeWindow !== 'undefined') ? unsafeWindow : window;

        const originalSetRequestHeader = win.XMLHttpRequest.prototype.setRequestHeader;
        win.XMLHttpRequest.prototype.setRequestHeader = function (header, value) {
            const lowerHeader = header.toLowerCase();
            if (lowerHeader === 'af-ac-enc-sz-token') {
                updateConfigIfChanged('afAcEncSzToken', value);
            } else if (lowerHeader === 'af-ac-enc-dat') {
                updateConfigIfChanged('afAcEncDat', value);
            } else if (lowerHeader === 'x-sap-sec') {
                updateConfigIfChanged('xSapSec', value);
            } else if (lowerHeader === 'x-sap-ri') {
                updateConfigIfChanged('xSapRi', value);
            } else if (lowerHeader === 'x-sz-sdk-version') {
                updateConfigIfChanged('xSzSdkVersion', value);
            }
            return originalSetRequestHeader.apply(this, arguments);
        };

        const originalFetch = win.fetch;
        win.fetch = async function (resource, init) {
            if (init && init.headers) {
                let headersObj = {};
                if (init.headers instanceof Headers) {
                    for (let [key, val] of init.headers.entries()) {
                        headersObj[key.toLowerCase()] = val;
                    }
                } else if (typeof init.headers === 'object') {
                    for (let key in init.headers) {
                        headersObj[key.toLowerCase()] = init.headers[key];
                    }
                }

                if (headersObj['af-ac-enc-sz-token']) {
                    updateConfigIfChanged('afAcEncSzToken', headersObj['af-ac-enc-sz-token']);
                }
                if (headersObj['af-ac-enc-dat']) {
                    updateConfigIfChanged('afAcEncDat', headersObj['af-ac-enc-dat']);
                }
                if (headersObj['x-sap-sec']) {
                    updateConfigIfChanged('xSapSec', headersObj['x-sap-sec']);
                }
                if (headersObj['x-sap-ri']) {
                    updateConfigIfChanged('xSapRi', headersObj['x-sap-ri']);
                }
                if (headersObj['x-sz-sdk-version']) {
                    updateConfigIfChanged('xSzSdkVersion', headersObj['x-sz-sdk-version']);
                }
            }
            return originalFetch.apply(this, arguments);
        };

        // Đăng ký Bridge lắng nghe tín hiệu ping từ trang Scraper
        GM_addValueChangeListener('scraper_ping_vn', (name, oldValue, newValue, remote) => {
            if (remote && newValue) {
                GM_setValue('scraper_pong_vn', Date.now());
            }
        });

        // Đăng ký Bridge lắng nghe yêu cầu gọi API từ trang Scraper
        GM_addValueChangeListener('scraper_request_vn', async (name, oldValue, newValue, remote) => {
            if (remote && newValue) {
                const { itemId, requestId } = newValue;
                console.log(`[Shopee VN Scraper Bridge] Nhận yêu cầu lấy hoa hồng cho SP: ${itemId}`);
                try {
                    // Gọi API trực tiếp bằng fetch của trang affiliate.shopee.vn.
                    const response = await win.fetch(`https://affiliate.shopee.vn/api/v3/offer/product?item_id=${itemId}`, {
                        headers: {
                            'accept': 'application/json, text/plain, */*',
                            'affiliate-program-type': '1',
                            'referer': `https://affiliate.shopee.vn/offer/product_offer/${itemId}`
                        }
                    });
                    const resText = await response.text();
                    console.log(`[Shopee VN Scraper Bridge] Phản hồi từ API cho SP ${itemId}: status ${response.status}`);
                    GM_setValue('scraper_response_vn', {
                        requestId,
                        itemId,
                        status: response.status,
                        responseText: resText,
                        timestamp: Date.now()
                    });
                } catch (err) {
                    console.error(`[Shopee VN Scraper Bridge] Lỗi gọi API cho SP ${itemId}:`, err);
                    GM_setValue('scraper_response_vn', {
                        requestId,
                        itemId,
                        status: 500,
                        error: err.message || 'Lỗi kết nối mạng',
                        timestamp: Date.now()
                    });
                }
            }
        });
    }
})();
