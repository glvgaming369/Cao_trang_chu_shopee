// ==UserScript==
// @name         Shopee Unified Market Scraper (VN + PH)
// @namespace    http://tampermonkey.net/
// @version      3.0
// @description  Cào dữ liệu sản phẩm Shopee VN & PH (tự nhận diện theo domain) theo danh mục hoặc từ khóa, áp dụng cơ chế mở tab ngầm bắt gói tin API để tránh ban tài khoản, và đồng bộ Google Sheets / Xuất Excel.
// @author       Antigravity
// @match        https://shopee.vn/*
// @match        https://affiliate.shopee.vn/*
// @match        https://shopee.ph/*
// @match        https://affiliate.shopee.ph/*
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_deleteValue
// @grant        GM_registerMenuCommand
// @grant        GM_xmlhttpRequest
// @grant        GM_addValueChangeListener
// @grant        GM_removeValueChangeListener
// @grant        GM_openInTab
// @connect      affiliate.shopee.vn
// @connect      affiliate.shopee.ph
// @connect      oauth2.googleapis.com
// @connect      sheets.googleapis.com
// @connect      videoai-api.devappnow.com
// @require      https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js
// @run-at       document-start
// ==/UserScript==

(function () {
    'use strict';

    // ============================================================
    // CẤU HÌNH HỆ THỐNG THEO THỊ TRƯỜNG
    // ============================================================
    const CATEGORIES_VN = {
        "Balo & Túi Ví Nam": ["Ba Lô Nam", "Ba Lô Laptop Nam", "Túi & Cặp Đựng Laptop", "Túi Chống Sốc Laptop Nam", "Túi Tote Nam", "Cặp Xách Công Sở Nam", "Ví Cầm Tay Nam", "Túi Đeo Hông & Túi Đeo Ngực Nam", "Túi Đeo Chéo Nam", "Bóp/Ví Nam", "Khác"],
        "Bách Hóa Online": ["Đồ ăn vặt", "Đồ chế biến sẵn", "Nhu yếu phẩm", "Nguyên liệu nấu ăn", "Đồ làm bánh", "Sữa - trứng", "Đồ uống", "Ngũ cốc & mứt", "Các loại bánh", "Đồ uống có cồn", "Bộ quà tặng", "Thực phẩm tươi sống và thực phẩm đông lạnh", "Khác"],
        "Chăm Sóc Thú Cưng": ["Thức ăn cho thú cưng", "Phụ kiện cho thú cưng", "Vệ sinh cho thú cưng", "Quần áo thú cưng", "Chăm sóc sức khỏe", "Làm đẹp cho thú cưng", "Khác"],
        "Dụng cụ và thiết bị tiện ích": ["Dụng cụ cầm tay", "Dụng cụ điện và thiết bị lớn", "Thiết bị mạch điện", "Vật liệu xây dựng", "Thiết bị và phụ kiện xây dựng"],
        "Điện Thoại & Phụ Kiện": ["Điện thoại", "Máy tính bảng", "Pin Dự Phòng", "Pin Gắn Trong, Cáp và Bộ Sạc", "Ốp lưng, bao da, Miếng dán điện thoại", "Bảo vệ màn hình", "Đế giữ điện thoại", "Thẻ nhớ", "Sim", "Phụ kiện khác", "Thiết bị khác"],
        "Đồ Chơi": ["Sở thích & Sưu tầm", "Đồ chơi giải trí", "Đồ chơi giáo dục", "Đồ chơi cho trẻ sơ sinh & trẻ nhỏ", "Đồ chơi vận động & ngoài trời", "Búp bê & Đồ chơi nhồi bông"],
        "Đồng Hồ": ["Đồng Hồ Nam", "Đồng Hồ Nữ", "Bộ Đồng Hồ & Đồng Hồ Cặp", "Đồng Hồ Trẻ Em", "Phụ Kiện Đồng Hồ", "Khác"],
        "Giày Dép Nam": ["Bốt", "Giày Thể Thao/ Sneakers", "Giày Sục", "Giày Tây Lười", "Giày Oxfords & Giày Buộc Dây", "Xăng-đan và Dép", "Phụ kiện giày dép", "Khác"],
        "Giày Dép Nữ": ["Bốt", "Giày Thể Thao/ Sneaker", "Giày Đế Bằng", "Giày Cao Gót", "Giày Đế Xuồng", "Xăng-đan Và Dép", "Phụ Kiện Giày", "Giày Khác"],
        "Giặt Giũ & Chăm Sóc Nhà Cửa": ["Giặt giũ & Chăm sóc nhà cửa", "Giấy vệ sinh, khăn giấy", "Vệ sinh nhà cửa", "Vệ sinh bát đĩa", "Dụng cụ vệ sinh", "Chất khử mùi, làm thơm", "Thuốc diệt côn trùng", "Túi, màng bọc thực phẩm", "Bao bì, túi đựng rác"],
        "Máy Tính & Laptop": ["Máy Tính Bàn", "Màn Hình", "Linh Kiện Máy Tính", "Thiết Bị Lưu Trữ", "Thiết Bị Mạng", "Máy In, Máy Scan & Máy Chiếu", "Phụ Kiện Máy Tính", "Laptop", "Khác", "Gaming"],
        "Máy Ảnh & Máy Quay Phim": ["Máy ảnh - Máy quay phim", "Camera giám sát & Camera hệ thống", "Thẻ nhớ", "Ống kính", "Phụ kiện máy ảnh", "Máy bay camera & Phụ kiện"],
        "Mẹ & Bé": ["Đồ dùng du lịch cho bé", "Đồ dùng ăn dặm cho bé", "Phụ kiện cho mẹ", "Chăm sóc sức khỏe mẹ", "Đồ dùng phòng tắm & Chăm sóc cơ thể bé", "Đồ dùng phòng ngủ cho bé", "An toàn cho bé", "Thực phẩm cho bé", "Chăm sóc sức khỏe bé", "Tã & bô em bé", "Đồ chơi", "Bộ & Gói quà tặng", "Khác", "Sữa công thức trên 24 tháng", "Sữa công thức 0-24 tháng tuổi"],
        "Nhà Cửa & Đời Sống": ["Chăn, Ga, Gối & Nệm", "Đồ nội thất", "Trang trí nhà cửa", "Dụng cụ & Thiết bị tiện ích", "Đồ dùng nhà bếp và hộp đựng thực phẩm", "Đèn", "Ngoài trời & Sân vườn", "Đồ dùng phòng tắm", "Vật phẩm thờ cúng", "Đồ trang trí tiệc", "Chăm sóc nhà cửa và giặt ủi", "Sắp xếp nhà cửa", "Dụng cụ pha chế", "Tinh dầu thơm phòng", "Đồ dùng phòng ăn"],
        "Nhà Sách Online": ["Sách Tiếng Việt", "Sách ngoại văn", "Gói Quà", "Bút viết", "Dụng cụ học sinh & văn phòng", "Màu, Họa Cụ và Đồ Thủ Công", "Sổ và Giấy Các Loại", "Quà Lưu Niệm", "Nhạc cụ và phụ kiện âm nhạc"],
        "Ô Tô & Xe Máy & Xe Đạp": ["Xe đạp, xe điện", "Mô tô, xe máy", "Xe Ô tô", "Mũ bảo hiểm", "Phụ kiện xe máy", "Phụ kiện xe đạp", "Phụ kiện bên trong ô tô", "Dầu nhớt & dầu nhờn", "Phụ tùng ô tô", "Phụ tùng xe máy", "Phụ kiện bên ngoài ô tô", "Chăm sóc ô tô", "Dịch vụ cho xe"],
        "Phụ Kiện & Trang Sức Nữ": ["Nhẫn", "Bông tai", "Khăn choàng", "Găng tay", "Phụ kiện tóc", "Vòng tay & Lắc tay", "Lắc chân", "Mũ", "Dây chuyền", "Kính mắt", "Kim loại quý", "Thắt lưu", "Cà vạt & Nơ cổ", "Phụ kiện thêm", "Bộ phụ kiện", "Khác", "Vớ/ Tất", "Ô/Dù"],
        "Sắc Đẹp": ["Chăm sóc da mặt", "Tắm & chăm sóc cơ thể", "Trang điểm", "Chăm sóc tóc", "Dụng cụ & Phụ kiện Làm đẹp", "Vệ sinh răng miệng", "Nước hoa", "Chăm sóc nam giới", "Khác", "Chăm sóc phụ nữ", "Bộ sản phẩm làm đẹp"],
        "Sức Khỏe": ["Vật tư y tế", "Chống muỗi & xua đuổi côn trùng", "Thực phẩm chức năng", "Tã người lớn", "Hỗ trợ làm đẹp", "Dụng cụ massage và trị liệu", "Khác"],
        "Thiết Bị Điện Gia Dụng": ["Đồ gia dụng nhà bếp", "Đồ gia dụng lớn", "Máy hút bụi & Thiết bị làm sạch", "Quạt & Máy nóng lạnh", "Thiết bị chăm sóc quần áo", "Khác", "Máy xay, ép, máy đánh trứng trộn bột, máy xay thực phẩm", "Bếp điện"],
        "Thiết Bị Điện Tử": ["Phụ kiện tivi", "Máy Game Console", "Phụ kiện Console", "Đĩa game", "Linh phụ kiện", "Tai nghe nhét tai", "Loa", "Tivi", "Tivi Box", "Headphones"],
        "Thể Thao & Du Lịch": ["Vali", "Túi du lịch", "Phụ kiện du lịch", "Dụng Cụ Thể Thao & Dã Ngoại", "Giày Thể Thao", "Thời Trang Thể Thao & Dã Ngoại", "Phụ Kiện Thể Thao & Dã Ngoại", "Khác"],
        "Thời Trang Nam": ["Áo Khoác", "Áo Vest và Blazer", "Áo Hoodie, Áo Len & Áo Nỉ", "Quần Jeans", "Quần Dài/Quần Âu", "Quần Short", "Áo", "Áo Ba Lỗ", "Đồ Lót", "Đồ Ngủ", "Đồ Bộ", "Vớ/Tất", "Trang Phục Truyền Thống", "Đồ Hóa Trang", "Trang Phục Ngành Nghề", "Khác", "Trang Sức Nam", "Kính Mắt Nam", "Thắt Lưng Nam", "Cà vạt & Nơ cổ", "Phụ Kiện Nam"],
        "Thời Trang Nữ": ["Quần", "Quần đùi", "Chân váy", "Quần jeans", "Đầm/Váy", "Váy cưới", "Đồ liền thân", "Áo khoác, Áo choàng & Vest", "Áo len & Cardigan", "Hoodie và Áo nỉ", "Bộ", "Đồ lót", "Đồ ngủ", "Áo", "Đồ tập", "Đồ Bầu", "Đồ truyền thống", "Đồ hóa trang", "Vải", "Vớ/ Tất", "Khác"],
        "Thời Trang Trẻ Em": ["Trang phục bé trai", "Trang phục bé gái", "Giày dép bé trai", "Giày dép bé gái", "Khác", "Quần áo em bé", "Giày tập đi & Tất sơ sinh", "Phụ kiện trẻ em"],
        "Túi Ví Nữ": ["Ba Lô Nữ", "Cặp Laptop", "Ví Dự Tiệc & Ví Cầm Tay", "Túi Đeo Hông & Túi Đeo Ngực", "Túi Tote", "Túi Quai Xách", "Túi Đeo Chéo & Túi Đeo Vai", "Ví/Bóp Nữ", "Phụ Kiện Túi", "Khác"],
        "Voucher & Dịch Vụ": ["Nhà hàng & Ăn uống", "Sự kiện & Giải trí", "Nạp tiền tài khoản", "Sức khỏe & Làm đẹp", "Gọi xe", "Khóa học", "Du lịch & Khách sạn", "Mua sắm", "Mã quà tặng Shopee", "Thanh toán hóa đơn", "Dịch vụ khác"]
    };

    const CATEGORIES_PH = {
        "Audio": ["Audio & Video Cables & Converters", "Earphones, Headphones & Headsets", "Amplifiers & Mixers", "Speakers and Karaoke", "Home Audio & Speakers", "Media Players"],
        "Babies & Kids": ["Baby Detergent", "Babies' Fashion", "Rain Gear", "Nursery", "Moms & Maternity", "Baby Gear", "Health & Safety", "Bath & Skin Care", "Boys' Fashion", "Girls' Fashion", "Feeding & Nursing", "Feeding", "Diapers & Wipes", "Others"],
        "Cameras": ["Car / Dash Camera", "Drones", "CCTV / IP Camera", "Action Camera", "Camera Accessories", "Digital Camera", "Others"],
        "Gaming": ["Computer Gaming", "Mobile Gaming", "Console Gaming", "Others"],
        "Groceries": ["Seasoning, Staple Foods & Baking Ingredients", "Gift Set & Hampers", "Dairy & Eggs", "Cigarettes", "Superfoods & Healthy Foods", "Breakfast Food", "Snack & Sweets", "Frozen & Fresh foods", "Alcoholic Beverages", "Laundry & Household Care", "Beverages", "Others"],
        "Health & Personal Care": ["Sexual Wellness", "Medical Supplies", "Men's Grooming", "Health Supplements", "Slimming", "Suncare", "Whitening", "Personal Care", "Bath & Body", "Hair Care", "Skin Care", "Others"],
        "Hobbies & Stationery": ["E-Books", "Books and Magazines", "Paper Supplies", "Writing Materials", "Religious Artifacts", "Packaging & Wrapping", "Arts & Crafts", "School & Office Supplies", "Musical Instruments", "Others"],
        "Home & Living": ["Hand Warmers, Hot Water Bags & Ice Bags", "Home Maintenance", "Furniture", "Lighting", "Party Supplies", "Beddings", "Bath", "Glassware & Drinkware", "Dinnerware", "Bakeware", "Kitchenware", "Sinkware", "Power Tools", "Home Improvement", "Storage & Organization", "Home Decor", "Garden Decor", "Outdoor & Garden", "Others"],
        "Home Appliances": ["Small Household Appliances", "Home Appliance Parts & Accessories", "Large Appliances", "Vacuum Cleaners & Floor Care", "Humidifier & Air Purifier", "Cooling & Heating", "Specialty Appliances", "Small kitchen Appliances", "Garment Care", "Others"],
        "Home Entertainment": ["Projectors", "TV Accessories", "Television", "Others"],
        "Laptops & Computers": ["USB Gadgets", "Computer Hardware", "Software", "Printers and Inks", "Storage", "Computer Accessories", "Network Components", "Laptops and Desktops", "Others"],
        "Makeup & Fragrances": ["Palettes & Makeup Sets", "Tools & Accessories", "Nails", "Fragrances", "Face Makeup", "Lip Makeup", "Eye Makeup", "Others"],
        "Men's Apparel": ["Tops", "Shorts", "Pants", "Jeans", "Underwear", "Socks", "Hoodies & Sweatshirts", "Jackets & Sweaters", "Sleepwear", "Suits", "Sets", "Occupational Attire", "Traditional Wear", "Costumes", "Others"],
        "Men's Bags & Accessories": ["Hats & Caps", "Wallets", "Eyewear", "Accessories", "Jewelry", "Watches", "Men's Bags", "Accessories Sets & Packages"],
        "Men's Shoes": ["Loafer & Boat Shoes", "Sneakers", "Sandals & Flip Flops", "Boots", "Formal", "Shoe Care & Accessories", "Others"],
        "Mobiles & Gadgets": ["Portable Audio", "Wearables", "E-Cigarettes", "Tablets", "Mobiles"],
        "Mobiles Accessories": ["Others Mobile Accessories", "Attachments", "Cases & Covers", "Powerbanks & Chargers"],
        "Motors": ["Car Care & Detailing", "Automotive Parts", "Engine Parts", "Ignition", "Exterior Car Accessories", "Oils, Coolants, & Fluids", "Car Electronics", "Moto Riding & Protective Gear", "Tools & Garage", "Motorcycle Accessories", "Motorcycle & ATV Parts", "Interior Car Accessories", "Others", "Motorcycles"],
        "Pet Care": ["Toys & Accessories", "Litter & Toilet", "Pet Essentials", "Pet Clothing & Accessories", "Pet Grooming Supplies", "Pet Toys & Accessories", "Pet Food & Treats", "Others"],
        "Sports & Travel": ["Travel Bags", "Travel Accessories", "Travel Organizer", "Kid's Activewear", "Boxing & MMA", "Weather Protection", "WinterSports Gear", "Outdoor Recreation", "Leisure Sports & Game Room", "Golf", "Racket Sports", "Sports Bags", "Women's Activewear", "Men's Activewear", "Cycling, Skates & Scooters", "Team Sports", "Water Sports", "Camping & Hiking", "Weightlifting", "Fitness Accessory", "Yoga", "Exercise & Fitness", "Others"],
        "Toys, Games & Collectibles": ["Celebrity Merchandise", "Dress Up & Pretend", "Blasters & Toy Guns", "Sports & Outdoor Toys", "Dolls", "Educational Toys", "Electronic Toys", "Boards & Family Games", "Collectibles", "Character", "Action Figure", "Others"],
        "Women Accessories": ["Jewelry", "Watches", "Hair Accessories", "Eyewear", "Wallets & Pouches", "Hats & Caps", "Belts & Scarves", "Gloves", "Accessories Sets & Packages", "Additional Accessories", "Watch & Jewelry Organizers", "Others"],
        "Women's Apparel": ["Dresses", "Tops", "Tees", "Shorts", "Pants", "Jeans", "Skirts", "Jumpsuits & Rompers", "Lingerie & Nightwear", "Sets", "Swimsuit", "Jackets & Outerwear", "Plus Size", "Sweater & Cardigans", "Maternity Wear", "Socks & Stockings", "Costumes", "Traditional Wear", "Fabric"],
        "Women's Bags": ["Shoulder Bags", "Tote Bags", "Handbags", "Clutches", "Backpacks", "Drawstrings", "Accessories", "Others"],
        "Women's Shoes": ["Flats", "Heels", "Flip Flops", "Sneakers", "Wedges & Platforms", "Boots", "Shoe Care & Accessories", "Others"]
    };

    const MARKETS = {
        vn: {
            code: 'vn',
            label: 'Việt Nam',
            domain: 'shopee.vn',
            affiliateDomain: 'affiliate.shopee.vn',
            currency: '₫',
            currencyBefore: false,
            currencyStrip: /[₫đ]/g,
            decimalMode: 'integer',
            imgDomain: 'down-vn',
            topSalesLabel: 'Bán chạy',
            categories: CATEGORIES_VN,
            defaults: { priceMin: 80000, commMin: 5000, soldMin: 10, sellerCommissionMin: 5000, priceStep: '1000', comStep: '1' }
        },
        ph: {
            code: 'ph',
            label: 'Philippines',
            domain: 'shopee.ph',
            affiliateDomain: 'affiliate.shopee.ph',
            currency: '₱',
            currencyBefore: true,
            currencyStrip: /[₱₫đ]/g,
            decimalMode: 'float',
            imgDomain: 'down-sg',
            topSalesLabel: 'Top Sales',
            categories: CATEGORIES_PH,
            defaults: { priceMin: 120, commMin: 5.0, soldMin: 10, sellerCommissionMin: 5.0, priceStep: '1', comStep: '0.1' }
        }
    };

    const HOST = window.location.hostname;
    const MARKET_CODE = HOST.endsWith('.ph') ? 'ph' : (HOST.endsWith('.vn') ? 'vn' : null);
    if (!MARKET_CODE) return;
    const SITE = MARKETS[MARKET_CODE];

    const NS = 'mktscraper';
    const PANEL_ID = `shopee-${NS}-panel`;

    // Hằng số VideoAI API — khai báo sớm (trước ROUTING) để tránh lỗi Temporal Dead Zone
    // khi runMainShopeeScraper/runAffiliatePortalScraper được gọi ngay tại document-start.
    const VIDEOAI_DEFAULT_ENDPOINT = 'https://videoai-api.devappnow.com/api/products/shopee-cache/batch';
    const VIDEOAI_MAX_BATCH = 200;
    const VIDEOAI_MIN_IMAGES = 3;

    function gmKey(name) {
        const sharedKeys = ['oauth_token', 'token_expiry', 'refresh_token', 'google_config'];
        if (sharedKeys.includes(name)) {
            return `${NS}_${name}`;
        }
        return `${NS}_${name}_${SITE.code}`;
    }

    function elId(name) { return `${NS}-${name}`; }

    function money(v) { return SITE.currencyBefore ? `${SITE.currency}${v}` : `${v}${SITE.currency}`; }

    // Cột dữ liệu lưu Google Sheet / Excel
    const C = {
        itemId: 'Mã SP (Item ID)',
        shopId: 'Mã Shop (Shop ID)',
        name: 'Tên sản phẩm',
        price: `Giá (${SITE.currency})`,
        sold: 'Lượt bán trong 30 ngày',
        url: 'Đường dẫn sản phẩm',
        commission: `Hoa hồng (${SITE.currency})`,
        sellerCom: `Hoa hồng người bán (${SITE.currency})`,
        shopeeCom: `Hoa hồng Shopee (${SITE.currency})`,
        group: 'Nhóm sản phẩm (ID-Group)',
        images: 'Danh sách ảnh',
        category: 'Category',
        stock: 'Tồn kho',
        rating: 'Đánh giá sao',
        ratingCount: 'Lượt đánh giá',
        postDate: 'Ngày đăng bán',
        linkType: 'Phân loại link',
        mergedLink: 'Link gộp'
    };
    const HEADERS = Object.values(C);
    const NUMERIC_COLS = new Set([C.price, C.sold, C.commission, C.sellerCom, C.shopeeCom, C.stock, C.rating, C.ratingCount]);

    // Trạng thái chạy toàn cục
    let isRunning = false;
    let crawledCache = GM_getValue(gmKey('crawled_cache')) || [];
    crawledCache = crawledCache.map(x => String(x));

    let skippedCache = GM_getValue(gmKey('skipped_cache')) || [];
    skippedCache = skippedCache.map(x => String(x));

    let blacklistConfig = GM_getValue(gmKey('blacklist_config'));
    if (!blacklistConfig) {
        blacklistConfig = SITE.code === 'vn'
            ? ['áo ngực', 'quần lót', 'quần chíp', 'quần sịp', 'áo lót', 'nội y', 'đồ lót', 'tất lưới', 'váy ngủ', 'đầm ngủ', 'đồ ngủ gợi cảm', 'váy ngủ gợi cảm', 'đầm ngủ xuyên thấu', 'đồ lót gợi cảm', 'bikini gợi cảm', 'bikini 2 mảnh', 'đồ bơi gợi cảm', 'áo tắm gợi cảm', 'bao cao su', 'gel bôi trơn', 'sextoy', 'đồ chơi tình dục', 'thuốc kích dục', 'tăng sinh lý', 'kéo dài thời gian', 'vòng rung']
            : [];
        GM_setValue(gmKey('blacklist_config'), blacklistConfig);
    }

    // Kiểm tra captcha toàn cục trước khi định tuyến
    const currentHref = window.location.href;
    const isGlobalCaptchaPage = currentHref.includes('/verify/captcha') || 
                                currentHref.includes('/verify/traffic') || 
                                currentHref.includes('/verify/security') ||
                                !!document.querySelector('.check-captcha-box') || 
                                !!document.querySelector('.shopee-captcha-wrapper') ||
                                !!document.getElementById('captcha-submit');

    // ============================================================
    // PHÂN LỘT TUYẾN DẪN CHẠY (ROUTING)
    // ============================================================
    const isProductDetailTab = /\/offer\/product_offer\/\d+/.test(window.location.pathname);
    const isAffiliatePortal = HOST.startsWith('affiliate.');

    // Trang captcha/verify: vẫn dựng scraper như thường (để KHÔNG mất panel điều khiển),
    // đồng thời chạy thêm handler captcha (banner + tự xoá cờ khi giải xong).
    if (isProductDetailTab) {
        runDetailTabHook();
    } else if (isAffiliatePortal) {
        runAffiliatePortalScraper();
    } else {
        runMainShopeeScraper();
    }

    if (isGlobalCaptchaPage) {
        runCaptchaTabHandler();
    }

    // ============================================================
    // HANDLER CHO TAB DÍNH CAPTCHA (tự focus + tự xoá cờ khi giải xong)
    // ============================================================
    function runCaptchaTabHandler() {
        console.log('[Shopee Scraper] Tab dính Captcha/Verify — kích hoạt chế độ chờ giải.');
        GM_setValue(gmKey('global_captcha_detected'), true);
        // Đưa tab captcha ra trước để người dùng thấy & giải ngay (best-effort)
        try { window.focus(); } catch (e) { }

        const addBanner = () => {
            if (document.getElementById(`${NS}-captcha-aff-banner`)) return;
            const banner = document.createElement('div');
            banner.id = `${NS}-captcha-aff-banner`;
            banner.style.cssText = `position: fixed; top: 0; left: 0; width: 100%; z-index: 9999999; background-color: #ee4d2d; color: white; padding: 15px; text-align: center; font-size: 16px; font-weight: bold; box-shadow: 0 2px 10px rgba(0,0,0,0.3);`;
            banner.innerText = `⚠️ PHÁT HIỆN CAPTCHA! Hãy giải captcha trên tab này — tool chính sẽ TỰ ĐỘNG tiếp tục ngay sau khi giải xong.`;
            (document.body || document.documentElement).appendChild(banner);
        };
        if (document.body) addBanner();
        else window.addEventListener('DOMContentLoaded', addBanner);

        // Theo dõi: ngay khi không còn dấu hiệu captcha (giải tại chỗ) -> tự xoá cờ để tool chính tiếp tục
        let cleared = false;
        const watcher = setInterval(() => {
            if (cleared) return;
            const stillCaptcha = window.location.href.includes('/verify/') ||
                !!document.querySelector('.check-captcha-box') ||
                !!document.querySelector('.shopee-captcha-wrapper') ||
                !!document.getElementById('captcha-submit') ||
                !!document.querySelector('iframe[src*="captcha"]');
            if (!stillCaptcha) {
                cleared = true;
                clearInterval(watcher);
                GM_setValue(gmKey('global_captcha_detected'), false);
                const b = document.getElementById(`${NS}-captcha-aff-banner`);
                if (b) b.remove();
            }
        }, 700);
    }

    // ============================================================
    // 1. LOGIC TRÊN TAB CHI TIẾT SẢN PHẨM (HOOK & BẮT GÓI TIN)
    // ============================================================
    function runDetailTabHook() {
        const itemIdMatch = window.location.pathname.match(/\/offer\/product_offer\/(\d+)/);
        const itemId = itemIdMatch ? itemIdMatch[1] : null;

        // Vào được trang chi tiết => không còn bị captcha chặn. Xoá cờ để tool chính tự tiếp tục (không cần bấm nút thủ công).
        if (GM_getValue(gmKey('global_captcha_detected')) === true) {
            GM_setValue(gmKey('global_captcha_detected'), false);
        }

        if (itemId) {
            console.log(`[Shopee Scraper Detail] Hooking XHR & Fetch for Item ID: ${itemId}`);
            const win = (typeof unsafeWindow !== 'undefined') ? unsafeWindow : window;

            // Hook Fetch API
            const originalFetch = win.fetch;
            win.fetch = async function (...args) {
                const response = await originalFetch.apply(this, args);
                const url = args[0];
                if (typeof url === 'string' && url.includes('/api/v3/offer/product')) {
                    try {
                        const clone = response.clone();
                        const text = await clone.text();
                        console.log(`[Shopee Scraper Detail] Captured fetch API response for Item ID: ${itemId}`);
                        
                        GM_setValue(`captured_${itemId}`, {
                            itemId: itemId,
                            status: response.status,
                            responseText: text,
                            timestamp: Date.now()
                        });

                        setTimeout(() => {
                            try { window.close(); } catch (e) { }
                        }, 500);
                    } catch (e) {
                        console.error('[Shopee Scraper Detail] Fetch hook error:', e);
                    }
                }
                return response;
            };

            // Hook XMLHttpRequest Prototype
            const originalOpen = win.XMLHttpRequest.prototype.open;
            win.XMLHttpRequest.prototype.open = function(method, url, ...rest) {
                const urlStr = String(url);
                this.addEventListener('readystatechange', async () => {
                    if (this.readyState === 4) {
                        if (urlStr.includes('/api/v3/offer/product')) {
                            console.log(`[Shopee Scraper Detail] Captured XHR API response for Item ID: ${itemId}`);
                            try {
                                GM_setValue(`captured_${itemId}`, {
                                    itemId: itemId,
                                    status: this.status,
                                    responseText: this.responseText,
                                    timestamp: Date.now()
                                });
                                setTimeout(() => {
                                    try { window.close(); } catch (e) { }
                                }, 500);
                            } catch (e) {
                                console.error('[Shopee Scraper Detail] XHR capture save error:', e);
                            }
                        }
                    }
                });
                return originalOpen.apply(this, [method, url, ...rest]);
            };
        }
    }

    // ============================================================
    // CÁC HÀM TIỆN ÍCH DÙNG CHUNG CỦA BẢNG ĐIỀU KHIỂN
    // ============================================================
    function sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    function cleanNumber(str) {
        if (!str) return 0;
        let s = String(str).trim().toLowerCase();

        const kMatch = s.match(/([\d.,]+)\s*k/);
        if (kMatch) {
            const num = parseFloat(kMatch[1].replace(/,/g, '.'));
            return isNaN(num) ? 0 : Math.round(num * 1000);
        }

        s = s.replace(SITE.currencyStrip, '');
        const numMatch = s.match(/[\d.,]+/g);
        if (!numMatch) return 0;

        let longest = numMatch.reduce((a, b) => (a.length > b.length ? a : b), '');

        if (SITE.decimalMode === 'integer') {
            longest = longest.replace(/[.,]/g, '');
            const parsed = parseInt(longest, 10);
            return isNaN(parsed) ? 0 : parsed;
        } else {
            longest = longest.replace(/,/g, '');
            const parsed = parseFloat(longest);
            return isNaN(parsed) ? 0 : Math.round(parsed * 100) / 100;
        }
    }

    function parseProductUrl(href) {
        if (!href) return null;
        let decodedHref = href;
        try { decodedHref = decodeURIComponent(href); } catch (e) { }

        // Dạng /product/{shopId}/{itemId} (link người dùng dán) hoặc ...-i.{shopId}.{itemId} (link cũ)
        let match = decodedHref.match(/\/product\/(\d+)\/(\d+)/);
        if (!match) match = decodedHref.match(/i\.(\d+)\.(\d+)/);
        if (match) {
            return {
                shopId: String(match[1]),
                itemId: String(match[2]),
                productUrl: `https://${SITE.domain}/product/${match[1]}/${match[2]}`
            };
        }
        return null;
    }

    function getCategoryFromUrl() {
        try {
            const match = window.location.href.match(/\/([a-zA-Z0-9-_]+)-cat\.\d+/);
            if (match && match[1]) return match[1].replace(/-/g, ' ');
        } catch (e) {
            console.error('Lỗi phân tích category từ URL:', e);
        }
        return '';
    }

    function getCategory(data, keyword = '') {
        try {
            if (data && data.batch_item_for_item_card_full) {
                const card = data.batch_item_for_item_card_full;
                if (card.categories && card.categories.length > 0) {
                    return card.categories.map(c => c.display_name || c.name).join(' > ');
                }
            }
        } catch (e) {
            console.error(e);
        }
        if (isAffiliatePortal) {
            return keyword || '';
        } else {
            const catUrl = getCategoryFromUrl();
            if (catUrl) return catUrl;
            const filterConfig = GM_getValue(`${NS}_filter_config_${SITE.code}`) || {};
            if (filterConfig.selectedCategory && filterConfig.selectedCategory !== 'no-category') {
                return filterConfig.selectedCategory;
            }
        }
        return '';
    }

    function formatDate(timestamp) {
        if (!timestamp) return '';
        try {
            const date = new Date(timestamp * 1000);
            const day = String(date.getDate()).padStart(2, '0');
            const month = String(date.getMonth() + 1).padStart(2, '0');
            return `${day}/${month}/${date.getFullYear()}`;
        } catch (e) {
            return '';
        }
    }

    async function smoothScrollToBottom() {
        return new Promise((resolve) => {
            let currentPosition = 0;
            const distance = 400;
            const interval = 200;
            let scrollCount = 0;
            const maxScrolls = 10;

            window.scrollTo(0, 0);
            const timer = setInterval(() => {
                currentPosition += distance;
                window.scrollTo(0, currentPosition);
                scrollCount++;
                const scrollHeight = document.documentElement.scrollHeight;
                if (scrollCount >= maxScrolls || currentPosition >= scrollHeight) {
                    clearInterval(timer);
                    resolve();
                }
            }, interval);
        });
    }

    function gmRequest(opts) {
        return new Promise((resolve, reject) => {
            GM_xmlhttpRequest({
                ...opts,
                onload: (res) => resolve(res),
                onerror: (err) => reject(err)
            });
        });
    }

    function imgUrls(ids) {
        const domain = SITE.imgDomain || 'down-sg';
        return (ids || []).map(id => `https://${domain}.img.susercontent.com/file/${id}`).join('|');
    }

    function getSpreadsheetId(url) {
        if (!url) return null;
        const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
        return match ? match[1] : null;
    }

    // ============================================================
    // GOOGLE SHEETS OAUTH2 & KHỞI CHẠY GOOGLE SHEETS API
    // ============================================================
    async function getGoogleAccessToken() {
        const storedToken = GM_getValue(gmKey('oauth_token'), null);
        const tokenExpiry = GM_getValue(gmKey('token_expiry'), 0);
        if (storedToken && Date.now() < tokenExpiry - 300000) {
            return storedToken;
        }

        const googleConfig = GM_getValue(gmKey('google_config')) || {};
        const client_id = googleConfig.clientId;
        const client_secret = googleConfig.clientSecret;
        const refresh_token = GM_getValue(gmKey('refresh_token'), null);

        if (!client_id || !client_secret) {
            addLog('Lỗi: Hãy cấu hình Client ID và Secret của Google trước!', 'error');
            throw new Error('Missing OAuth credentials');
        }
        if (!refresh_token) {
            addLog('Chưa có Refresh Token. Vui lòng bấm kết nối Google Sheets!', 'error');
            throw new Error('No refresh token available');
        }

        addLog('Đang tự động làm mới Access Token Google...', 'info');
        const response = await gmRequest({
            method: 'POST',
            url: 'https://oauth2.googleapis.com/token',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            data: `refresh_token=${encodeURIComponent(refresh_token)}` +
                `&client_id=${encodeURIComponent(client_id)}` +
                `&client_secret=${encodeURIComponent(client_secret)}` +
                `&grant_type=refresh_token`
        });
        const res = JSON.parse(response.responseText);
        if (res.access_token) {
            GM_setValue(gmKey('oauth_token'), res.access_token);
            GM_setValue(gmKey('token_expiry'), Date.now() + (res.expires_in * 1000));
            addLog('Làm mới Access Token thành công!', 'success');
            return res.access_token;
        }
        addLog(`Lỗi làm mới Access Token: ${response.responseText}`, 'error');
        GM_setValue(gmKey('refresh_token'), null);
        throw new Error('Refresh token invalid');
    }

    function startGoogleAuthFlow(clientId) {
        if (!clientId) {
            addLog('Lỗi: Điền Client ID trước!', 'error');
            alert('Vui lòng điền Google Client ID trước!');
            return;
        }
        const redirect_uri = 'urn:ietf:wg:oauth:2.0:oob';
        const authUrl = `https://accounts.google.com/o/oauth2/v2/auth?` +
            `client_id=${encodeURIComponent(clientId)}` +
            `&redirect_uri=${encodeURIComponent(redirect_uri)}` +
            `&response_type=code` +
            `&scope=${encodeURIComponent('https://www.googleapis.com/auth/spreadsheets')}` +
            `&access_type=offline&prompt=consent`;

        addLog('Đang mở trang đăng nhập Google cấp quyền...', 'info');
        window.open(authUrl, '_blank');
        addLog('Hãy dán mã Authorization Code nhận được vào ô nhập ở panel cấu hình và nhấn OK.', 'warn');
    }

    async function exchangeAuthCodeForTokens(authCode, clientId, clientSecret) {
        try {
            addLog('Đang trao đổi mã lấy token truy cập...', 'info');
            const response = await gmRequest({
                method: 'POST',
                url: 'https://oauth2.googleapis.com/token',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                data: `code=${encodeURIComponent(authCode)}` +
                    `&client_id=${encodeURIComponent(clientId)}` +
                    `&client_secret=${encodeURIComponent(clientSecret)}` +
                    `&redirect_uri=${encodeURIComponent('urn:ietf:wg:oauth:2.0:oob')}` +
                    `&grant_type=authorization_code`
            });
            const res = JSON.parse(response.responseText);
            if (res.access_token) {
                GM_setValue(gmKey('oauth_token'), res.access_token);
                GM_setValue(gmKey('token_expiry'), Date.now() + (res.expires_in * 1000));
                if (res.refresh_token) GM_setValue(gmKey('refresh_token'), res.refresh_token);
                addLog('KẾT NỐI GOOGLE SPREADSHEETS THÀNH CÔNG!', 'success');
                alert('Đăng nhập và kết nối Google thành công!');
            } else {
                addLog(`Lỗi Google API: ${response.responseText}`, 'error');
                alert(`Lỗi: ${res.error_description || 'Không lấy được Token. Hãy kiểm tra Client ID & Client Secret.'}`);
            }
        } catch (err) {
            addLog(`Lỗi xác thực: ${err.message}`, 'error');
        }
    }

    async function testGoogleSheetsConnection() {
        try {
            const googleConfig = GM_getValue(gmKey('google_config')) || {};
            const spreadsheetId = getSpreadsheetId(googleConfig.sheetUrl);
            if (!spreadsheetId) {
                addLog('Lỗi kiểm tra: Đường dẫn Google Sheets không đúng!', 'error');
                alert('Vui lòng nhập đường dẫn Google Sheets hợp lệ!');
                return;
            }

            addLog('Đang kết nối thử nghiệm đến Google Sheets API...', 'info');
            const token = await getGoogleAccessToken();
            const response = await gmRequest({
                method: 'GET',
                url: `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?fields=properties.title,sheets.properties.title`,
                headers: { 'Authorization': `Bearer ${token}` }
            });

            if (response.status !== 200) {
                addLog(`Thử nghiệm thất bại (Status ${response.status})`, 'error');
                alert(`Kết nối thử nghiệm thất bại! Mã lỗi: ${response.status}`);
                if (response.status === 401) GM_setValue(gmKey('oauth_token'), null);
                return;
            }

            const data = JSON.parse(response.responseText);
            const title = data.properties.title;
            const sheets = data.sheets.map(s => s.properties.title);
            addLog('TEST KẾT NỐI SPREADSHEET THÀNH CÔNG!', 'success');
            addLog(`Tên bảng tính: "${title}" | Danh sách Tab: [${sheets.join(', ')}]`, 'info');

            const tabName = googleConfig.tabName || 'Sheet1';
            if (sheets.includes(tabName)) {
                addLog(`Tab "${tabName}": Sẵn sàng!`, 'success');
                alert(`Thử nghiệm thành công!\nSpreadsheet: ${title}\nTab "${tabName}" đã sẵn sàng.`);
            } else {
                addLog(`Tab "${tabName}": Chưa tạo (sẽ tự động tạo mới khi lưu sản phẩm)`, 'warn');
                alert(`Thử nghiệm thành công!\nSpreadsheet: ${title}\nCảnh báo: Tab "${tabName}" chưa tồn tại.`);
            }
        } catch (e) {
            addLog(`Lỗi kiểm tra kết nối: ${e.message}`, 'error');
        }
    }

    async function syncCacheFromGoogleSheets() {
        try {
            const googleConfig = GM_getValue(gmKey('google_config')) || {};
            const spreadsheetId = getSpreadsheetId(googleConfig.sheetUrl);
            if (!spreadsheetId) {
                addLog('Lỗi đồng bộ: URL Spreadsheet không hợp lệ!', 'error');
                alert('Vui lòng cấu hình URL Google Sheet hợp lệ trước!');
                return;
            }

            addLog('Đang tải dữ liệu Item ID từ Google Sheets để đồng bộ cache...', 'info');
            const token = await getGoogleAccessToken();
            const tabName = googleConfig.tabName || 'Sheet1';
            const range = `${tabName}!A2:A`;
            const response = await gmRequest({
                method: 'GET',
                url: `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}`,
                headers: { 'Authorization': `Bearer ${token}` }
            });

            if (response.status !== 200) {
                addLog(`Đồng bộ thất bại (Mã lỗi ${response.status})`, 'error');
                if (response.status === 401) GM_setValue(gmKey('oauth_token'), null);
                return;
            }

            const data = JSON.parse(response.responseText);
            const values = data.values;
            if (!values || values.length === 0) {
                addLog('Không tìm thấy Item ID nào để đồng bộ.', 'warn');
                alert('Không có dữ liệu Item ID nào trên Google Sheets!');
                return;
            }

            const sheetIds = values
                .map(row => row[0])
                .filter(val => val !== undefined && val !== null)
                .map(val => String(val).trim())
                .filter(val => /^\d+$/.test(val));

            if (sheetIds.length === 0) {
                addLog('Không có Item ID hợp lệ ở cột A.', 'warn');
                alert('Không tìm thấy Item ID hợp lệ ở cột A!');
                return;
            }

            const beforeCount = crawledCache.length;
            crawledCache = Array.from(new Set([...crawledCache, ...sheetIds]));
            const addedCount = crawledCache.length - beforeCount;

            GM_setValue(gmKey('crawled_cache'), crawledCache);
            updateCacheCount();

            addLog('ĐỒNG BỘ CACHE THÀNH CÔNG!', 'success');
            addLog(`- Quét từ Sheet: ${sheetIds.length} mã. Thêm mới: ${addedCount} mã.`, 'success');
            alert(`Đồng bộ cache thành công!\nThêm mới: ${addedCount} mã.\nTổng Cache: ${crawledCache.length}.`);
        } catch (e) {
            console.error(e);
            addLog(`Lỗi đồng bộ cache: ${e.message}`, 'error');
        }
    }

    async function pushToGoogleSheets(dataList) {
        if (!dataList || dataList.length === 0) return;
        const googleConfig = GM_getValue(gmKey('google_config')) || {};
        const spreadsheetId = getSpreadsheetId(googleConfig.sheetUrl);
        if (!spreadsheetId) {
            throw new Error('Đường dẫn Google Sheets không đúng!');
        }

        const token = await getGoogleAccessToken();
        const tabName = googleConfig.tabName || 'Sheet1';

        const rows = dataList.map(item => HEADERS.map(col => {
            const v = item[col];
            if (v === undefined || v === null) return NUMERIC_COLS.has(col) ? 0 : '';
            return v;
        }));

        let needHeader = true;
        try {
            const checkRes = await gmRequest({
                method: 'GET',
                url: `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(tabName + '!A1:A1')}`,
                headers: { 'Authorization': `Bearer ${token}` }
            });
            if (checkRes.status === 200) {
                const d = JSON.parse(checkRes.responseText);
                needHeader = !(d.values && d.values.length > 0 && d.values[0].length > 0 && String(d.values[0][0]).trim() !== '');
            } else if (checkRes.status === 401) {
                GM_setValue(gmKey('oauth_token'), null);
            }
        } catch (e) {
            needHeader = true;
        }

        const values = needHeader ? [HEADERS, ...rows] : rows;
        const range = `${tabName}!A:R`;

        const response = await gmRequest({
            method: 'POST',
            url: `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}:append?valueInputOption=USER_ENTERED`,
            headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
            data: JSON.stringify({ values })
        });

        if (response.status === 200) {
            addLog(`Ghi thành công ${rows.length} dòng lên Google Sheets!`, 'success');
        } else {
            if (response.status === 401) GM_setValue(gmKey('oauth_token'), null);
            throw new Error(`Google Sheets API error: Status ${response.status}`);
        }
    }

    // ============================================================
    // ĐẨY DỮ LIỆU LÊN VIDEOAI API (shopee-cache/batch)
    // ============================================================
    // Chuyển danh sách sản phẩm nội bộ -> mảng item đúng định dạng API.
    // Trả về { items, skippedFewImages } để lọc trước item có < 3 ảnh (server sẽ skip).
    function buildVideoAIItems(dataList) {
        const items = [];
        const skippedFewImages = [];

        (dataList || []).forEach(it => {
            const imagesStr = it[C.images] || '';
            const images = imagesStr ? String(imagesStr).split('|').map(s => s.trim()).filter(Boolean) : [];

            if (images.length < VIDEOAI_MIN_IMAGES) {
                skippedFewImages.push(it[C.url] || it[C.itemId] || '(không rõ)');
                return;
            }

            const item = {
                url: it[C.url] || '',
                title: it[C.name] || '',
                price: Number(it[C.price]) || 0,
                rating: Number(it[C.rating]) || 0,
                soldCount: Number(it._soldTotal != null ? it._soldTotal : it[C.sold]) || 0,
                stock: Number(it[C.stock]) || 0,
                shopName: it._shopName || '',
                images: images,
                reviewCount: Number(it._reviewCount) || 0,
                ratingCount: Number(it[C.ratingCount]) || 0
            };

            if (it._originalPrice) item.originalPrice = Number(it._originalPrice) || 0;
            if (it[C.category]) {
                const cats = String(it[C.category]).split('>').map(s => s.trim()).filter(Boolean);
                if (cats.length) item.categories = cats;
            }
            if (it._catId) item.categoryIds = [Number(it._catId)].filter(n => !isNaN(n) && n > 0);

            items.push(item);
        });

        return { items, skippedFewImages };
    }

    async function pushToVideoAI(dataList) {
        if (!dataList || dataList.length === 0) return;

        const cfg = GM_getValue(gmKey('videoai_config')) || {};
        const apiKey = (cfg.apiKey || '').trim();
        const endpoint = (cfg.endpoint || VIDEOAI_DEFAULT_ENDPOINT).trim();

        if (!apiKey) {
            throw new Error('Chưa cấu hình VideoAI API Key!');
        }

        const { items, skippedFewImages } = buildVideoAIItems(dataList);
        if (skippedFewImages.length > 0) {
            addLog(`VideoAI: bỏ qua ${skippedFewImages.length} SP do < ${VIDEOAI_MIN_IMAGES} ảnh (client-side).`, 'warn');
        }
        if (items.length === 0) {
            addLog('VideoAI: không có sản phẩm hợp lệ để đẩy.', 'warn');
            return;
        }

        let totalUpserted = 0;
        let totalSkipped = 0;

        for (let i = 0; i < items.length; i += VIDEOAI_MAX_BATCH) {
            const chunk = items.slice(i, i + VIDEOAI_MAX_BATCH);
            const response = await gmRequest({
                method: 'POST',
                url: endpoint,
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': `Bearer ${apiKey}`
                },
                data: JSON.stringify({ items: chunk })
            });

            if (response.status === 401) {
                throw new Error('VideoAI API Key không hợp lệ (401).');
            }
            if (response.status !== 200) {
                throw new Error(`VideoAI API lỗi: Status ${response.status}`);
            }

            let res;
            try { res = JSON.parse(response.responseText); } catch (e) { res = {}; }
            totalUpserted += Number(res.upserted) || 0;
            totalSkipped += Number(res.skipped) || 0;

            if (Array.isArray(res.errors) && res.errors.length > 0) {
                res.errors.slice(0, 10).forEach(err => {
                    addLog(`VideoAI skip: ${err.url} — ${err.reason}`, 'warn');
                });
                if (res.errors.length > 10) {
                    addLog(`VideoAI: ...và ${res.errors.length - 10} lỗi khác.`, 'warn');
                }
            }
        }

        addLog(`VideoAI: ghi thành công ${totalUpserted} SP, bỏ qua ${totalSkipped} SP.`, 'success');
    }

    async function testVideoAIConnection() {
        try {
            const cfg = GM_getValue(gmKey('videoai_config')) || {};
            const apiKey = (cfg.apiKey || '').trim();
            const endpoint = (cfg.endpoint || VIDEOAI_DEFAULT_ENDPOINT).trim();

            if (!apiKey) {
                addLog('Lỗi test VideoAI: chưa nhập API Key!', 'error');
                alert('Vui lòng nhập VideoAI API Key trước!');
                return;
            }

            addLog('Đang test kết nối VideoAI API...', 'info');
            const response = await gmRequest({
                method: 'POST',
                url: endpoint,
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': `Bearer ${apiKey}`
                },
                data: JSON.stringify({ items: [] })
            });

            // Server xác thực key TRƯỚC, rồi mới kiểm payload:
            //   - Sai key  -> 401
            //   - Đúng key + items rỗng -> 400 ("items array is required")
            // => 401 = key sai; 200 hoặc 400 = key HỢP LỆ (đã qua xác thực).
            if (response.status === 401) {
                addLog('Test VideoAI thất bại: API Key không hợp lệ (401).', 'error');
                alert('API Key VideoAI không hợp lệ!');
                return;
            }
            if (response.status === 200 || response.status === 400) {
                addLog('TEST KẾT NỐI VIDEOAI THÀNH CÔNG! API Key hợp lệ.', 'success');
                alert('Kết nối VideoAI thành công! API Key hợp lệ.');
                return;
            }
            addLog(`Test VideoAI: phản hồi bất thường (Status ${response.status}).`, 'warn');
            alert(`Phản hồi VideoAI: Status ${response.status}`);
        } catch (e) {
            addLog(`Lỗi test kết nối VideoAI: ${e.message}`, 'error');
        }
    }

    function exportExcelForData(dataList, prefix = 'data') {
        if (!dataList || dataList.length === 0) {
            addLog('Không có dữ liệu để xuất Excel!', 'warn');
            return;
        }
        const worksheet = XLSX.utils.json_to_sheet(dataList, { header: HEADERS });
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Sản phẩm");

        const now = new Date();
        const dateStr = `${now.getDate()}_${now.getMonth() + 1}_${now.getFullYear()}_${now.getHours()}h${now.getMinutes()}m`;
        const fileName = `${prefix}_${SITE.code}_${dateStr}.xlsx`;

        XLSX.writeFile(workbook, fileName);
        addLog(`Đã xuất file Excel: ${fileName}`, 'success');
        alert(`Đã xuất file Excel thành công: ${fileName}`);
    }

    // ============================================================
    // CHỨC NĂNG MỞ TAB VÀ THU THẬP DỮ LIỆU HOA HỒNG
    // ============================================================
    async function waitForCapturedData(itemId, tab, tabUrl, retryCount = 0) {
        const timeoutMs = 12000;
        const checkInterval = 250;
        
        let elapsed = 0;
        while (elapsed < timeoutMs) {
            if (!isRunning) return null;
            
            // Kiểm tra xem có bị dính captcha ở tab chính hoặc tab phụ không
            const isGlobalCaptcha = GM_getValue(gmKey('global_captcha_detected')) === true;
            if (isGlobalCaptcha) {
                // Tạm dừng bộ đếm timeout và chờ giải captcha
                await checkAndWaitForCaptcha();
                if (!isRunning) return null;
                // Sau khi giải xong, reset lại thời gian chờ để tab phụ tải lại trang và hook bắt được gói tin
                elapsed = 0; 
            }

            const captured = GM_getValue(`captured_${itemId}`);
            if (captured) {
                GM_deleteValue(`captured_${itemId}`);
                try { tab.close(); } catch (e) { }
                return captured;
            }
            
            await sleep(checkInterval);
            elapsed += checkInterval;
        }

        try { tab.close(); } catch (e) { }

        if (retryCount < 1) {
            addLog(`[Timeout] Quá 12s không bắt được gói tin SP ${itemId}. Tiến hành thử lại lần 2 (mở tab nổi)...`, 'warn');
            await sleep(2000);
            const newTab = GM_openInTab(tabUrl, { active: true, insert: true });
            return await waitForCapturedData(itemId, newTab, tabUrl, retryCount + 1);
        }

        addLog(`[Thất bại] Bỏ qua SP ${itemId} do Timeout sau khi thử lại.`, 'error');
        return null;
    }

    async function processProduct(itemData, tabUrl, sellerComMin, keyword = '') {
        const itemId = itemData[C.itemId];
        addLog(`Mở sản phẩm ${itemId}: "${itemData[C.name].substring(0, 15)}..."`, 'info');

        const tab = GM_openInTab(tabUrl, { active: false, insert: true });
        const captured = await waitForCapturedData(itemId, tab, tabUrl, 0);

        if (!captured) return null;

        try {
            const res = JSON.parse(captured.responseText);
            if (res.code === 0 && res.data) {
                const d = res.data;
                const cr = d.commission_rate || {};
                const card = d.batch_item_for_item_card_full;

                const sellerCom = cleanNumber(cr.seller_commission || money(0));
                const shopeeCom = cleanNumber(cr.shopee_commission || money(0));
                const totalCom = cleanNumber(d.commission || money(0));

                if (sellerCom < sellerComMin) {
                    addLog(`Lọc bỏ SP "${itemData[C.name].substring(0, 15)}..." do hoa hồng thực tế (${money(sellerCom)} < ${money(sellerComMin)})`, 'warn');
                    if (!skippedCache.includes(String(itemId))) {
                        skippedCache.push(String(itemId));
                        GM_setValue(gmKey('skipped_cache'), skippedCache);
                        updateCacheCount();
                    }
                    return null;
                }

                const ratingCountArr = (card && card.item_rating) ? card.item_rating.rating_count : [];
                const ratingCount = (Array.isArray(ratingCountArr) && ratingCountArr.length > 0) ? Math.max(...ratingCountArr) : 0;

                if (card) {
                    itemData[C.price] = parseFloat(card.price) / 100000;
                    itemData[C.shopId] = String(card.shopid || itemData[C.shopId]);
                    itemData[C.url] = `https://${SITE.domain}/product/${card.shopid}/${itemId}`;
                    // Điền tên & lượt bán từ card khi chưa có sẵn (trường hợp cào theo link không có DOM nguồn)
                    if (!itemData[C.name]) itemData[C.name] = card.name || '';
                    if (itemData[C.sold] == null || itemData[C.sold] === '') itemData[C.sold] = card.sold || card.historical_sold || 0;
                }
                itemData[C.commission] = totalCom;
                itemData[C.sellerCom] = sellerCom;
                itemData[C.shopeeCom] = shopeeCom;
                itemData[C.images] = imgUrls(card ? card.images : []);
                itemData[C.stock] = card ? card.stock : 0;
                itemData[C.rating] = card && card.item_rating ? Math.round(card.item_rating.rating_star * 100) / 100 : 0;
                itemData[C.ratingCount] = ratingCount;
                // Field thô bổ sung cho VideoAI API (không xuất ra Excel/Sheet)
                if (card) {
                    itemData._shopName = card.shop_name || '';
                    itemData._reviewCount = card.cmt_count || 0;
                    itemData._soldTotal = (card.historical_sold != null ? card.historical_sold : card.sold) || 0;
                    itemData._originalPrice = card.price_before_discount ? parseFloat(card.price_before_discount) / 100000 : 0;
                    itemData._catId = card.catid || 0;
                }
                itemData[C.postDate] = formatDate(card ? card.ctime : 0);
                itemData[C.linkType] = 'L1';
                itemData[C.category] = getCategory(d, keyword);

                addLog(`ĐẠT CHUẨN API L1: "${itemData[C.name].substring(0, 20)}..." | Hoa hồng người bán: ${money(sellerCom)}`, 'success');

                const productsToReturn = [itemData];

                // Cào sản phẩm tương tự L2
                const similarOffers = (d.similar_product_offers && d.similar_product_offers.list) || [];
                if (similarOffers && similarOffers.length > 0) {
                    const parsedSimilar = similarOffers.map(sim => {
                        const simCard = sim.batch_item_for_item_card_full;
                        if (!simCard) return null;

                        const simPriceVal = parseFloat(simCard.price) / 100000;
                        const simSellerRate = parseFloat(sim.seller_commission_rate || "0%") || 0;
                        const simDefaultRate = parseFloat(sim.default_commission_rate || "0%") || 0;

                        const simSellerCom = Math.round((simPriceVal * (simSellerRate / 100)) * 100) / 100;
                        const simTotalCom = Math.round((simPriceVal * (simDefaultRate / 100)) * 100) / 100;
                        const simShopeeCom = Math.max(0, Math.round((simTotalCom - simSellerCom) * 100) / 100);

                        const simItemId = String(sim.item_id || simCard.itemid || '');
                        const simShopId = String(simCard.shopid || '');
                        const simUrl = `https://${SITE.domain}/product/${simShopId}/${simItemId}`;
                        const simRatingCountArr = (simCard.item_rating) ? simCard.item_rating.rating_count : [];
                        const simRatingCount = (Array.isArray(simRatingCountArr) && simRatingCountArr.length > 0) ? Math.max(...simRatingCountArr) : 0;

                        if (crawledCache.includes(simItemId) || skippedCache.includes(simItemId)) {
                            return null;
                        }

                        if (simSellerCom < sellerComMin) {
                            if (!skippedCache.includes(simItemId)) {
                                skippedCache.push(simItemId);
                            }
                            return null;
                        }

                        return {
                            [C.itemId]: simItemId,
                            [C.shopId]: simShopId,
                            [C.name]: simCard.name || '',
                            [C.price]: simPriceVal,
                            [C.sold]: simCard.sold || 0,
                            [C.url]: simUrl,
                            [C.commission]: simTotalCom,
                            [C.sellerCom]: simSellerCom,
                            [C.shopeeCom]: simShopeeCom,
                            [C.group]: `gr-${itemId}`,
                            [C.images]: imgUrls(simCard.images || []),
                            [C.category]: itemData[C.category] || '',
                            [C.stock]: simCard.stock || 0,
                            [C.rating]: simCard.item_rating ? Math.round(simCard.item_rating.rating_star * 100) / 100 : 0,
                            [C.ratingCount]: simRatingCount,
                            [C.postDate]: formatDate(simCard.ctime || 0),
                            [C.linkType]: 'L2',
                            [C.mergedLink]: '',
                            _shopName: simCard.shop_name || '',
                            _reviewCount: simCard.cmt_count || 0,
                            _soldTotal: (simCard.historical_sold != null ? simCard.historical_sold : simCard.sold) || 0,
                            _originalPrice: simCard.price_before_discount ? parseFloat(simCard.price_before_discount) / 100000 : 0,
                            _catId: simCard.catid || 0
                        };
                    }).filter(x => x !== null);

                    const top5Similar = parsedSimilar.slice(0, 5);
                    if (top5Similar.length > 0) {
                        productsToReturn.push(...top5Similar);
                        addLog(`-> Lấy thêm ${top5Similar.length} SP tương tự (L2) đạt chuẩn từ SP gốc ${itemId}`, 'success');
                    }
                    GM_setValue(gmKey('skipped_cache'), skippedCache);
                }

                return productsToReturn;
            } else {
                addLog(`API trả về mã lỗi ${res.code} cho SP ${itemId}.`, 'error');
                if (!skippedCache.includes(String(itemId))) {
                    skippedCache.push(String(itemId));
                    GM_setValue(gmKey('skipped_cache'), skippedCache);
                    updateCacheCount();
                }
            }
        } catch (err) {
            addLog(`Lỗi parse gói tin API SP ${itemId}: ${err.message}`, 'error');
            if (!skippedCache.includes(String(itemId))) {
                skippedCache.push(String(itemId));
                GM_setValue(gmKey('skipped_cache'), skippedCache);
                updateCacheCount();
            }
        }
        return null;
    }

    function addLog(msg, type = 'info') {
        const logBox = document.getElementById(elId('logs'));
        if (!logBox) {
            console.log(`[Shopee Scraper] [${type}] ${msg}`);
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

    function setStatus(text, color = '#aaa') {
        const statusEl = document.getElementById(elId('status'));
        if (statusEl) {
            statusEl.innerText = text;
            statusEl.style.backgroundColor = color;
            statusEl.style.color = '#fff';
        }
    }

    function updateCacheCount() {
        const cacheEl = document.getElementById(elId('cache-count'));
        if (cacheEl) cacheEl.innerText = crawledCache ? crawledCache.length : 0;
        const skipEl = document.getElementById(elId('skip-count'));
        if (skipEl) skipEl.innerText = skippedCache ? skippedCache.length : 0;
    }

    function detectCaptchaPage() {
        const href = window.location.href;
        const isCaptchaUrl = href.includes('/verify/captcha') || href.includes('/verify/traffic') || href.includes('/verify/security') || href.includes('verify/traffic');
        const hasCaptchaDom = !!document.querySelector('.check-captcha-box') || 
                             !!document.querySelector('.shopee-captcha-wrapper') || 
                             !!document.getElementById('captcha-submit') ||
                             !!document.querySelector('iframe[src*="captcha"]');
        return isCaptchaUrl || hasCaptchaDom;
    }

    function playWarningBeep() {
        try {
            const audioCtx = new (window.AudioContext || window.webkitAudioContext)();
            const oscillator = audioCtx.createOscillator();
            const gainNode = audioCtx.createGain();
            oscillator.connect(gainNode);
            gainNode.connect(audioCtx.destination);
            oscillator.type = 'sine';
            oscillator.frequency.setValueAtTime(880, audioCtx.currentTime);
            gainNode.gain.setValueAtTime(0.2, audioCtx.currentTime);
            oscillator.start();
            oscillator.stop(audioCtx.currentTime + 0.3);
        } catch (e) {
            console.log('Audio error:', e);
        }
    }

    function showCaptchaOverlay(message) {
        if (document.getElementById(elId('captcha-overlay-alert'))) return;
        
        const overlay = document.createElement('div');
        overlay.id = elId('captcha-overlay-alert');
        overlay.style.cssText = `position: fixed; top: 0; left: 0; width: 100%; height: 100%; background-color: rgba(238, 77, 45, 0.95); z-index: 9999999; display: flex; flex-direction: column; justify-content: center; align-items: center; color: white; font-family: Arial, sans-serif; text-align: center; padding: 20px; box-sizing: border-box;`;
        overlay.innerHTML = `
            <div style="font-size: 60px; margin-bottom: 20px; animation: blinker 1s linear infinite;">⚠️ CAPTCHA DETECTED! ⚠️</div>
            <div style="font-size: 24px; font-weight: bold; margin-bottom: 20px; max-width: 800px;">${message}</div>
            <div style="font-size: 18px; color: #ffeb3b; border: 2px dashed #ffeb3b; padding: 15px; border-radius: 8px; margin-bottom: 20px;">
                Hãy giải Captcha trên tab phụ đang mở (hoặc ngay trên trang này).<br>
                Sau khi giải xong, tool sẽ <b>TỰ ĐỘNG</b> tiếp tục. Nút bên dưới chỉ dùng khi tool không tự nhận diện.
            </div>
            <button id="${elId('btn-resume-captcha')}" style="padding: 15px 30px; border: none; border-radius: 8px; background-color: #4caf50; color: white; font-size: 18px; font-weight: bold; cursor: pointer; box-shadow: 0 4px 10px rgba(0,0,0,0.3);">
                Tiếp tục cào (thủ công)
            </button>
            <style>
                @keyframes blinker { 50% { opacity: 0; } }
            </style>
        `;
        document.body.appendChild(overlay);
        
        document.getElementById(elId('btn-resume-captcha')).addEventListener('click', () => {
            GM_setValue(gmKey('global_captcha_detected'), false);
            hideCaptchaOverlay();
        });
    }

    function hideCaptchaOverlay() {
        const overlay = document.getElementById(elId('captcha-overlay-alert'));
        if (overlay) overlay.remove();
    }

    async function checkAndWaitForCaptcha() {
        let hasShownOverlay = false;
        let lastBeepTime = 0;
        
        while (isRunning) {
            const isMainCaptcha = detectCaptchaPage();
            const isGlobalCaptcha = GM_getValue(gmKey('global_captcha_detected')) === true;
            
            if (isMainCaptcha || isGlobalCaptcha) {
                if (!hasShownOverlay) {
                    const msg = isMainCaptcha 
                        ? "Phát hiện CAPTCHA xác minh robot trên Trang chính Shopee!" 
                        : "Phát hiện CAPTCHA trên Trang liên kết Affiliate (tab phụ)!";
                    addLog(`[CẢNH BÁO] ${msg} Tạm dừng cào để giải captcha...`, 'error');
                    setStatus('CAPTCHA!', '#f44336');
                    showCaptchaOverlay(msg);
                    hasShownOverlay = true;
                }
                
                const now = Date.now();
                if (now - lastBeepTime > 5000) {
                    playWarningBeep();
                    lastBeepTime = now;
                }
                await sleep(1000);
            } else {
                if (hasShownOverlay) {
                    addLog('Đã xác minh Captcha! Tiếp tục cào...', 'success');
                    setStatus('Running...', '#4caf50');
                    hideCaptchaOverlay();
                }
                break;
            }
        }
    }

    function updateMergedLinks(dataList) {
        if (!dataList || dataList.length === 0) return;
        const groupLinksMap = {};
        for (const item of dataList) {
            const groupId = item[C.group];
            const link = item[C.url];
            if (groupId && link) {
                if (!groupLinksMap[groupId]) {
                    groupLinksMap[groupId] = [];
                }
                if (!groupLinksMap[groupId].includes(link)) {
                    groupLinksMap[groupId].push(link);
                }
            }
        }

        for (const item of dataList) {
            const groupId = item[C.group];
            if (groupId && groupLinksMap[groupId]) {
                item[C.mergedLink] = groupLinksMap[groupId].join('|');
            } else {
                item[C.mergedLink] = '';
            }
        }
    }

    function getFilters() {
        const isAff = HOST.startsWith('affiliate.');
        if (isAff) {
            const commEl = document.getElementById(elId('comm-min'));
            const soldEl = document.getElementById(elId('sold-min'));
            const comEl = document.getElementById(elId('com-min'));
            return {
                commMin: commEl ? (parseFloat(commEl.value) || 0) : (SITE.defaults.commMin || 0),
                soldMin: soldEl ? (parseInt(soldEl.value, 10) || 0) : (SITE.defaults.soldMin || 0),
                sellerComMin: comEl ? (parseFloat(comEl.value) || 0) : (SITE.defaults.sellerCommissionMin || 0)
            };
        } else {
            const priceEl = document.getElementById(elId('price-min'));
            const soldEl = document.getElementById(elId('sold-min'));
            const comEl = document.getElementById(elId('com-min'));
            return {
                priceMin: priceEl ? (parseFloat(priceEl.value) || 0) : (SITE.defaults.priceMin || 0),
                soldMin: soldEl ? (parseInt(soldEl.value, 10) || 0) : (SITE.defaults.soldMin || 0),
                sellerComMin: comEl ? (parseFloat(comEl.value) || 0) : (SITE.defaults.sellerCommissionMin || 0)
            };
        }
    }

    function showDoneOverlay() {
        const existing = document.getElementById(`${NS}-done-overlay`);
        if (existing) existing.remove();

        const overlay = document.createElement('div');
        overlay.id = `${NS}-done-overlay`;
        overlay.style.cssText = `position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; background-color: rgba(46,125,50,0.15); backdrop-filter: blur(6px); -webkit-backdrop-filter: blur(6px); z-index: 9999999; display: flex; justify-content: center; align-items: center; transition: opacity 0.4s ease; opacity: 0;`;

        const card = document.createElement('div');
        card.style.cssText = `background-color: #1a1a1a; border: 2px solid #4caf50; border-radius: 12px; padding: 35px 60px; box-shadow: 0 10px 30px rgba(0,0,0,0.6); text-align: center; color: #fff; font-family: Arial, sans-serif; transform: scale(0.8); transition: transform 0.3s cubic-bezier(0.175,0.885,0.32,1.275);`;
        card.innerHTML = `
            <div style="font-size: 50px; color: #4caf50; margin-bottom: 12px; animation: ${NS}-bounce 0.8s infinite alternate;">🎉</div>
            <div style="font-size: 32px; font-weight: bold; margin-bottom: 8px; color: #4caf50; letter-spacing: 1px;">Done!</div>
            <div style="font-size: 14px; color: #ccc; margin-bottom: 25px;">Đã hoàn thành cào dữ liệu và đồng bộ (${SITE.label}).</div>
            <button id="${NS}-btn-close-overlay" style="padding: 10px 30px; font-size: 14px; font-weight: bold; color: white; background-color: #4caf50; border: none; border-radius: 6px; cursor: pointer; transition: background-color 0.2s; box-shadow: 0 4px 10px rgba(76,175,80,0.3);">OK</button>
        `;

        const style = document.createElement('style');
        style.id = `${NS}-overlay-style`;
        style.innerHTML = `
            @keyframes ${NS}-bounce { from { transform: translateY(0); } to { transform: translateY(-10px); } }
            #${NS}-btn-close-overlay:hover { background-color: #388e3c !important; }
        `;
        document.head.appendChild(style);

        overlay.appendChild(card);
        document.body.appendChild(overlay);

        setTimeout(() => {
            overlay.style.opacity = '1';
            card.style.transform = 'scale(1)';
        }, 50);

        card.querySelector(`#${NS}-btn-close-overlay`).addEventListener('click', () => {
            overlay.style.opacity = '0';
            card.style.transform = 'scale(0.8)';
            setTimeout(() => {
                overlay.remove();
                const styleEl = document.getElementById(`${NS}-overlay-style`);
                if (styleEl) styleEl.remove();
            }, 400);
        });
    }

    // ============================================================
    // 2. KỊCH BẢN A: CÀO SẢN PHẨM THEO DANH MỤC (RUNS ON SHOPEE.*)
    // ============================================================
    // ============================================================
    // 2. KỊCH BẢN A: CÀO SẢN PHẨM TRÊN TRANG SHOPEE CHÍNH (DANH MỤC & TỪ KHÓA)
    // ============================================================
    function runMainShopeeScraper() {
        let crawledData = [];
        let keywordsList = [];
        let currentKeywordIndex = 0;
        let isTopSales = false;

        const panel = document.createElement('div');
        panel.id = PANEL_ID;
        panel.style.cssText = `position: fixed; top: 80px; right: 20px; z-index: 999999; background-color: #1a1a1a; color: #fff; padding: 15px; border-radius: 8px; box-shadow: 0 4px 15px rgba(0,0,0,0.5); font-family: Arial, sans-serif; font-size: 13px; width: 300px; border: 1px solid #ee4d2d;`;

        let googleConfig = GM_getValue(gmKey('google_config')) || { sheetUrl: '', tabName: 'Sheet1', clientId: '', clientSecret: '' };
        let videoaiConfig = GM_getValue(gmKey('videoai_config')) || { apiKey: '', endpoint: VIDEOAI_DEFAULT_ENDPOINT, autoPush: false };
        let filterConfig = GM_getValue(`${NS}_filter_config_${SITE.code}`) || {
            priceMin: SITE.defaults.priceMin,
            soldMin: SITE.defaults.soldMin,
            sellerCommissionMin: SITE.defaults.sellerCommissionMin,
            selectedCategory: 'no-category',
            selectedMode: 'category',
            keywordsText: '',
            linksText: ''
        };

        let sheetConfig = {
            sheetUrl: googleConfig.sheetUrl || '',
            tabName: googleConfig.tabName || 'Sheet1',
            clientId: googleConfig.clientId || '',
            clientSecret: googleConfig.clientSecret || '',
            priceMin: filterConfig.priceMin,
            soldMin: filterConfig.soldMin,
            sellerCommissionMin: filterConfig.sellerCommissionMin,
            selectedCategory: filterConfig.selectedCategory || 'no-category',
            selectedMode: filterConfig.selectedMode || 'category',
            keywordsText: filterConfig.keywordsText || '',
            linksText: filterConfig.linksText || ''
        };

        function isTopSalesSelected() {
            return window.location.search.includes('sortBy=sales') || isTopSales;
        }

        function saveRunningState(step) {
            GM_setValue(gmKey('running_state'), {
                isRunning: isRunning,
                selectedMode: sheetConfig.selectedMode,
                crawledData: crawledData,
                currentStep: step,
                selectedCategory: sheetConfig.selectedCategory,
                keywordsList: keywordsList,
                currentKeywordIndex: currentKeywordIndex,
                isTopSales: isTopSales
            });
        }

        function updateKeywordProgressUI() {
            const wrapper = document.getElementById(elId('keyword-progress-wrapper'));
            if (!wrapper) return;

            const isKeywordMode = sheetConfig.selectedMode === 'keyword';
            wrapper.style.display = (isKeywordMode && isRunning) ? 'block' : 'none';

            if (isKeywordMode && isRunning && keywordsList.length > 0) {
                const curDisplay = document.getElementById(elId('current-keyword-display'));
                if (curDisplay) {
                    const filterText = isTopSales ? ` (${SITE.topSalesLabel})` : ' (Liên quan)';
                    curDisplay.innerText = `${keywordsList[currentKeywordIndex] || '-'}${filterText}`;
                }
                const ratioDisplay = document.getElementById(elId('keyword-progress-ratio'));
                if (ratioDisplay) {
                    ratioDisplay.innerText = `${currentKeywordIndex + 1}/${keywordsList.length}`;
                }
            }
        }

        function checkAndResumeState() {
            try {
                const savedState = GM_getValue(gmKey('running_state'));
                if (savedState && savedState.selectedMode === 'link') {
                    // Cào theo link chạy trong bộ nhớ, không khôi phục qua reload
                    GM_setValue(gmKey('running_state'), null);
                    return;
                }
                if (savedState && savedState.isRunning) {
                    addLog('Khôi phục phiên cào trước đó đang chạy dở...', 'warn');
                    isRunning = true;
                    sheetConfig.selectedMode = savedState.selectedMode || 'category';
                    crawledData = savedState.crawledData || [];
                    
                    const modeSelect = document.getElementById(elId('run-mode-select'));
                    if (modeSelect) {
                        modeSelect.value = sheetConfig.selectedMode;
                        document.getElementById(elId('category-select-wrapper')).style.display = sheetConfig.selectedMode === 'category' ? 'block' : 'none';
                        document.getElementById(elId('keywords-input-wrapper')).style.display = sheetConfig.selectedMode === 'keyword' ? 'block' : 'none';
                    }

                    if (sheetConfig.selectedMode === 'keyword') {
                        keywordsList = savedState.keywordsList || [];
                        currentKeywordIndex = savedState.currentKeywordIndex || 0;
                        isTopSales = !!savedState.isTopSales;
                    }

                    syncButtonsRunning(true);
                    setStatus('Resuming...', '#ff9800');

                    const countEl = document.getElementById(elId('count'));
                    if (countEl) countEl.innerText = crawledData.length;

                    updateKeywordProgressUI();

                    setTimeout(async () => {
                        if (sheetConfig.selectedMode === 'category') {
                            if (savedState.currentStep === 'waiting_topsales') {
                                await transitionToTopSalesAndCrawl();
                            } else {
                                await crawlCurrentPage();
                            }
                        } else {
                            await checkKeywordPageAndCrawl();
                        }
                    }, 3000);
                }
            } catch (e) {
                console.error('Lỗi khôi phục chạy:', e);
            }
        }

        function syncButtonsRunning(running) {
            const startBtn = document.getElementById(elId('btn-start'));
            const stopBtn = document.getElementById(elId('btn-stop'));
            if (startBtn) {
                startBtn.disabled = running;
                startBtn.style.backgroundColor = running ? '#555' : '#ee4d2d';
            }
            if (stopBtn) {
                stopBtn.disabled = !running;
                stopBtn.style.backgroundColor = running ? '#ee4d2d' : '#555';
            }
        }

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
                    addLog(`Chuyển hướng đến danh mục con: ${aTag.href}`, 'info');
                    window.location.href = aTag.href;
                    return true;
                }
                el.click();
                return true;
            }
            return false;
        }

        function getElementByXPath(xpath) {
            return document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
        }

        function initPanel() {
            if (document.getElementById(PANEL_ID)) return;

            panel.innerHTML = `
                <div style="font-weight: bold; font-size: 15px; margin-bottom: 12px; color: #ee4d2d; border-bottom: 1px solid #333; padding-bottom: 5px; display: flex; justify-content: space-between; align-items: center;">
                    <span>Shopee Main Scraper</span>
                    <span id="${elId('status')}" style="font-size: 10px; padding: 2px 6px; background-color: #333; border-radius: 4px; color: #aaa;">Idle</span>
                </div>

                <div style="font-size: 10px; color: #888; margin-bottom: 10px;">Thị trường: <strong style="color:#ee4d2d;">${SITE.label}</strong> · Tiền tệ: ${SITE.currency}</div>

                <!-- Chọn chế độ cào -->
                <div style="margin-bottom: 10px;">
                    <label style="display: block; margin-bottom: 4px; font-weight: bold; font-size: 11px;">Chế độ cào:</label>
                    <select id="${elId('run-mode-select')}" style="width: 98%; padding: 5px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white; font-size: 12px;">
                        <option value="category" ${sheetConfig.selectedMode === 'category' ? 'selected' : ''}>Cào theo danh mục</option>
                        <option value="keyword" ${sheetConfig.selectedMode === 'keyword' ? 'selected' : ''}>Cào theo từ khóa</option>
                        <option value="link" ${sheetConfig.selectedMode === 'link' ? 'selected' : ''}>Cào theo link</option>
                    </select>
                </div>

                <!-- Bộ lọc -->
                <div style="margin-bottom: 10px; display: flex; gap: 8px;">
                    <div style="flex: 1; display: flex; flex-direction: column;">
                        <label style="margin-bottom: 4px; font-weight: bold; font-size: 11px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; color: #ccc;" title="Giá tối thiểu (${SITE.currency})">Giá Min (${SITE.currency}):</label>
                        <input type="number" step="${SITE.defaults.priceStep}" id="${elId('price-min')}" value="${sheetConfig.priceMin}" style="width: 100%; padding: 5px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white; font-size: 12px; box-sizing: border-box;">
                    </div>
                    <div style="flex: 1; display: flex; flex-direction: column;">
                        <label style="margin-bottom: 4px; font-weight: bold; font-size: 11px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; color: #ccc;" title="Đã bán tối thiểu">Đã bán Min:</label>
                        <input type="number" id="${elId('sold-min')}" value="${sheetConfig.soldMin}" style="width: 100%; padding: 5px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white; font-size: 12px; box-sizing: border-box;">
                    </div>
                    <div style="flex: 1; display: flex; flex-direction: column;">
                        <label style="margin-bottom: 4px; font-weight: bold; font-size: 11px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; color: #ccc;" title="Hoa hồng người bán tối thiểu (${SITE.currency})">HH Min (${SITE.currency}):</label>
                        <input type="number" step="${SITE.defaults.comStep}" id="${elId('com-min')}" value="${sheetConfig.sellerCommissionMin}" style="width: 100%; padding: 5px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white; font-size: 12px; box-sizing: border-box;">
                    </div>
                </div>

                <!-- Chọn danh mục (Category select wrapper) -->
                <div id="${elId('category-select-wrapper')}" style="margin-bottom: 10px; display: ${sheetConfig.selectedMode === 'category' ? 'block' : 'none'};">
                    <label style="display: block; margin-bottom: 4px; font-weight: bold; font-size: 11px;">Chế độ cào nhóm ngành:</label>
                    <select id="${elId('category-select')}" style="width: 98%; padding: 5px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white; font-size: 12px;">
                        <option value="no-category" ${sheetConfig.selectedCategory === 'no-category' ? 'selected' : ''}>No Category (Chỉ cào trang hiện tại)</option>
                        ${Object.keys(SITE.categories).map(cat => `<option value="${cat}" ${sheetConfig.selectedCategory === cat ? 'selected' : ''}>${cat}</option>`).join('')}
                    </select>
                </div>

                <!-- Từ khóa tìm kiếm (Keyword textarea wrapper) -->
                <div id="${elId('keywords-input-wrapper')}" style="margin-bottom: 10px; display: ${sheetConfig.selectedMode === 'keyword' ? 'block' : 'none'};">
                    <label style="display: block; margin-bottom: 4px; font-weight: bold; font-size: 11px; color: #ccc;">Từ khóa tìm kiếm (mỗi từ khóa 1 dòng):</label>
                    <textarea id="${elId('keywords-input')}" rows="4" placeholder="Nhập danh sách từ khóa..." style="width: 96%; padding: 6px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white; font-size: 11px; font-family: monospace; resize: vertical; box-sizing: border-box;">${sheetConfig.keywordsText}</textarea>
                </div>

                <!-- Danh sách link sản phẩm (Link textarea wrapper) -->
                <div id="${elId('links-input-wrapper')}" style="margin-bottom: 10px; display: ${sheetConfig.selectedMode === 'link' ? 'block' : 'none'};">
                    <label style="display: block; margin-bottom: 4px; font-weight: bold; font-size: 11px; color: #ccc;">Danh sách link sản phẩm (mỗi link 1 dòng):</label>
                    <textarea id="${elId('links-input')}" rows="5" placeholder="https://shopee.vn/product/714557060/20969872164" style="width: 96%; padding: 6px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white; font-size: 11px; font-family: monospace; resize: vertical; box-sizing: border-box;">${sheetConfig.linksText}</textarea>
                    <div id="${elId('link-progress')}" style="margin-top: 6px; font-size: 11px; color: #4caf50; display: none;">Tiến độ link: <strong id="${elId('link-progress-ratio')}">0/0</strong></div>
                </div>

                <!-- Cấu hình Google Sheets -->
                <div style="margin-bottom: 10px; border: 1px solid #333; border-radius: 4px;">
                    <div id="${elId('sheet-toggle')}" style="background-color: #2b2b2b; padding: 6px 10px; font-weight: bold; cursor: pointer; display: flex; justify-content: space-between; align-items: center; border-radius: 4px 4px 0 0;">
                        <span>⚙️ Cấu hình Google Sheets</span>
                        <span id="${elId('toggle-arrow')}">▲</span>
                    </div>
                    <div id="${elId('sheet-fields')}" style="padding: 10px; display: block; background-color: #1f1f1f;">
                        <div style="margin-bottom: 8px;">
                            <label style="display: block; margin-bottom: 3px; font-size: 11px;">Đường dẫn Spreadsheet:</label>
                            <input type="text" id="${elId('sheet-url')}" placeholder="https://docs.google.com/spreadsheets/d/..." value="${sheetConfig.sheetUrl}" style="width: 95%; padding: 5px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white; font-size: 11px;">
                        </div>
                        <div style="margin-bottom: 8px;">
                            <label style="display: block; margin-bottom: 3px; font-size: 11px;">Tên Sheet Tab:</label>
                            <input type="text" id="${elId('sheet-tab')}" placeholder="Sheet1" value="${sheetConfig.tabName}" style="width: 95%; padding: 5px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white; font-size: 11px;">
                        </div>
                        <div style="margin-bottom: 8px;">
                            <label style="display: block; margin-bottom: 3px; font-size: 11px;">Google Client ID:</label>
                            <input type="password" id="${elId('client-id')}" placeholder="Client ID" value="${sheetConfig.clientId}" style="width: 95%; padding: 5px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white; font-size: 11px;">
                        </div>
                        <div style="margin-bottom: 8px;">
                            <label style="display: block; margin-bottom: 3px; font-size: 11px;">Google Client Secret:</label>
                            <input type="password" id="${elId('client-secret')}" placeholder="Client Secret" value="${sheetConfig.clientSecret}" style="width: 95%; padding: 5px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white; font-size: 11px;">
                        </div>
                        <div style="display: flex; gap: 5px; justify-content: flex-end; margin-bottom: 8px; flex-wrap: wrap;">
                            <button id="${elId('btn-save-sheet')}" style="padding: 4px 6px; border: none; border-radius: 3px; background-color: #0288d1; color: white; cursor: pointer; font-size: 11px;">Lưu cấu hình</button>
                            <button id="${elId('btn-connect-google')}" style="padding: 4px 6px; border: none; border-radius: 3px; background-color: #e65100; color: white; cursor: pointer; font-size: 11px; font-weight: bold;">Kết nối Google</button>
                            <button id="${elId('btn-test-sheet')}" style="padding: 4px 6px; border: none; border-radius: 3px; background-color: #6a1b9a; color: white; cursor: pointer; font-size: 11px; font-weight: bold;">Test Sheet</button>
                        </div>

                        <div id="${elId('auth-code-wrapper')}" style="border-top: 1px dashed #444; padding-top: 8px; display: none;">
                            <label style="display: block; margin-bottom: 3px; font-size: 11px; color: #ffeb3b;">Nhập mã Authorization Code:</label>
                            <div style="display: flex; gap: 5px;">
                                <input type="text" id="${elId('auth-code-input')}" placeholder="Dán mã code tại đây" style="flex: 1; padding: 4px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white; font-size: 11px;">
                                <button id="${elId('btn-confirm-code')}" style="padding: 4px 8px; border: none; border-radius: 3px; background-color: #4caf50; color: white; cursor: pointer; font-size: 11px; font-weight: bold;">Xác nhận</button>
                            </div>
                        </div>
                    </div>
                </div>

                <!-- Cấu hình VideoAI API -->
                <div style="margin-bottom: 10px; border: 1px solid #333; border-radius: 4px;">
                    <div id="${elId('videoai-toggle')}" style="background-color: #2b2b2b; padding: 6px 10px; font-weight: bold; cursor: pointer; display: flex; justify-content: space-between; align-items: center; border-radius: 4px 4px 0 0;">
                        <span>🚀 Cấu hình VideoAI API</span>
                        <span id="${elId('videoai-toggle-arrow')}">▼</span>
                    </div>
                    <div id="${elId('videoai-fields')}" style="padding: 10px; display: none; background-color: #1f1f1f;">
                        <div style="margin-bottom: 8px;">
                            <label style="display: block; margin-bottom: 3px; font-size: 11px;">VideoAI API Key:</label>
                            <input type="password" id="${elId('videoai-key')}" placeholder="api_xxxxxxxx" value="${videoaiConfig.apiKey}" style="width: 95%; padding: 5px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white; font-size: 11px;">
                        </div>
                        <div style="margin-bottom: 8px;">
                            <label style="display: block; margin-bottom: 3px; font-size: 11px;">Endpoint:</label>
                            <input type="text" id="${elId('videoai-endpoint')}" placeholder="${VIDEOAI_DEFAULT_ENDPOINT}" value="${videoaiConfig.endpoint}" style="width: 95%; padding: 5px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white; font-size: 11px;">
                        </div>
                        <div style="margin-bottom: 8px; display: flex; align-items: center; gap: 6px;">
                            <input type="checkbox" id="${elId('videoai-autopush')}" ${videoaiConfig.autoPush ? 'checked' : ''} style="cursor: pointer;">
                            <label for="${elId('videoai-autopush')}" style="font-size: 11px; cursor: pointer;">Tự động đẩy lên VideoAI sau mỗi đợt cào</label>
                        </div>
                        <div style="display: flex; gap: 5px; justify-content: flex-end; flex-wrap: wrap;">
                            <button id="${elId('btn-save-videoai')}" style="padding: 4px 6px; border: none; border-radius: 3px; background-color: #0288d1; color: white; cursor: pointer; font-size: 11px;">Lưu cấu hình</button>
                            <button id="${elId('btn-test-videoai')}" style="padding: 4px 6px; border: none; border-radius: 3px; background-color: #6a1b9a; color: white; cursor: pointer; font-size: 11px; font-weight: bold;">Test kết nối</button>
                            <button id="${elId('btn-push-videoai')}" style="padding: 4px 6px; border: none; border-radius: 3px; background-color: #ee4d2d; color: white; cursor: pointer; font-size: 11px; font-weight: bold;">Đẩy lên VideoAI</button>
                        </div>
                    </div>
                </div>

                <!-- Khung hiển thị tiến độ cào từ khóa (Chỉ hiện ở chế độ keyword khi đang chạy) -->
                <div id="${elId('keyword-progress-wrapper')}" style="margin-bottom: 10px; background-color: #2b2b2b; padding: 8px; border-radius: 4px; border: 1px solid #ee4d2d; display: none;">
                    <div style="font-weight: bold; font-size: 11px; color: #ee4d2d; margin-bottom: 4px;">Tiến độ cào từ khóa:</div>
                    <div style="display: flex; justify-content: space-between; font-size: 12px; color: #fff;">
                        <span style="white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 170px;">
                            Từ khóa: <strong id="${elId('current-keyword-display')}" style="color: #ffeb3b;">-</strong>
                        </span>
                        <span>
                            Tiến độ: <strong id="${elId('keyword-progress-ratio')}" style="color: #4caf50;">0/0</strong>
                        </span>
                    </div>
                </div>

                <div style="margin-bottom: 12px; display: flex; gap: 5px; justify-content: space-between; align-items: center; background-color: #2b2b2b; padding: 6px; border-radius: 4px;">
                    <span>Đã cào: <strong id="${elId('count')}" style="color: #ee4d2d;">0</strong> SP</span>
                    <span style="font-size: 11px; white-space: nowrap;">Cache: <strong id="${elId('cache-count')}" style="color: #4caf50;">0</strong> | Skip: <strong id="${elId('skip-count')}" style="color: #f44336;">0</strong></span>
                </div>

                <div style="display: flex; gap: 5px; margin-bottom: 8px;">
                    <button id="${elId('btn-start')}" style="flex: 1; padding: 8px; border: none; border-radius: 4px; background-color: #ee4d2d; color: white; font-weight: bold; cursor: pointer;">Start Crawl</button>
                    <button id="${elId('btn-stop')}" style="flex: 1; padding: 8px; border: none; border-radius: 4px; background-color: #555; color: white; font-weight: bold; cursor: pointer;" disabled>Stop</button>
                </div>
                <div style="display: flex; gap: 5px; margin-bottom: 8px;">
                    <button id="${elId('btn-clear-cache')}" style="width: 33%; padding: 6px; border: none; border-radius: 4px; background-color: #444; color: #ccc; cursor: pointer; font-size: 11px;">Xóa Cache</button>
                    <button id="${elId('btn-sync-cache')}" style="width: 34%; padding: 6px; border: none; border-radius: 4px; background-color: #0288d1; color: white; cursor: pointer; font-size: 11px; font-weight: bold;">Đồng bộ Cache</button>
                    <button id="${elId('btn-export')}" style="width: 33%; padding: 6px; border: none; border-radius: 4px; background-color: #2e7d32; color: white; cursor: pointer; font-size: 11px; font-weight: bold;">Xuất Excel</button>
                </div>
                <div style="margin-top: 10px;">
                    <div style="font-weight: bold; margin-bottom: 4px;">System Logs:</div>
                    <div id="${elId('logs')}" style="height: 120px; overflow-y: auto; background-color: #0c0c0c; border: 1px solid #333; padding: 6px; font-family: monospace; font-size: 11px; border-radius: 4px; word-break: break-all;">
                        <div style="color: #888;">Chờ lệnh Start...</div>
                    </div>
                </div>
            `;
            document.body.appendChild(panel);

            // Gắn sự kiện giao diện
            document.getElementById(elId('sheet-toggle')).addEventListener('click', () => {
                const fields = document.getElementById(elId('sheet-fields'));
                const arrow = document.getElementById(elId('toggle-arrow'));
                if (fields.style.display === 'none') {
                    fields.style.display = 'block';
                    arrow.innerText = '▲';
                } else {
                    fields.style.display = 'none';
                    arrow.innerText = '▼';
                }
            });

            // Lắng nghe đổi chế độ cào
            const modeSelect = document.getElementById(elId('run-mode-select'));
            if (modeSelect) {
                modeSelect.addEventListener('change', (e) => {
                    const mode = e.target.value;
                    sheetConfig.selectedMode = mode;
                    document.getElementById(elId('category-select-wrapper')).style.display = mode === 'category' ? 'block' : 'none';
                    document.getElementById(elId('keywords-input-wrapper')).style.display = mode === 'keyword' ? 'block' : 'none';
                    document.getElementById(elId('links-input-wrapper')).style.display = mode === 'link' ? 'block' : 'none';

                    saveFilterConfig();
                    updateKeywordProgressUI();
                    const modeLabel = mode === 'category' ? 'Cào danh mục' : (mode === 'keyword' ? 'Cào từ khóa' : 'Cào theo link');
                    addLog(`Đã chuyển chế độ cào sang: ${modeLabel}`, 'info');
                });
            }

            document.getElementById(elId('btn-connect-google')).addEventListener('click', () => {
                readSheetInputsToConfig();
                GM_setValue(gmKey('google_config'), googleConfig);
                startGoogleAuthFlow(googleConfig.clientId);
                document.getElementById(elId('auth-code-wrapper')).style.display = 'block';
            });

            document.getElementById(elId('btn-confirm-code')).addEventListener('click', () => {
                const code = document.getElementById(elId('auth-code-input')).value.trim();
                if (!code) {
                    alert('Hãy dán mã Authorization Code!');
                    return;
                }
                exchangeAuthCodeForTokens(code, googleConfig.clientId, googleConfig.clientSecret);
                document.getElementById(elId('auth-code-input')).value = '';
                document.getElementById(elId('auth-code-wrapper')).style.display = 'none';
            });

            document.getElementById(elId('btn-test-sheet')).addEventListener('click', async () => {
                readSheetInputsToConfig();
                GM_setValue(gmKey('google_config'), googleConfig);
                await testGoogleSheetsConnection();
            });

            document.getElementById(elId('btn-save-sheet')).addEventListener('click', () => {
                readSheetInputsToConfig();
                saveFilterConfig();
                GM_setValue(gmKey('google_config'), googleConfig);
                addLog('Đã lưu cấu hình và bộ lọc!', 'success');
                alert('Lưu cấu hình thành công!');
            });

            document.getElementById(elId('btn-start')).addEventListener('click', startScraper);
            document.getElementById(elId('btn-stop')).addEventListener('click', stopScraper);
            document.getElementById(elId('btn-sync-cache')).addEventListener('click', syncCacheFromGoogleSheets);

            document.getElementById(elId('btn-clear-cache')).addEventListener('click', () => {
                if (confirm(`Bạn có chắc chắn muốn xóa bộ nhớ đệm cache của ${SITE.label} không?`)) {
                    crawledCache = [];
                    skippedCache = [];
                    GM_setValue(gmKey('crawled_cache'), crawledCache);
                    GM_setValue(gmKey('skipped_cache'), skippedCache);
                    updateCacheCount();
                    addLog('Đã xóa sạch bộ nhớ đệm cache thị trường hiện tại.', 'warn');
                    alert('Xóa cache thành công!');
                }
            });

            document.getElementById(elId('btn-export')).addEventListener('click', () => {
                updateMergedLinks(crawledData);
                exportExcelForData(crawledData, `${sheetConfig.selectedMode}_crawled`);
            });

            document.getElementById(elId('videoai-toggle')).addEventListener('click', () => {
                const fields = document.getElementById(elId('videoai-fields'));
                const arrow = document.getElementById(elId('videoai-toggle-arrow'));
                if (fields.style.display === 'none') {
                    fields.style.display = 'block';
                    arrow.innerText = '▲';
                } else {
                    fields.style.display = 'none';
                    arrow.innerText = '▼';
                }
            });

            document.getElementById(elId('btn-save-videoai')).addEventListener('click', () => {
                readVideoAIInputsToConfig();
                GM_setValue(gmKey('videoai_config'), videoaiConfig);
                addLog('Đã lưu cấu hình VideoAI API!', 'success');
                alert('Lưu cấu hình VideoAI thành công!');
            });

            document.getElementById(elId('btn-test-videoai')).addEventListener('click', async () => {
                readVideoAIInputsToConfig();
                GM_setValue(gmKey('videoai_config'), videoaiConfig);
                await testVideoAIConnection();
            });

            document.getElementById(elId('btn-push-videoai')).addEventListener('click', async () => {
                readVideoAIInputsToConfig();
                GM_setValue(gmKey('videoai_config'), videoaiConfig);
                if (!crawledData || crawledData.length === 0) {
                    alert('Chưa có sản phẩm nào để đẩy!');
                    return;
                }
                updateMergedLinks(crawledData);
                addLog(`Đẩy thủ công ${crawledData.length} SP lên VideoAI...`, 'info');
                try {
                    await pushToVideoAI(crawledData);
                } catch (err) {
                    addLog(`Lỗi đẩy VideoAI: ${err.message}`, 'error');
                    alert(`Lỗi đẩy VideoAI: ${err.message}`);
                }
            });

            const categorySelect = document.getElementById(elId('category-select'));
            if (categorySelect) {
                categorySelect.addEventListener('change', (e) => {
                    sheetConfig.selectedCategory = e.target.value;
                    saveFilterConfig();
                    addLog(`Chuyển chế độ cào danh mục: ${e.target.value === 'no-category' ? 'Chỉ cào trang hiện tại' : e.target.value}`, 'info');
                });
            }

            updateCacheCount();
            checkAndResumeState();
        }

        function readSheetInputsToConfig() {
            googleConfig.sheetUrl = document.getElementById(elId('sheet-url')).value.trim();
            googleConfig.tabName = document.getElementById(elId('sheet-tab')).value.trim();
            googleConfig.clientId = document.getElementById(elId('client-id')).value.trim();
            googleConfig.clientSecret = document.getElementById(elId('client-secret')).value.trim();
        }

        function readVideoAIInputsToConfig() {
            videoaiConfig.apiKey = document.getElementById(elId('videoai-key')).value.trim();
            videoaiConfig.endpoint = document.getElementById(elId('videoai-endpoint')).value.trim() || VIDEOAI_DEFAULT_ENDPOINT;
            videoaiConfig.autoPush = document.getElementById(elId('videoai-autopush')).checked;
        }

        function saveFilterConfig() {
            sheetConfig.priceMin = parseFloat(document.getElementById(elId('price-min')).value) || 0;
            sheetConfig.soldMin = parseInt(document.getElementById(elId('sold-min')).value, 10) || 0;
            sheetConfig.sellerCommissionMin = parseFloat(document.getElementById(elId('com-min')).value) || 0;
            sheetConfig.selectedCategory = document.getElementById(elId('category-select')).value;
            
            const kwInput = document.getElementById(elId('keywords-input'));
            sheetConfig.keywordsText = kwInput ? kwInput.value : '';

            const linkInput = document.getElementById(elId('links-input'));
            sheetConfig.linksText = linkInput ? linkInput.value : sheetConfig.linksText;

            GM_setValue(`${NS}_filter_config_${SITE.code}`, {
                priceMin: sheetConfig.priceMin,
                soldMin: sheetConfig.soldMin,
                sellerCommissionMin: sheetConfig.sellerCommissionMin,
                selectedCategory: sheetConfig.selectedCategory,
                selectedMode: sheetConfig.selectedMode,
                keywordsText: sheetConfig.keywordsText,
                linksText: sheetConfig.linksText
            });
        }

        // ============================================================
        // CÀO THEO LINK CÓ SẴN
        // ============================================================
        function parseLinksToItems(rawText) {
            const lines = (rawText || '').split('\n').map(l => l.trim()).filter(Boolean);
            const seen = new Set();
            const items = [];
            let invalidCount = 0;
            for (const line of lines) {
                const parsed = parseProductUrl(line);
                if (!parsed) { invalidCount++; continue; }
                if (seen.has(parsed.itemId)) continue;
                seen.add(parsed.itemId);
                items.push(parsed);
            }
            return { items, invalidCount };
        }

        function updateLinkProgressUI(done, total) {
            const wrap = document.getElementById(elId('link-progress'));
            const ratio = document.getElementById(elId('link-progress-ratio'));
            if (wrap) wrap.style.display = total > 0 ? 'block' : 'none';
            if (ratio) ratio.innerText = `${done}/${total}`;
        }

        async function runLinkScrapingLoop() {
            const linkInput = document.getElementById(elId('links-input'));
            const raw = linkInput ? linkInput.value : sheetConfig.linksText;
            const { items, invalidCount } = parseLinksToItems(raw);

            if (invalidCount > 0) addLog(`Bỏ qua ${invalidCount} dòng link không hợp lệ.`, 'warn');

            if (items.length === 0) {
                addLog('Không có link sản phẩm hợp lệ để cào!', 'error');
                alert('Không có link hợp lệ!\nHãy dán link dạng: https://shopee.vn/product/{shopId}/{itemId}');
                isRunning = false;
                syncButtonsRunning(false);
                return;
            }

            addLog(`Bắt đầu cào ${items.length} link sản phẩm (không lọc hoa hồng)...`, 'success');
            setStatus('Đang cào link...', '#2196f3');
            updateLinkProgressUI(0, items.length);

            const BATCH_SIZE = 3;
            for (let i = 0; i < items.length; i += BATCH_SIZE) {
                if (!isRunning) break;
                await checkAndWaitForCaptcha();
                if (!isRunning) break;

                const batch = items.slice(i, i + BATCH_SIZE);
                addLog(`Đang cào nhóm link ${Math.floor(i / BATCH_SIZE) + 1}/${Math.ceil(items.length / BATCH_SIZE)}...`, 'info');

                const promises = batch.map(p => {
                    const itemData = {
                        [C.itemId]: p.itemId,
                        [C.shopId]: p.shopId,
                        [C.name]: '',
                        [C.url]: p.productUrl
                    };
                    // sellerComMin = 0 -> bỏ lọc hoa hồng (cào hết link đã dán)
                    const tabUrl = `https://${SITE.affiliateDomain}/offer/product_offer/${p.itemId}`;
                    return processProduct(itemData, tabUrl, 0, '');
                });
                const results = await Promise.all(promises);

                const batchData = [];
                for (const arr of results) {
                    if (arr && arr.length > 0) {
                        for (const it of arr) {
                            batchData.push(it);
                            crawledData.push(it);
                            const id = String(it[C.itemId]);
                            if (!crawledCache.includes(id)) crawledCache.push(id);
                        }
                    }
                }

                GM_setValue(gmKey('crawled_cache'), crawledCache);
                updateCacheCount();
                const countEl = document.getElementById(elId('count'));
                if (countEl) countEl.innerText = crawledData.length;
                updateLinkProgressUI(Math.min(i + BATCH_SIZE, items.length), items.length);

                if (batchData.length > 0) {
                    updateMergedLinks(batchData);
                    addLog(`Đẩy ${batchData.length} SP lên Google Sheets...`, 'info');
                    try {
                        await pushToGoogleSheets(batchData);
                        addLog('Đồng bộ Google Sheets thành công!', 'success');
                    } catch (err) {
                        addLog(`Lỗi đồng bộ Sheets: ${err.message}`, 'error');
                        exportExcelForData(batchData, 'backup_link');
                    }
                    if (videoaiConfig.autoPush) {
                        addLog(`Tự động đẩy ${batchData.length} SP lên VideoAI...`, 'info');
                        try {
                            await pushToVideoAI(batchData);
                        } catch (err) {
                            addLog(`Lỗi tự động đẩy VideoAI: ${err.message}`, 'error');
                        }
                    }
                }

                await sleep(2000 + Math.random() * 1500);
            }

            if (isRunning) {
                addLog('========================================', 'success');
                addLog(`ĐÃ HOÀN THÀNH CÀO ${items.length} LINK!`, 'success');
                setStatus('Finished!', '#4caf50');
                await stopScraper();
                showDoneOverlay();
            }
        }

        // ============================================================
        // VÒNG LẶP CÀO SẢN PHẨM (PHƯƠNG ÁN B: CÀO ĐÂU MỞ TAB ĐÓ)
        // ============================================================
        async function scrapeVisiblePage() {
            addLog('Đang cuộn xuống đáy trang để tải thêm sản phẩm...', 'info');
            setStatus('Scrolling...', '#ff9800');
            await smoothScrollToBottom();
            await sleep(1500);

            setStatus('Analyzing...', '#2196f3');
            const items = document.querySelectorAll('[data-sqe="item"]');
            addLog(`Tìm thấy ${items.length} thẻ sản phẩm trên trang này`, 'info');

            if (items.length === 0) {
                addLog('Không tìm thấy sản phẩm nào! Đảm bảo bạn đang ở trang tìm kiếm/danh mục.', 'error');
                return [];
            }

            const { priceMin, soldMin } = getFilters();
            const eligible = [];

            for (let i = 0; i < items.length; i++) {
                if (!isRunning) return [];
                const item = items[i];
                try {
                    const aTag = item.querySelector('a.contents, a');
                    if (!aTag) continue;

                    const parsed = parseProductUrl(aTag.getAttribute('href'));
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

                    if (crawledCache.includes(parsed.itemId) || skippedCache.includes(parsed.itemId)) {
                        addLog(`Bỏ qua: "${shortTitle}" (Trùng Cache)`, 'warn');
                        continue;
                    }

                    // Áp dụng Blacklist từ khóa nhạy cảm tại VN
                    if (SITE.code === 'vn' && title) {
                        const lowercaseTitle = title.toLowerCase();
                        const matchedKeyword = (blacklistConfig || []).find(keyword => {
                            const kw = keyword.trim().toLowerCase();
                            return kw && lowercaseTitle.includes(kw);
                        });
                        if (matchedKeyword) {
                            addLog(`Bỏ qua: "${shortTitle}" (Blacklist: "${matchedKeyword}")`, 'warn');
                            continue;
                        }
                    }

                    let priceVal = 0;
                    const promotionPriceNode = item.querySelector('[aria-label="promotion price"]');
                    if (promotionPriceNode && promotionPriceNode.parentElement) {
                        priceVal = cleanNumber(promotionPriceNode.parentElement.textContent);
                    } else {
                        const textNodes = Array.from(item.querySelectorAll('*'));
                        const priceNode = textNodes.find(node => node.textContent.includes(SITE.currency) && node.children.length === 0);
                        if (priceNode) priceVal = cleanNumber(priceNode.textContent);
                    }

                    if (priceVal < priceMin) {
                        addLog(`Bỏ qua: "${shortTitle}" (Giá ${money(priceVal)} < ${money(priceMin)})`, 'warn');
                        continue;
                    }

                    let soldVal = 0;
                    const textNodes = Array.from(item.querySelectorAll('*'));
                    const soldNode = textNodes.find(node => {
                        const text = node.textContent.trim().toLowerCase();
                        const hasKeyword = text.includes('sold') || text.includes('đã bán');
                        return hasKeyword && /\d/.test(text) && text.length < 30 && node.children.length === 0;
                    });
                    if (soldNode) soldVal = cleanNumber(soldNode.textContent);

                    if (soldVal < soldMin) {
                        addLog(`Bỏ qua: "${shortTitle}" (Bán ${soldVal} < ${soldMin})`, 'warn');
                        continue;
                    }

                    const productUrl = parsed.productUrl;
                    const itemData = {
                        [C.itemId]: parsed.itemId,
                        [C.shopId]: parsed.shopId,
                        [C.name]: title,
                        [C.price]: priceVal,
                        [C.sold]: soldVal,
                        [C.url]: productUrl,
                        [C.commission]: 0,
                        [C.sellerCom]: 0,
                        [C.shopeeCom]: 0,
                        [C.group]: `gr-${parsed.itemId}`,
                        [C.images]: '',
                        [C.category]: '',
                        [C.stock]: 0,
                        [C.rating]: 0,
                        [C.ratingCount]: 0,
                        [C.postDate]: '',
                        [C.linkType]: 'L1',
                        [C.mergedLink]: ''
                    };

                    const tabUrl = `https://${SITE.affiliateDomain}/offer/product_offer/${parsed.itemId}`;
                    eligible.push({ itemData, tabUrl, itemId: parsed.itemId });
                } catch (err) {
                    addLog(`Lỗi xử lý thẻ sản phẩm: ${err.message}`, 'error');
                }
            }
            return eligible;
        }

        function getAllVisibleItemIds() {
            const items = document.querySelectorAll('[data-sqe="item"]');
            const ids = [];
            for (let i = 0; i < items.length; i++) {
                const aTag = items[i].querySelector('a.contents, a');
                if (aTag) {
                    const parsed = parseProductUrl(aTag.getAttribute('href'));
                    if (parsed && parsed.itemId) {
                        ids.push(String(parsed.itemId));
                    }
                }
            }
            return ids;
        }

        function isPageIdentical(ids1, ids2) {
            if (ids1.length === 0 || ids2.length === 0) return false;
            let matchCount = 0;
            for (const id of ids1) {
                if (ids2.includes(id)) matchCount++;
            }
            const matchRatio = matchCount / ids1.length;
            return matchRatio > 0.85; // Trên 85% trùng là coi như trang không đổi
        }

        async function crawlCurrentPage() {
            try {
                if (!isRunning) return;

                let lastPageItemIds = [];
                let consecutiveSamePageCount = 0;
                let crawledPagesCount = 0;
                const MAX_PAGES_PER_FILTER = 20; // Giới hạn tối đa 20 trang mỗi bộ lọc

                while (isRunning) {
                    // Check captcha trước khi xử lý trang
                    await checkAndWaitForCaptcha();
                    if (!isRunning) return;

                    const eligible = await scrapeVisiblePage();
                    if (!isRunning) return;

                    // Lấy danh sách toàn bộ ID sản phẩm hiển thị trên trang hiện hành
                    const currentVisibleIds = getAllVisibleItemIds();

                    // Kiểm tra kẹt trang
                    if (lastPageItemIds.length > 0) {
                        const isIdentical = isPageIdentical(currentVisibleIds, lastPageItemIds);
                        if (isIdentical) {
                            consecutiveSamePageCount++;
                            addLog(`[Cảnh báo] Phát hiện nội dung trang trùng khớp với trang trước (Lần ${consecutiveSamePageCount}/2).`, 'warn');
                            if (consecutiveSamePageCount >= 2) {
                                addLog('Trình duyệt bị kẹt trang (Next Page nhưng nội dung không đổi). Dừng cào bộ lọc này để chống lag.', 'error');
                                break;
                            }
                        } else {
                            consecutiveSamePageCount = 0;
                        }
                    }
                    lastPageItemIds = currentVisibleIds;

                    if (eligible.length > 0) {
                        addLog(`Tìm thấy ${eligible.length} SP đạt bộ lọc sơ bộ trên trang. Bắt đầu xử lý song song theo nhóm 3...`, 'info');
                        const pageData = [];
                        const BATCH_SIZE = 3;

                        for (let i = 0; i < eligible.length; i += BATCH_SIZE) {
                            if (!isRunning) return;
                            const batch = eligible.slice(i, i + BATCH_SIZE);
                            addLog(`Đang cào nhóm sản phẩm ${Math.floor(i / BATCH_SIZE) + 1}/${Math.ceil(eligible.length / BATCH_SIZE)}...`, 'info');

                            const sellerComMin = getFilters().sellerComMin;
                            const promises = batch.map(prod => processProduct(prod.itemData, prod.tabUrl, sellerComMin, sheetConfig.selectedMode === 'keyword' ? keywordsList[currentKeywordIndex] : ''));
                            const results = await Promise.all(promises);

                            for (const itemArray of results) {
                                if (itemArray && itemArray.length > 0) {
                                    for (const finalItem of itemArray) {
                                        pageData.push(finalItem);
                                        crawledData.push(finalItem);
                                        const finalItemId = String(finalItem[C.itemId]);
                                        if (!crawledCache.includes(finalItemId)) {
                                            crawledCache.push(finalItemId);
                                        }
                                    }
                                }
                            }

                            GM_setValue(gmKey('crawled_cache'), crawledCache);
                            updateCacheCount();
                            const countEl = document.getElementById(elId('count'));
                            if (countEl) countEl.innerText = crawledData.length;

                            saveRunningState(sheetConfig.selectedMode === 'category' ? 'crawling_popular' : 'processing');
                            await sleep(3000 + Math.random() * 1500);
                        }

                        // Đồng bộ Sheets trang hiện tại ngay lập tức (Phương án B)
                        if (pageData.length > 0) {
                            updateMergedLinks(pageData);
                            addLog(`Đẩy ${pageData.length} SP đạt chuẩn lên Google Sheets...`, 'info');
                            try {
                                await pushToGoogleSheets(pageData);
                                addLog(`Đồng bộ dữ liệu trang lên Google Sheets thành công!`, 'success');
                            } catch (err) {
                                addLog(`Lỗi đồng bộ Sheets: ${err.message}`, 'error');
                                addLog('Tải file Excel dự phòng cho trang này...', 'warn');
                                exportExcelForData(pageData, `backup_page`);
                            }
                            if (videoaiConfig.autoPush) {
                                addLog(`Tự động đẩy ${pageData.length} SP lên VideoAI...`, 'info');
                                try {
                                    await pushToVideoAI(pageData);
                                } catch (err) {
                                    addLog(`Lỗi tự động đẩy VideoAI: ${err.message}`, 'error');
                                }
                            }
                        }
                    } else {
                        addLog('Không có sản phẩm nào đạt bộ lọc sơ bộ trên trang này.', 'warn');
                    }

                    crawledPagesCount++;
                    if (crawledPagesCount >= MAX_PAGES_PER_FILTER) {
                        addLog(`Đã đạt giới hạn tối đa ${MAX_PAGES_PER_FILTER} trang cào cho bộ lọc này.`, 'warn');
                        break;
                    }

                    // Chuyển trang tiếp theo
                    const nextPageBtn = document.querySelector('.shopee-icon-button--right:not([aria-disabled="true"]):not([disabled])');
                    if (nextPageBtn) {
                        addLog('Phát hiện nút chuyển trang. Chuyển trang sau 3 giây...', 'info');
                        setStatus('Next Page...', '#9c27b0');
                        nextPageBtn.click();
                        await sleep(3500);
                    } else {
                        break;
                    }
                }

                if (!isRunning) return;
                addLog('Đã cào xong tất cả các trang của bộ lọc hiện tại.', 'success');

                if (sheetConfig.selectedMode === 'category') {
                    if (!isTopSalesSelected()) {
                        await transitionToTopSalesAndCrawl();
                    } else {
                        await finishCrawlingSession();
                    }
                } else {
                    if (!isTopSalesSelected()) {
                        isTopSales = true;
                        await transitionToTopSalesAndCrawl();
                    } else {
                        await finishKeywordSession();
                    }
                }
            } catch (err) {
                console.error('Lỗi cào sản phẩm:', err);
                addLog(`Lỗi cào sản phẩm: ${err.message}. Dừng cào.`, 'error');
                stopScraper();
            }
        }

        async function checkKeywordPageAndCrawl() {
            if (!isRunning) return;
            updateKeywordProgressUI();
            if (currentKeywordIndex >= keywordsList.length) {
                await finishKeywordScrapingSession();
                return;
            }

            const keyword = keywordsList[currentKeywordIndex];
            const urlParams = new URLSearchParams(window.location.search);
            const urlKeyword = urlParams.get('keyword') || '';
            const urlSortBy = urlParams.get('sortBy') || '';

            const isCorrectKeyword = urlKeyword.toLowerCase() === keyword.toLowerCase();
            const isCorrectSort = isTopSales ? (urlSortBy === 'sales') : (urlSortBy !== 'sales');

            if (!isCorrectKeyword || !isCorrectSort) {
                const actionText = isTopSales ? 'bán chạy' : 'liên quan';
                addLog(`Chuyển hướng sang trang tìm kiếm từ khóa "${keyword}" (Chế độ: ${actionText})...`, 'info');
                setStatus('Redirecting...', '#2196f3');
                
                saveRunningState('keyword_redirecting');
                await sleep(1000);

                const targetUrl = `https://${HOST}/search?keyword=${encodeURIComponent(keyword)}${isTopSales ? '&sortBy=sales' : ''}`;
                window.location.href = targetUrl;
                return;
            }

            addLog(`Bắt đầu cào từ khóa "${keyword}" (${isTopSales ? 'Bán chạy' : 'Liên quan'})...`, 'info');
            await crawlCurrentPage();
        }

        async function transitionToTopSalesAndCrawl() {
            if (!isRunning) return;

            if (isTopSalesSelected()) {
                if (sheetConfig.selectedMode === 'category') {
                    await finishCrawlingSession();
                } else {
                    await finishKeywordSession();
                }
                return;
            }

            if (sheetConfig.selectedMode === 'keyword') {
                isTopSales = true;
                const keyword = keywordsList[currentKeywordIndex];
                addLog(`Chuyển hướng sang bộ lọc Bán chạy của từ khóa "${keyword}"...`, 'info');
                saveRunningState('waiting_topsales');
                await sleep(1000);
                window.location.href = `https://${HOST}/search?keyword=${encodeURIComponent(keyword)}&sortBy=sales`;
                return;
            }

            addLog(`Chuyển sang bộ lọc "${SITE.topSalesLabel}" để cào tiếp...`, 'info');
            const targetBtn = getElementByXPath(`//span[text()="${SITE.topSalesLabel}"]`) || getElementByXPath(`//div[text()="${SITE.topSalesLabel}"]`);
            if (targetBtn) {
                saveRunningState('waiting_topsales');
                targetBtn.click();
                addLog(`Đợi tải trang ${SITE.topSalesLabel}...`, 'info');
                setStatus('Loading Top Sales...', '#2196f3');

                let activated = false;
                for (let attempt = 1; attempt <= 10; attempt++) {
                    await sleep(1000);
                    if (!isRunning) return;
                    if (isTopSalesSelected()) { activated = true; break; }
                }

                if (activated) {
                    addLog(`Chuyển sang bộ lọc ${SITE.topSalesLabel} thành công! Bắt đầu cào...`, 'success');
                    await sleep(1500);
                    await crawlCurrentPage();
                } else {
                    addLog(`Không xác nhận được bộ lọc ${SITE.topSalesLabel} (Timeout). Kết thúc phiên cào.`, 'warn');
                    if (sheetConfig.selectedMode === 'category') {
                        await finishCrawlingSession();
                    } else {
                        await finishKeywordSession();
                    }
                }
            } else {
                addLog(`Không tìm thấy nút bộ lọc ${SITE.topSalesLabel}. Kết thúc phiên cào.`, 'warn');
                if (sheetConfig.selectedMode === 'category') {
                    await finishCrawlingSession();
                } else {
                    await finishKeywordSession();
                }
            }
        }

        async function finishCrawlingSession() {
            addLog('Hoàn thành cào danh mục/trang hiện tại.', 'success');

            // Chuyển sang danh mục con tiếp theo nếu cào toàn ngành
            if (isRunning) {
                const selectedCat = sheetConfig.selectedCategory;
                if (selectedCat && selectedCat !== 'no-category') {
                    const subCats = SITE.categories[selectedCat];
                    if (subCats && subCats.length > 0) {
                        const activeSub = getActiveSubCategoryText();
                        let nextIndex = 0;
                        if (activeSub) {
                            const idx = subCats.indexOf(activeSub);
                            if (idx !== -1) {
                                nextIndex = idx + 1;
                            } else {
                                addLog(`[Cảnh báo] Danh mục con đang mở "${activeSub}" không nằm trong nhóm ngành "${selectedCat}"!`, 'warn');
                                nextIndex = 0;
                            }
                        } else {
                            addLog(`Không có danh mục con nào đang mở. Tiến hành cào danh mục con đầu tiên của "${selectedCat}"...`, 'info');
                        }

                        if (nextIndex < subCats.length) {
                            const nextSubName = subCats[nextIndex];
                            addLog(`Chuyển hướng sang danh mục con tiếp theo: "${nextSubName}"...`, 'info');
                            setStatus(`Moving to: ${nextSubName}`, '#9c27b0');

                            crawledData = [];
                            saveRunningState('crawling_popular');

                            if (clickSubCategory(nextSubName)) {
                                addLog(`Click chuyển danh mục con "${nextSubName}" thành công. Đợi 4 giây...`, 'success');
                                await sleep(4000);
                                await crawlCurrentPage();
                                return;
                            } else {
                                addLog(`[Lỗi] Không thể click vào danh mục con "${nextSubName}"! Dừng kịch bản.`, 'error');
                                setStatus('Error!', '#d32f2f');
                                stopScraper();
                                return;
                            }
                        } else {
                            addLog(`Đã hoàn thành cào tất cả danh mục con của nhóm ngành "${selectedCat}"!`, 'success');
                        }
                    }
                }
            }

            setStatus('Done!', '#4caf50');
            stopScraper();
            showDoneOverlay();
        }

        async function finishKeywordSession() {
            const keyword = keywordsList[currentKeywordIndex];
            addLog(`Hoàn thành cào từ khóa: "${keyword}"`, 'success');
            
            currentKeywordIndex++;
            isTopSales = false; 
            updateKeywordProgressUI();
            
            if (currentKeywordIndex < keywordsList.length) {
                const nextKeyword = keywordsList[currentKeywordIndex];
                addLog(`Chuẩn bị chuyển sang từ khóa tiếp theo: "${nextKeyword}"...`, 'info');
                saveRunningState('keyword_redirecting');
                
                await sleep(2000);
                window.location.href = `https://${HOST}/search?keyword=${encodeURIComponent(nextKeyword)}`;
            } else {
                await finishKeywordScrapingSession();
            }
        }

        async function finishKeywordScrapingSession() {
            addLog('ĐÃ HOÀN THÀNH TOÀN BỘ DANH SÁCH TỪ KHÓA TÌM KIẾM!', 'success');
            setStatus('Done!', '#4caf50');
            stopScraper();
            showDoneOverlay();
        }

        async function startScraper() {
            if (isRunning) return;
            saveFilterConfig();

            isRunning = true;
            crawledData = [];
            GM_setValue(gmKey('global_captcha_detected'), false);

            syncButtonsRunning(true);

            if (sheetConfig.selectedMode === 'category') {
                saveRunningState('crawling_popular');
                addLog(`Bắt đầu cào danh mục...`, 'success');
                await crawlCurrentPage();
            } else if (sheetConfig.selectedMode === 'link') {
                // Cào theo link không điều hướng tab chính -> không dùng resume state
                GM_setValue(gmKey('running_state'), null);
                await runLinkScrapingLoop();
            } else {
                const kwInput = document.getElementById(elId('keywords-input'));
                const kwText = kwInput ? kwInput.value.trim() : '';
                keywordsList = kwText.split('\n').map(k => k.trim()).filter(k => k.length > 0);

                if (keywordsList.length === 0) {
                    alert('Vui lòng nhập ít nhất 1 từ khóa!');
                    isRunning = false;
                    syncButtonsRunning(false);
                    return;
                }

                currentKeywordIndex = 0;
                isTopSales = false;
                saveRunningState('keyword_redirecting');
                updateKeywordProgressUI();
                
                addLog(`Bắt đầu cào danh sách gồm ${keywordsList.length} từ khóa...`, 'success');
                addLog(`Chuyển hướng sang từ khóa đầu tiên: "${keywordsList[0]}"...`, 'info');
                await sleep(1000);

                window.location.href = `https://${HOST}/search?keyword=${encodeURIComponent(keywordsList[0])}`;
            }
        }

        function stopScraper() {
            if (!isRunning) return;
            isRunning = false;
            syncButtonsRunning(false);
            setStatus('Stopped', '#d32f2f');
            GM_setValue(gmKey('running_state'), null);
            updateKeywordProgressUI();
            addLog('Đã dừng cào.', 'warn');
        }

        // Khởi chạy Panel điều khiển
        if (document.readyState === 'complete' || document.readyState === 'interactive') {
            initPanel();
        } else {
            window.addEventListener('DOMContentLoaded', initPanel);
        }

        let reinitTimer = null;
        const observer = new MutationObserver(() => {
            if (reinitTimer) return;
            reinitTimer = setTimeout(() => {
                reinitTimer = null;
                if (!document.getElementById(PANEL_ID) && document.body) initPanel();
            }, 500);
        });
        observer.observe(document.documentElement, { childList: true, subtree: true });
    }

    // ============================================================
    // 3. KỊCH BẢN B: CÀO THEO TỪ KHÓA TÌM KIẾM (RUNS ON AFFILIATE.*)
    // ============================================================
    function runAffiliatePortalScraper() {
        let crawledData = [];
        let allCrawledData = [];
        let keywordsList = [];
        let currentKeywordIndex = 0;

        const panel = document.createElement('div');
        panel.id = PANEL_ID;
        panel.style.cssText = `position: fixed; top: 80px; right: 20px; z-index: 999999; background-color: #1a1a1a; color: #fff; padding: 15px; border-radius: 8px; box-shadow: 0 4px 15px rgba(0,0,0,0.5); font-family: Arial, sans-serif; font-size: 13px; width: 320px; border: 1px solid #ee4d2d;`;

        let googleConfig = GM_getValue(gmKey('google_config')) || { sheetUrl: '', tabName: 'Sheet1', clientId: '', clientSecret: '' };
        let videoaiConfig = GM_getValue(gmKey('videoai_config')) || { apiKey: '', endpoint: VIDEOAI_DEFAULT_ENDPOINT, autoPush: false };
        let filterConfig = GM_getValue(`${NS}_aff_filter_config_${SITE.code}`) || {
            commMin: SITE.defaults.commMin || 5000,
            soldMin: SITE.defaults.soldMin,
            sellerCommissionMin: SITE.defaults.sellerCommissionMin
        };

        let sheetConfig = {
            sheetUrl: googleConfig.sheetUrl || '',
            tabName: googleConfig.tabName || 'Sheet1',
            clientId: googleConfig.clientId || '',
            clientSecret: googleConfig.clientSecret || '',
            commMin: filterConfig.commMin,
            soldMin: filterConfig.soldMin,
            sellerCommissionMin: filterConfig.sellerCommissionMin
        };

        function saveRunningState(step) {
            GM_setValue(gmKey('running_state'), {
                isRunning: isRunning,
                keywords: keywordsList,
                currentKeywordIndex: currentKeywordIndex,
                crawledData: crawledData,
                allCrawledData: allCrawledData,
                currentStep: step
            });
        }

        function checkAndResumeState() {
            try {
                const savedState = GM_getValue(gmKey('running_state'));
                if (savedState && savedState.isRunning) {
                    addLog('Khôi phục phiên cào từ khóa trước đó bị gián đoạn...', 'warn');
                    isRunning = true;
                    keywordsList = savedState.keywords || [];
                    currentKeywordIndex = savedState.currentKeywordIndex || 0;
                    crawledData = savedState.crawledData || [];
                    allCrawledData = savedState.allCrawledData || [];

                    syncButtonsRunning(true);
                    setStatus('Resuming...', '#ff9800');

                    const kwTextarea = document.getElementById(elId('keywords-input'));
                    if (kwTextarea) kwTextarea.value = keywordsList.join('\n');

                    setTimeout(async () => {
                        addLog(`Tiếp tục từ khóa thứ ${currentKeywordIndex + 1}: "${keywordsList[currentKeywordIndex]}"...`, 'info');
                        await runKeywordScrapingLoop();
                    }, 3000);
                }
            } catch (e) {
                console.error('Lỗi khôi phục chạy:', e);
            }
        }

        function syncButtonsRunning(running) {
            const startBtn = document.getElementById(elId('btn-start'));
            const stopBtn = document.getElementById(elId('btn-stop'));
            if (startBtn) {
                startBtn.disabled = running;
                startBtn.style.backgroundColor = running ? '#555' : '#ee4d2d';
            }
            if (stopBtn) {
                stopBtn.disabled = !running;
                stopBtn.style.backgroundColor = running ? '#ee4d2d' : '#555';
            }
        }

        function initPanel() {
            if (document.getElementById(PANEL_ID)) return;

            panel.innerHTML = `
                <div style="font-weight: bold; font-size: 15px; margin-bottom: 12px; color: #ee4d2d; border-bottom: 1px solid #333; padding-bottom: 5px; display: flex; justify-content: space-between; align-items: center;">
                    <span>Affiliate Keyword Scraper</span>
                    <span id="${elId('status')}" style="font-size: 10px; padding: 2px 6px; background-color: #333; border-radius: 4px; color: #aaa;">Idle</span>
                </div>

                <div style="font-size: 10px; color: #888; margin-bottom: 10px;">Thị trường: <strong style="color:#ee4d2d;">${SITE.label}</strong> · Tiền tệ: ${SITE.currency}</div>

                <!-- Bộ lọc -->
                <div style="margin-bottom: 10px; display: flex; gap: 8px;">
                    <div style="flex: 1; display: flex; flex-direction: column;">
                        <label style="margin-bottom: 4px; font-weight: bold; font-size: 11px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; color: #ccc;" title="Hoa hồng tối thiểu trên trang danh sách (${SITE.currency})">Comm Min (${SITE.currency}):</label>
                        <input type="number" step="${SITE.defaults.comStep}" id="${elId('comm-min')}" value="${sheetConfig.commMin}" style="width: 100%; padding: 5px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white; font-size: 12px; box-sizing: border-box;">
                    </div>
                    <div style="flex: 1; display: flex; flex-direction: column;">
                        <label style="margin-bottom: 4px; font-weight: bold; font-size: 11px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; color: #ccc;" title="Đã bán tối thiểu">Đã bán Min:</label>
                        <input type="number" id="${elId('sold-min')}" value="${sheetConfig.soldMin}" style="width: 100%; padding: 5px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white; font-size: 12px; box-sizing: border-box;">
                    </div>
                    <div style="flex: 1; display: flex; flex-direction: column;">
                        <label style="margin-bottom: 4px; font-weight: bold; font-size: 11px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; color: #ccc;" title="Hoa hồng người bán tối thiểu khi cào chi tiết (${SITE.currency})">HH Real Min (${SITE.currency}):</label>
                        <input type="number" step="${SITE.defaults.comStep}" id="${elId('com-min')}" value="${sheetConfig.sellerCommissionMin}" style="width: 100%; padding: 5px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white; font-size: 12px; box-sizing: border-box;">
                    </div>
                </div>

                <!-- Từ khóa -->
                <div style="margin-bottom: 10px;">
                    <label style="display: block; margin-bottom: 4px; font-weight: bold; font-size: 11px; color: #ccc;">Từ khóa tìm kiếm (mỗi từ khóa 1 dòng):</label>
                    <textarea id="${elId('keywords-input')}" rows="4" placeholder="Nhập danh sách từ khóa..." style="width: 96%; padding: 6px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white; font-size: 11px; font-family: monospace; resize: vertical; box-sizing: border-box;"></textarea>
                </div>

                <!-- Cấu hình Google Sheets -->
                <div style="margin-bottom: 10px; border: 1px solid #333; border-radius: 4px;">
                    <div id="${elId('sheet-toggle')}" style="background-color: #2b2b2b; padding: 6px 10px; font-weight: bold; cursor: pointer; display: flex; justify-content: space-between; align-items: center; border-radius: 4px 4px 0 0;">
                        <span>⚙️ Cấu hình Google Sheets</span>
                        <span id="${elId('toggle-arrow')}">▲</span>
                    </div>
                    <div id="${elId('sheet-fields')}" style="padding: 10px; display: block; background-color: #1f1f1f;">
                        <div style="margin-bottom: 8px;">
                            <label style="display: block; margin-bottom: 3px; font-size: 11px;">Đường dẫn Spreadsheet:</label>
                            <input type="text" id="${elId('sheet-url')}" placeholder="https://docs.google.com/spreadsheets/d/..." value="${sheetConfig.sheetUrl}" style="width: 95%; padding: 5px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white; font-size: 11px;">
                        </div>
                        <div style="margin-bottom: 8px;">
                            <label style="display: block; margin-bottom: 3px; font-size: 11px;">Tên Sheet Tab:</label>
                            <input type="text" id="${elId('sheet-tab')}" placeholder="Sheet1" value="${sheetConfig.tabName}" style="width: 95%; padding: 5px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white; font-size: 11px;">
                        </div>
                        <div style="margin-bottom: 8px;">
                            <label style="display: block; margin-bottom: 3px; font-size: 11px;">Google Client ID:</label>
                            <input type="password" id="${elId('client-id')}" placeholder="Client ID" value="${sheetConfig.clientId}" style="width: 95%; padding: 5px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white; font-size: 11px;">
                        </div>
                        <div style="margin-bottom: 8px;">
                            <label style="display: block; margin-bottom: 3px; font-size: 11px;">Google Client Secret:</label>
                            <input type="password" id="${elId('client-secret')}" placeholder="Client Secret" value="${sheetConfig.clientSecret}" style="width: 95%; padding: 5px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white; font-size: 11px;">
                        </div>
                        <div style="display: flex; gap: 5px; justify-content: flex-end; margin-bottom: 8px; flex-wrap: wrap;">
                            <button id="${elId('btn-save-sheet')}" style="padding: 4px 6px; border: none; border-radius: 3px; background-color: #0288d1; color: white; cursor: pointer; font-size: 11px;">Lưu cấu hình</button>
                            <button id="${elId('btn-connect-google')}" style="padding: 4px 6px; border: none; border-radius: 3px; background-color: #e65100; color: white; cursor: pointer; font-size: 11px; font-weight: bold;">Kết nối Google</button>
                            <button id="${elId('btn-test-sheet')}" style="padding: 4px 6px; border: none; border-radius: 3px; background-color: #6a1b9a; color: white; cursor: pointer; font-size: 11px; font-weight: bold;">Test Sheet</button>
                        </div>

                        <div id="${elId('auth-code-wrapper')}" style="border-top: 1px dashed #444; padding-top: 8px; display: none;">
                            <label style="display: block; margin-bottom: 3px; font-size: 11px; color: #ffeb3b;">Nhập mã Authorization Code:</label>
                            <div style="display: flex; gap: 5px;">
                                <input type="text" id="${elId('auth-code-input')}" placeholder="Dán mã code tại đây" style="flex: 1; padding: 4px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white; font-size: 11px;">
                                <button id="${elId('btn-confirm-code')}" style="padding: 4px 8px; border: none; border-radius: 3px; background-color: #4caf50; color: white; cursor: pointer; font-size: 11px; font-weight: bold;">OK</button>
                            </div>
                        </div>
                    </div>
                </div>

                <!-- Cấu hình VideoAI API -->
                <div style="margin-bottom: 10px; border: 1px solid #333; border-radius: 4px;">
                    <div id="${elId('videoai-toggle')}" style="background-color: #2b2b2b; padding: 6px 10px; font-weight: bold; cursor: pointer; display: flex; justify-content: space-between; align-items: center; border-radius: 4px 4px 0 0;">
                        <span>🚀 Cấu hình VideoAI API</span>
                        <span id="${elId('videoai-toggle-arrow')}">▼</span>
                    </div>
                    <div id="${elId('videoai-fields')}" style="padding: 10px; display: none; background-color: #1f1f1f;">
                        <div style="margin-bottom: 8px;">
                            <label style="display: block; margin-bottom: 3px; font-size: 11px;">VideoAI API Key:</label>
                            <input type="password" id="${elId('videoai-key')}" placeholder="api_xxxxxxxx" value="${videoaiConfig.apiKey}" style="width: 95%; padding: 5px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white; font-size: 11px;">
                        </div>
                        <div style="margin-bottom: 8px;">
                            <label style="display: block; margin-bottom: 3px; font-size: 11px;">Endpoint:</label>
                            <input type="text" id="${elId('videoai-endpoint')}" placeholder="${VIDEOAI_DEFAULT_ENDPOINT}" value="${videoaiConfig.endpoint}" style="width: 95%; padding: 5px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white; font-size: 11px;">
                        </div>
                        <div style="margin-bottom: 8px; display: flex; align-items: center; gap: 6px;">
                            <input type="checkbox" id="${elId('videoai-autopush')}" ${videoaiConfig.autoPush ? 'checked' : ''} style="cursor: pointer;">
                            <label for="${elId('videoai-autopush')}" style="font-size: 11px; cursor: pointer;">Tự động đẩy lên VideoAI sau mỗi đợt cào</label>
                        </div>
                        <div style="display: flex; gap: 5px; justify-content: flex-end; flex-wrap: wrap;">
                            <button id="${elId('btn-save-videoai')}" style="padding: 4px 6px; border: none; border-radius: 3px; background-color: #0288d1; color: white; cursor: pointer; font-size: 11px;">Lưu cấu hình</button>
                            <button id="${elId('btn-test-videoai')}" style="padding: 4px 6px; border: none; border-radius: 3px; background-color: #6a1b9a; color: white; cursor: pointer; font-size: 11px; font-weight: bold;">Test kết nối</button>
                            <button id="${elId('btn-push-videoai')}" style="padding: 4px 6px; border: none; border-radius: 3px; background-color: #ee4d2d; color: white; cursor: pointer; font-size: 11px; font-weight: bold;">Đẩy lên VideoAI</button>
                        </div>
                    </div>
                </div>

                <!-- Cấu hình Blacklist từ khóa -->
                <div style="margin-bottom: 10px; border: 1px solid #333; border-radius: 4px;">
                    <div id="${elId('blacklist-toggle')}" style="background-color: #2b2b2b; padding: 6px 10px; font-weight: bold; cursor: pointer; display: flex; justify-content: space-between; align-items: center; border-radius: 4px 4px 0 0;">
                        <span>🚫 Cấu hình Blacklist (Chỉ VN)</span>
                        <span id="${elId('blacklist-toggle-arrow')}">▼</span>
                    </div>
                    <div id="${elId('blacklist-fields')}" style="padding: 10px; display: none; background-color: #1f1f1f;">
                        <div style="margin-bottom: 8px;">
                            <label style="display: block; margin-bottom: 4px; font-size: 11px; color: #ccc;">Từ khóa loại trừ (mỗi từ khóa 1 dòng):</label>
                            <textarea id="${elId('blacklist-input')}" rows="4" placeholder="áo ngực, bao cao su..." style="width: 95%; padding: 6px; border-radius: 4px; border: 1px solid #444; background-color: #2b2b2b; color: white; font-size: 11px; font-family: monospace; resize: vertical; box-sizing: border-box;">${blacklistConfig.join('\n')}</textarea>
                        </div>
                        <div style="display: flex; gap: 5px; justify-content: flex-end;">
                            <button id="${elId('btn-save-blacklist')}" style="padding: 4px 8px; border: none; border-radius: 3px; background-color: #0288d1; color: white; cursor: pointer; font-size: 11px; font-weight: bold;">Lưu Blacklist</button>
                        </div>
                    </div>
                </div>

                <div style="margin-bottom: 12px; display: flex; gap: 5px; justify-content: space-between; align-items: center; background-color: #2b2b2b; padding: 6px; border-radius: 4px;">
                    <span>Đã cào: <strong id="${elId('count')}" style="color: #ee4d2d;">0</strong> SP</span>
                    <span style="font-size: 11px; white-space: nowrap;">Cache: <strong id="${elId('cache-count')}" style="color: #4caf50;">0</strong> | Skip: <strong id="${elId('skip-count')}" style="color: #f44336;">0</strong></span>
                </div>

                <div style="display: flex; gap: 5px; margin-bottom: 8px;">
                    <button id="${elId('btn-start')}" style="flex: 1; padding: 8px; border: none; border-radius: 4px; background-color: #ee4d2d; color: white; font-weight: bold; cursor: pointer;">Start Crawl</button>
                    <button id="${elId('btn-stop')}" style="flex: 1; padding: 8px; border: none; border-radius: 4px; background-color: #555; color: white; font-weight: bold; cursor: pointer;" disabled>Stop</button>
                </div>
                <div style="display: flex; gap: 5px; margin-bottom: 8px;">
                    <button id="${elId('btn-clear-cache')}" style="width: 33%; padding: 6px; border: none; border-radius: 4px; background-color: #444; color: #ccc; cursor: pointer; font-size: 11px;">Xóa Cache</button>
                    <button id="${elId('btn-sync-cache')}" style="width: 34%; padding: 6px; border: none; border-radius: 4px; background-color: #0288d1; color: white; cursor: pointer; font-size: 11px; font-weight: bold;">Đồng bộ Cache</button>
                    <button id="${elId('btn-export')}" style="width: 33%; padding: 6px; border: none; border-radius: 4px; background-color: #2e7d32; color: white; cursor: pointer; font-size: 11px; font-weight: bold;">Xuất Excel</button>
                </div>
                <div style="margin-top: 10px;">
                    <div style="font-weight: bold; margin-bottom: 4px;">Logs chạy:</div>
                    <div id="${elId('logs')}" style="height: 110px; overflow-y: auto; background-color: #0c0c0c; border: 1px solid #333; padding: 6px; font-family: monospace; font-size: 11px; border-radius: 4px; word-break: break-all;">
                        <div style="color: #888;">Chờ lệnh Start...</div>
                    </div>
                </div>
            `;
            document.body.appendChild(panel);

            document.getElementById(elId('sheet-toggle')).addEventListener('click', () => {
                const fields = document.getElementById(elId('sheet-fields'));
                const arrow = document.getElementById(elId('toggle-arrow'));
                if (fields.style.display === 'none') {
                    fields.style.display = 'block';
                    arrow.innerText = '▲';
                } else {
                    fields.style.display = 'none';
                    arrow.innerText = '▼';
                }
            });

            document.getElementById(elId('blacklist-toggle')).addEventListener('click', () => {
                const fields = document.getElementById(elId('blacklist-fields'));
                const arrow = document.getElementById(elId('blacklist-toggle-arrow'));
                if (fields.style.display === 'none') {
                    fields.style.display = 'block';
                    arrow.innerText = '▲';
                } else {
                    fields.style.display = 'none';
                    arrow.innerText = '▼';
                }
            });

            document.getElementById(elId('btn-save-blacklist')).addEventListener('click', () => {
                const inputVal = document.getElementById(elId('blacklist-input')).value;
                const list = inputVal.split('\n').map(item => item.trim()).filter(item => item.length > 0);
                blacklistConfig = list;
                GM_setValue(gmKey('blacklist_config'), blacklistConfig);
                addLog('Đã lưu danh sách Blacklist!', 'success');
                alert('Đã lưu Blacklist thành công!');
            });

            document.getElementById(elId('btn-connect-google')).addEventListener('click', () => {
                readSheetInputsToConfig();
                GM_setValue(gmKey('google_config'), googleConfig);
                startGoogleAuthFlow(googleConfig.clientId);
                document.getElementById(elId('auth-code-wrapper')).style.display = 'block';
            });

            document.getElementById(elId('btn-confirm-code')).addEventListener('click', () => {
                const code = document.getElementById(elId('auth-code-input')).value.trim();
                if (!code) {
                    alert('Hãy dán mã Authorization Code!');
                    return;
                }
                exchangeAuthCodeForTokens(code, googleConfig.clientId, googleConfig.clientSecret);
                document.getElementById(elId('auth-code-input')).value = '';
                document.getElementById(elId('auth-code-wrapper')).style.display = 'none';
            });

            document.getElementById(elId('btn-test-sheet')).addEventListener('click', async () => {
                readSheetInputsToConfig();
                GM_setValue(gmKey('google_config'), googleConfig);
                await testGoogleSheetsConnection();
            });

            document.getElementById(elId('btn-save-sheet')).addEventListener('click', () => {
                readSheetInputsToConfig();
                sheetConfig.commMin = parseFloat(document.getElementById(elId('comm-min')).value) || 0;
                sheetConfig.soldMin = parseInt(document.getElementById(elId('sold-min')).value, 10) || 0;
                sheetConfig.sellerCommissionMin = parseFloat(document.getElementById(elId('com-min')).value) || 0;

                GM_setValue(gmKey('google_config'), googleConfig);
                GM_setValue(`${NS}_aff_filter_config_${SITE.code}`, {
                    commMin: sheetConfig.commMin,
                    soldMin: sheetConfig.soldMin,
                    sellerCommissionMin: sheetConfig.sellerCommissionMin
                });

                addLog('Đã lưu cấu hình và bộ lọc!', 'success');
                alert('Lưu cấu hình thành công!');
            });

            document.getElementById(elId('btn-start')).addEventListener('click', startScraper);
            document.getElementById(elId('btn-stop')).addEventListener('click', stopScraper);
            document.getElementById(elId('btn-sync-cache')).addEventListener('click', syncCacheFromGoogleSheets);

            document.getElementById(elId('btn-clear-cache')).addEventListener('click', () => {
                if (confirm(`Bạn có chắc chắn muốn xóa bộ nhớ đệm cache của ${SITE.label} không?`)) {
                    crawledCache = [];
                    skippedCache = [];
                    GM_setValue(gmKey('crawled_cache'), crawledCache);
                    GM_setValue(gmKey('skipped_cache'), skippedCache);
                    updateCacheCount();
                    addLog('Đã xóa sạch bộ nhớ đệm cache.', 'warn');
                    alert('Xóa cache thành công!');
                }
            });

            document.getElementById(elId('btn-export')).addEventListener('click', () => {
                updateMergedLinks(allCrawledData);
                exportExcelForData(allCrawledData, `all_keywords`);
            });

            document.getElementById(elId('videoai-toggle')).addEventListener('click', () => {
                const fields = document.getElementById(elId('videoai-fields'));
                const arrow = document.getElementById(elId('videoai-toggle-arrow'));
                if (fields.style.display === 'none') {
                    fields.style.display = 'block';
                    arrow.innerText = '▲';
                } else {
                    fields.style.display = 'none';
                    arrow.innerText = '▼';
                }
            });

            document.getElementById(elId('btn-save-videoai')).addEventListener('click', () => {
                readVideoAIInputsToConfig();
                GM_setValue(gmKey('videoai_config'), videoaiConfig);
                addLog('Đã lưu cấu hình VideoAI API!', 'success');
                alert('Lưu cấu hình VideoAI thành công!');
            });

            document.getElementById(elId('btn-test-videoai')).addEventListener('click', async () => {
                readVideoAIInputsToConfig();
                GM_setValue(gmKey('videoai_config'), videoaiConfig);
                await testVideoAIConnection();
            });

            document.getElementById(elId('btn-push-videoai')).addEventListener('click', async () => {
                readVideoAIInputsToConfig();
                GM_setValue(gmKey('videoai_config'), videoaiConfig);
                if (!allCrawledData || allCrawledData.length === 0) {
                    alert('Chưa có sản phẩm nào để đẩy!');
                    return;
                }
                updateMergedLinks(allCrawledData);
                addLog(`Đẩy thủ công ${allCrawledData.length} SP lên VideoAI...`, 'info');
                try {
                    await pushToVideoAI(allCrawledData);
                } catch (err) {
                    addLog(`Lỗi đẩy VideoAI: ${err.message}`, 'error');
                    alert(`Lỗi đẩy VideoAI: ${err.message}`);
                }
            });

            updateCacheCount();
            checkAndResumeState();
        }

        function readSheetInputsToConfig() {
            googleConfig.sheetUrl = document.getElementById(elId('sheet-url')).value.trim();
            googleConfig.tabName = document.getElementById(elId('sheet-tab')).value.trim();
            googleConfig.clientId = document.getElementById(elId('client-id')).value.trim();
            googleConfig.clientSecret = document.getElementById(elId('client-secret')).value.trim();
        }

        function readVideoAIInputsToConfig() {
            videoaiConfig.apiKey = document.getElementById(elId('videoai-key')).value.trim();
            videoaiConfig.endpoint = document.getElementById(elId('videoai-endpoint')).value.trim() || VIDEOAI_DEFAULT_ENDPOINT;
            videoaiConfig.autoPush = document.getElementById(elId('videoai-autopush')).checked;
        }

        // ============================================================
        // VÒNG LẶP CÀO SẢN PHẨM THEO TỪ KHÓA
        // ============================================================
        function getElementByXPath(xpath) {
            return document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
        }

        function findElementByText(textList) {
            for (const text of textList) {
                const el = getElementByXPath(`//*[text()="${text}"]`);
                if (el) return el;
            }
            return null;
        }

        function getProductItems() {
            return Array.from(document.querySelectorAll('.product-offer-item'));
        }

        function setReactInput(inputEl, value) {
            if (!inputEl) return;
            const nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, "value").set;
            nativeInputValueSetter.call(inputEl, value);
            
            const ev = new Event('input', { bubbles: true });
            inputEl.dispatchEvent(ev);
            
            const evChange = new Event('change', { bubbles: true });
            inputEl.dispatchEvent(evChange);
        }

        async function performSearch(keyword) {
            addLog(`Đang tìm kiếm từ khóa: "${keyword}"...`, 'info');
            const searchInput = document.querySelector('input.ant-input-lg') || getElementByXPath('//*[@class="ant-input ant-input-lg"]');
            if (!searchInput) {
                addLog('Không tìm thấy ô nhập từ khóa trên giao diện!', 'error');
                return false;
            }

            searchInput.focus();
            setReactInput(searchInput, '');
            await sleep(300);
            setReactInput(searchInput, keyword);
            await sleep(500);

            searchInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', keyCode: 13, bubbles: true }));
            searchInput.dispatchEvent(new KeyboardEvent('keypress', { key: 'Enter', code: 'Enter', keyCode: 13, bubbles: true }));
            searchInput.dispatchEvent(new KeyboardEvent('keyup', { key: 'Enter', code: 'Enter', keyCode: 13, bubbles: true }));

            const searchBtn = document.querySelector('.ant-input-search-button') || getElementByXPath('//button[contains(@class, "ant-input-search-button")]');
            if (searchBtn) {
                searchBtn.click();
            }

            await sleep(4000);
            return true;
        }

        async function applyFiltersOnPage() {
            addLog('Đang kiểm tra và áp dụng bộ lọc Top Sales & Comm Xtra...', 'info');

            const topSalesBtn = findElementByText(['Top Sales', 'Bán chạy']);
            if (topSalesBtn) {
                addLog('Click chọn bộ lọc Top Sales...', 'info');
                topSalesBtn.click();
                await sleep(2500);
            } else {
                addLog('Không tìm thấy bộ lọc Top Sales.', 'warn');
            }

            const isCommXtraChecked = !!getElementByXPath('//*[@class="ant-radio-button-wrapper ant-radio-button-wrapper-checked"]//*[text()="Comm Xtra"]') ||
                                      !!getElementByXPath('//*[@class="ant-radio-button-wrapper ant-radio-button-wrapper-checked"]//*[text()="Hoa hồng Xtra"]');
            if (!isCommXtraChecked) {
                const commXtraBtn = SITE.code === 'vn'
                    ? getElementByXPath('//*[@class="ant-radio-button-wrapper"]//*[text()="Hoa hồng Xtra"]')
                    : findElementByText(['Comm Xtra', 'Hoa hồng Xtra']);
                if (commXtraBtn) {
                    addLog('Click chọn bộ lọc Comm Xtra...', 'info');
                    commXtraBtn.click();
                    await sleep(3000);
                } else {
                    addLog('Không tìm thấy nút lọc Comm Xtra.', 'warn');
                }
            } else {
                addLog('Bộ lọc Comm Xtra đã được chọn sẵn.', 'info');
            }

            return true;
        }

        async function scrapeAllPagesOfKeyword() {
            let pageCount = 1;
            while (isRunning) {
                addLog(`Đang cào trang ${pageCount} cho từ khóa: "${keywordsList[currentKeywordIndex]}"...`, 'info');
                setStatus(`Page ${pageCount} - ${keywordsList[currentKeywordIndex]}`, '#009688');

                await smoothScrollToBottom();
                await sleep(1500);

                const items = getProductItems();
                addLog(`Tìm thấy ${items.length} sản phẩm trên trang kết quả này.`, 'info');

                if (items.length === 0) {
                    addLog('Không tìm thấy sản phẩm nào.', 'warn');
                    break;
                }

                const { soldMin, commMin, sellerComMin } = getFilters();
                const eligibleProducts = [];

                for (let i = 0; i < items.length; i++) {
                    const item = items[i];
                    try {
                        let soldText = '';
                        const textElements = Array.from(item.querySelectorAll('*'));
                        const soldNode = textElements.find(node => {
                            const t = node.textContent.toLowerCase();
                            const isSoldLabel = SITE.code === 'vn'
                                ? (t.includes('sold') || t.includes('đã bán') || t.includes('lượt bán') || t.includes('bán/tháng'))
                                : (t.includes('sold') || t.includes('đã bán'));
                            return isSoldLabel && /\d/.test(t) && node.children.length === 0;
                        });
                        if (soldNode) soldText = soldNode.textContent;
                        const soldVal = cleanNumber(soldText);

                        let title = '';
                        const titleEl = item.querySelector('.name, .product-name, div[class*="name"]');
                        if (titleEl) {
                            title = titleEl.textContent.trim();
                        } else {
                            const h3 = item.querySelector('h3');
                            if (h3) title = h3.textContent.trim();
                        }
                        if (!title) {
                            const possibleTitle = textElements.find(node => node.textContent.length > 8 && node.children.length === 0);
                            if (possibleTitle) title = possibleTitle.textContent.trim();
                        }
                        const shortTitle = title.length > 20 ? title.substring(0, 20) + '...' : title;

                        if (SITE.code === 'vn' && title) {
                            const lowercaseTitle = title.toLowerCase();
                            const matchedKeyword = (blacklistConfig || []).find(keyword => {
                                const kw = keyword.trim().toLowerCase();
                                return kw && lowercaseTitle.includes(kw);
                            });
                            if (matchedKeyword) {
                                addLog(`Bỏ qua: "${shortTitle}" (Chứa từ khóa nhạy cảm: "${matchedKeyword}")`, 'warn');
                                continue;
                            }
                        }

                        const priceEl = item.querySelector('.price, div[class*="price"]');
                        const priceVal = priceEl ? cleanNumber(priceEl.textContent) : 0;

                        const commRateEl = item.querySelector('.commRate, div[class*="commRate"]');
                        let commRateVal = 0;
                        if (commRateEl) {
                            const m = commRateEl.textContent.match(/[\d.,]+/);
                            if (m) {
                                commRateVal = parseFloat(m[0].replace(/,/g, '.'));
                            }
                        }

                        const estimatedComm = priceVal * (commRateVal / 100);

                        const aTag = item.querySelector('a');
                        if (!aTag) continue;
                        const href = aTag.getAttribute('href');
                        let itemId = '';
                        let shopId = '';
                        if (href) {
                            const detailMatch = href.match(/\/offer\/product_offer\/(\d+)/);
                            if (detailMatch) {
                                itemId = detailMatch[1];
                            } else {
                                const shopItemMatch = href.match(/i\.(\d+)\.(\d+)/);
                                if (shopItemMatch) {
                                    shopId = shopItemMatch[1];
                                    itemId = shopItemMatch[2];
                                }
                            }
                        }

                        if (!itemId) {
                            addLog(`Bỏ qua: Không thấy Item ID sản phẩm "${shortTitle}".`, 'warn');
                            continue;
                        }

                        if (crawledCache.includes(String(itemId)) || skippedCache.includes(String(itemId))) {
                            addLog(`Bỏ qua: "${shortTitle}" (Trùng Cache)`, 'warn');
                            continue;
                        }

                        if (soldVal < soldMin) {
                            addLog(`Bỏ qua: "${shortTitle}" (Lượt bán ${soldVal} < ${soldMin})`, 'warn');
                            continue;
                        }

                        if (estimatedComm < commMin) {
                            addLog(`Bỏ qua: "${shortTitle}" (HH ước tính ${money(estimatedComm.toFixed(2))} < ${money(commMin)})`, 'warn');
                            continue;
                        }

                        const productUrl = `https://${SITE.domain}/product/${shopId || 'shop'}/${itemId}`;
                        const itemData = {
                            [C.itemId]: itemId,
                            [C.shopId]: shopId || 'N/A',
                            [C.name]: title,
                            [C.price]: priceVal,
                            [C.sold]: soldVal,
                            [C.url]: productUrl,
                            [C.commission]: 0,
                            [C.sellerCom]: 0,
                            [C.shopeeCom]: 0,
                            [C.group]: `gr-${itemId}`,
                            [C.images]: '',
                            [C.category]: '',
                            [C.stock]: 0,
                            [C.rating]: 0,
                            [C.ratingCount]: 0,
                            [C.postDate]: '',
                            [C.linkType]: 'L1',
                            [C.mergedLink]: ''
                        };

                        const tabUrl = window.location.origin + href;
                        eligibleProducts.push({ itemData, tabUrl, itemId });
                    } catch (e) {
                        addLog(`Lỗi lọc sơ bộ từ khóa: ${e.message}`, 'error');
                    }
                }

                const BATCH_SIZE = 3;
                if (eligibleProducts.length > 0) {
                    addLog(`Tìm thấy ${eligibleProducts.length} SP đạt lọc sơ bộ. Mở tab chi tiết song song (batch = 3)...`, 'info');
                    
                    for (let i = 0; i < eligibleProducts.length; i += BATCH_SIZE) {
                        if (!isRunning) return;
                        const batch = eligibleProducts.slice(i, i + BATCH_SIZE);
                        addLog(`Đang cào nhóm sản phẩm ${Math.floor(i / BATCH_SIZE) + 1}/${Math.ceil(eligibleProducts.length / BATCH_SIZE)}...`, 'info');
                        
                        const promises = batch.map(prod => processProduct(prod.itemData, prod.tabUrl, sellerComMin, keywordsList[currentKeywordIndex]));
                        const results = await Promise.all(promises);

                        for (const itemArray of results) {
                            if (itemArray && itemArray.length > 0) {
                                for (const finalItem of itemArray) {
                                    crawledData.push(finalItem);
                                    allCrawledData.push(finalItem);
                                    const finalItemId = String(finalItem[C.itemId]);
                                    if (!crawledCache.includes(finalItemId)) {
                                        crawledCache.push(finalItemId);
                                    }
                                }

                                GM_setValue(gmKey('crawled_cache'), crawledCache);
                                updateCacheCount();

                                const countEl = document.getElementById(elId('count'));
                                if (countEl) countEl.innerText = crawledData.length;

                                saveRunningState('processing');
                            }
                        }

                        await sleep(3000 + Math.random() * 1500);
                    }
                } else {
                    addLog('Không có sản phẩm nào đạt chuẩn sơ bộ trên trang này.', 'warn');
                }

                const nextBtn = document.querySelector('.ant-pagination-next:not(.ant-pagination-disabled) button, .ant-pagination-next:not(.ant-pagination-disabled), .fa.page-item.page-next:not(.disabled)') || 
                                getElementByXPath('//*[@class="fa page-item page-next"]');
                if (nextBtn) {
                    addLog('Chuyển sang trang kết quả tiếp theo sau 3 giây...', 'info');
                    setStatus('Next Page...', '#9c27b0');
                    nextBtn.click();
                    pageCount++;
                    await sleep(3500);
                } else {
                    addLog('Đã hết trang kết quả tìm kiếm cho từ khóa.', 'success');
                    break;
                }
            }
        }

        async function runKeywordScrapingLoop() {
            while (isRunning && currentKeywordIndex < keywordsList.length) {
                const keyword = keywordsList[currentKeywordIndex];
                addLog(`========================================`, 'info');
                addLog(`Bắt đầu cào từ khóa (${currentKeywordIndex + 1}/${keywordsList.length}): "${keyword}"`, 'info');
                setStatus(`Keyword: ${keyword}`, '#009688');

                crawledData = [];
                saveRunningState('processing');

                const searchSuccess = await performSearch(keyword);
                if (!searchSuccess) {
                    addLog(`Không tìm kiếm được từ khóa "${keyword}". Bỏ qua.`, 'error');
                    currentKeywordIndex++;
                    saveRunningState('next_keyword');
                    continue;
                }

                await applyFiltersOnPage();
                await scrapeAllPagesOfKeyword();

                if (!isRunning) return;

                if (crawledData.length > 0) {
                    updateMergedLinks(crawledData);
                    addLog(`Đẩy ${crawledData.length} SP cào được của từ khóa "${keyword}" lên Google Sheets...`, 'info');
                    try {
                        await pushToGoogleSheets(crawledData);
                        addLog(`Đồng bộ dữ liệu từ khóa "${keyword}" thành công!`, 'success');
                    } catch (err) {
                        addLog(`Lỗi đồng bộ Sheets từ khóa "${keyword}": ${err.message}`, 'error');
                        addLog('Xuất file Excel dự phòng cho từ khóa này...', 'warn');
                        exportExcelForData(crawledData, `backup_${keyword}`);
                    }
                    if (videoaiConfig.autoPush) {
                        addLog(`Tự động đẩy ${crawledData.length} SP từ khóa "${keyword}" lên VideoAI...`, 'info');
                        try {
                            await pushToVideoAI(crawledData);
                        } catch (err) {
                            addLog(`Lỗi tự động đẩy VideoAI: ${err.message}`, 'error');
                        }
                    }
                } else {
                    addLog(`Không cào được sản phẩm nào đạt chuẩn cho từ khóa "${keyword}".`, 'warn');
                }

                currentKeywordIndex++;
                saveRunningState('next_keyword');
                await sleep(3000);
            }

            if (isRunning) {
                addLog('========================================', 'success');
                addLog('ĐÃ HOÀN THÀNH TOÀN BỘ DANH SÁCH TỪ KHÓA TÌM KIẾM!', 'success');
                setStatus('Finished!', '#4caf50');
                await stopScraper();
                showDoneOverlay();
            }
        }

        async function startScraper() {
            if (isRunning) return;

            const kwInput = document.getElementById(elId('keywords-input'));
            if (!kwInput) return;
            keywordsList = kwInput.value.split('\n').map(k => k.trim()).filter(k => k.length > 0);

            if (keywordsList.length === 0) {
                alert('Vui lòng nhập ít nhất 1 từ khóa!');
                return;
            }

            isRunning = true;
            currentKeywordIndex = 0;
            crawledData = [];
            allCrawledData = [];
            GM_setValue(gmKey('global_captcha_detected'), false);

            syncButtonsRunning(true);
            saveRunningState('starting');

            if (!window.location.pathname.startsWith('/offer/product_offer')) {
                addLog('Tự động chuyển hướng sang trang tìm kiếm sản phẩm của affiliate...', 'warn');
                await sleep(1000);
                window.location.href = `https://${window.location.hostname}/offer/product_offer`;
                return;
            }

            addLog(`Bắt đầu xử lý danh sách gồm ${keywordsList.length} từ khóa...`, 'success');
            await runKeywordScrapingLoop();
        }

        async function stopScraper() {
            if (!isRunning) return;
            isRunning = false;

            await sleep(500);

            if (crawledData && crawledData.length > 0) {
                updateMergedLinks(crawledData);
                setStatus('Stopping & Syncing...', '#e65100');
                const stopBtn = document.getElementById(elId('btn-stop'));
                if (stopBtn) stopBtn.disabled = true;

                addLog(`Đang dừng kịch bản... Đồng bộ nốt ${crawledData.length} SP đã cào của từ khóa hiện tại lên Google Sheets...`, 'warn');

                try {
                    await pushToGoogleSheets(crawledData);
                    addLog(`Đồng bộ dữ liệu trước khi dừng lên Google Sheets thành công!`, 'success');
                } catch (err) {
                    addLog(`Lỗi đồng bộ Sheets khi dừng: ${err.message}`, 'error');
                    addLog('Tải file Excel dự phòng trước khi dừng...', 'warn');
                    const currentKeyword = keywordsList[currentKeywordIndex] || 'stopped';
                    exportExcelForData(crawledData, `backup_stop_${currentKeyword}`);
                }
                if (videoaiConfig.autoPush) {
                    try {
                        await pushToVideoAI(crawledData);
                    } catch (err) {
                        addLog(`Lỗi tự động đẩy VideoAI khi dừng: ${err.message}`, 'error');
                    }
                }
                crawledData = [];
            }

            syncButtonsRunning(false);
            setStatus('Stopped', '#d32f2f');
            GM_setValue(gmKey('running_state'), null);
            addLog('Đã dừng cào sản phẩm theo từ khóa.', 'warn');
        }

        // Khởi chạy Panel điều khiển
        if (document.readyState === 'complete' || document.readyState === 'interactive') {
            initPanel();
        } else {
            window.addEventListener('DOMContentLoaded', initPanel);
        }

        let reinitTimer = null;
        const observer = new MutationObserver(() => {
            if (reinitTimer) return;
            reinitTimer = setTimeout(() => {
                reinitTimer = null;
                if (!document.getElementById(PANEL_ID) && document.body) initPanel();
            }, 500);
        });
        observer.observe(document.documentElement, { childList: true, subtree: true });
    }
})();
