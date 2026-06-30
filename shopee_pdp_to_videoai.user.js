// ==UserScript==
// @name         Shopee PDP → VideoAI Pusher
// @namespace    http://tampermonkey.net/shopee-pdp-videoai
// @version      1.0
// @description  Nhận link sản phẩm Shopee, mở tab ngầm bắt API get_pc (trang sản phẩm thường) rồi đẩy dữ liệu lên VideoAI. Độc lập với tool affiliate.
// @author       Antigravity
// @match        https://shopee.vn/*
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_deleteValue
// @grant        GM_xmlhttpRequest
// @grant        GM_openInTab
// @grant        GM_registerMenuCommand
// @connect      videoai-api-dev.devappnow.com
// @connect      api.telegram.org
// @run-at       document-start
// ==/UserScript==

(function () {
    'use strict';

    // ============================================================
    // HẰNG SỐ & TIỆN ÍCH
    // ============================================================
    const NS = 'pdpvai';
    const PANEL_ID = `${NS}-panel`;
    const IMG_DOMAIN = 'down-vn';
    const MIN_IMAGES = 3;
    const MAX_BATCH = 200;
    const FLUSH_THRESHOLD = 30; // gom đủ ~30 SP thì đẩy 1 đợt lên VideoAI
    const CAPTURE_TIMEOUT_MS = 15000;
    const TG_COOLDOWN_MS = 60000;            // chống spam Telegram khi nhiều tab cùng dính captcha
    const SESSION_MAX_AGE_MS = 6 * 3600 * 1000; // bỏ qua session resume quá cũ
    const RESUME_DELAY_MS = 3000;            // chờ trước khi tự khôi phục phiên
    const CAPTCHA_DOM_DELAY_MS = 2500;       // chờ sau khi tải trang rồi mới kiểm tra captcha kiểu DOM
    const PDP_API = '/api/v4/pdp/get_pc';
    const VIDEOAI_DEFAULT_ENDPOINT = 'https://videoai-api-dev.devappnow.com/api/products/shopee-cache/batch';

    const sleep = (ms) => new Promise(r => setTimeout(r, ms));

    function imgUrl(hash) {
        if (!hash) return null;
        if (String(hash).startsWith('http')) return hash;
        return `https://${IMG_DOMAIN}.img.susercontent.com/file/${hash}`;
    }

    // Parse link -> { shopId, itemId, productUrl }. Hỗ trợ /product/{shop}/{item} và i.{shop}.{item}
    function parseLink(href) {
        if (!href) return null;
        let s = href.trim();
        try { s = decodeURIComponent(s); } catch (e) { }
        let m = s.match(/\/product\/(\d+)\/(\d+)/);
        if (!m) m = s.match(/i\.(\d+)\.(\d+)/);
        if (!m) return null;
        return {
            shopId: String(m[1]),
            itemId: String(m[2]),
            productUrl: `https://shopee.vn/product/${m[1]}/${m[2]}`
        };
    }

    // Loại link trùng theo itemId (link hợp lệ) hoặc theo nội dung dòng (link lạ). Giữ thứ tự, bỏ dòng trống.
    function dedupeLinksText(text) {
        const seen = new Set();
        const out = [];
        let removed = 0;
        (text || '').split('\n').forEach(line => {
            const t = line.trim();
            if (!t) return;
            const p = parseLink(t);
            const key = p ? `id:${p.itemId}` : `raw:${t.toLowerCase()}`;
            if (seen.has(key)) { removed++; return; }
            seen.add(key);
            out.push(t);
        });
        return { text: out.join('\n'), removed };
    }

    function gmRequest(opts) {
        return new Promise((resolve, reject) => {
            GM_xmlhttpRequest({
                ...opts,
                onload: resolve,
                onerror: reject,
                ontimeout: () => reject(new Error('timeout'))
            });
        });
    }

    function num(v) {
        const n = Number(v);
        return isNaN(n) ? 0 : n;
    }

    // ============================================================
    // TRẠNG THÁI (STATE) — gom về 1 chỗ, khởi tạo sớm
    // ============================================================
    let running = false;
    let resumeChecked = false;
    // Cache itemID đã xử lý xong (đẩy thành công HOẶC bị loại) -> không làm lại giữa các phiên
    let doneCache = (GM_getValue(`${NS}_done`, []) || []).map(String);
    // Tập itemId đang chờ worker bắt (đồng bộ xuống GM để worker nhiều tab cùng nhận diện)
    const expectingSet = new Set();
    function syncExpecting() {
        GM_setValue(`${NS}_expecting`, Array.from(expectingSet));
    }

    // ============================================================
    // ROUTING: trang sản phẩm = worker (hook), còn lại = panel điều khiển
    // ============================================================
    const isProductPage = /\/product\/\d+\/\d+/.test(window.location.pathname) ||
                          /i\.\d+\.\d+/.test(window.location.pathname);

    // Trang captcha/verify: chỉ báo Telegram + banner, KHÔNG dựng panel/worker (tránh resume cascade trên tab phụ dính captcha)
    if (isCaptchaUrl()) {
        handleCaptchaPage();
    } else {
        if (isProductPage) {
            runWorkerHook();
        } else {
            runControlPanel();
        }
        // Bắt captcha kiểu DOM (chèn sau khi tải trang) trên các trang bình thường
        window.addEventListener('load', () => {
            setTimeout(() => { if (isCaptchaDom()) handleCaptchaPage(); }, CAPTCHA_DOM_DELAY_MS);
        });
    }

    // ============================================================
    // PHÁT HIỆN CAPTCHA & THÔNG BÁO TELEGRAM
    // ============================================================
    function isCaptchaUrl() {
        const h = window.location.href;
        return h.includes('/verify/captcha') || h.includes('/verify/traffic') || h.includes('/verify/security');
    }

    function isCaptchaDom() {
        return !!document.querySelector('.check-captcha-box') ||
               !!document.querySelector('.shopee-captcha-wrapper') ||
               !!document.getElementById('captcha-submit') ||
               !!document.querySelector('iframe[src*="captcha"]');
    }

    function sendTelegram(text) {
        const token = (GM_getValue(`${NS}_tg_token`, '') || '').trim();
        const chat = (GM_getValue(`${NS}_tg_chat`, '') || '').trim();
        if (!token || !chat) return Promise.resolve({ ok: false, reason: 'Chưa cấu hình token/chat id' });
        return new Promise(resolve => {
            GM_xmlhttpRequest({
                method: 'POST',
                url: `https://api.telegram.org/bot${token}/sendMessage`,
                headers: { 'Content-Type': 'application/json' },
                data: JSON.stringify({ chat_id: chat, text, parse_mode: 'HTML', disable_web_page_preview: true }),
                onload: r => resolve({ ok: r.status === 200, status: r.status, body: r.responseText }),
                onerror: () => resolve({ ok: false, reason: 'network error' }),
                ontimeout: () => resolve({ ok: false, reason: 'timeout' })
            });
        });
    }

    function notifyCaptcha() {
        const token = (GM_getValue(`${NS}_tg_token`, '') || '').trim();
        const chat = (GM_getValue(`${NS}_tg_chat`, '') || '').trim();
        if (!token || !chat) return; // chưa cấu hình Telegram -> không báo, không tiêu cooldown
        const last = GM_getValue(`${NS}_tg_last`, 0);
        if (Date.now() - last < TG_COOLDOWN_MS) return; // cooldown tránh spam khi nhiều tab cùng dính
        GM_setValue(`${NS}_tg_last`, Date.now());
        const name = (GM_getValue(`${NS}_tg_name`, '') || 'PDP→VideoAI').trim();
        const msg = `⚠️ <b>[${name}]</b> Gặp CAPTCHA trên Shopee — cần giải để tiếp tục cào.\n` +
                    `🔗 ${window.location.href}\n🕒 ${new Date().toLocaleString()}`;
        sendTelegram(msg);
    }

    function showCaptchaBanner() {
        if (document.getElementById(`${NS}-captcha-banner`)) return;
        const banner = document.createElement('div');
        banner.id = `${NS}-captcha-banner`;
        banner.style.cssText = `position:fixed; top:0; left:0; width:100%; z-index:9999999; background:#ee4d2d; color:#fff; padding:14px; text-align:center; font-size:15px; font-weight:bold; box-shadow:0 2px 10px rgba(0,0,0,0.3);`;
        banner.innerText = '⚠️ PHÁT HIỆN CAPTCHA! Hãy giải để tool tiếp tục. (Đã gửi thông báo Telegram)';
        (document.body || document.documentElement).appendChild(banner);
    }

    function handleCaptchaPage() {
        try { window.focus(); } catch (e) { }
        notifyCaptcha();
        if (document.body) showCaptchaBanner();
        else window.addEventListener('DOMContentLoaded', showCaptchaBanner);
    }

    // ============================================================
    // WORKER: hook fetch/XHR bắt get_pc, lưu lại rồi đóng tab
    // ============================================================
    function runWorkerHook() {
        // Chỉ hoạt động khi tab được mở bởi tool (tránh hook trang user đang xem thủ công).
        // expecting là DANH SÁCH itemId đang chờ (hỗ trợ mở nhiều tab song song).
        const expectingRaw = GM_getValue(`${NS}_expecting`, []);
        const expectingList = (Array.isArray(expectingRaw) ? expectingRaw : [expectingRaw]).map(String);
        const pathMatch = window.location.pathname.match(/\/product\/(\d+)\/(\d+)/) ||
                          window.location.pathname.match(/i\.(\d+)\.(\d+)/);
        const itemId = pathMatch ? pathMatch[2] : null;
        if (!itemId || !expectingList.includes(String(itemId))) {
            return; // không phải tab do tool mở -> bỏ qua, không can thiệp
        }

        const win = (typeof unsafeWindow !== 'undefined') ? unsafeWindow : window;
        const store = (text) => {
            try {
                const json = JSON.parse(text);
                if (json && json.data && json.data.item && json.data.item.title) {
                    GM_setValue(`${NS}_captured_${itemId}`, { responseText: text, ts: Date.now() });
                    setTimeout(() => { try { window.close(); } catch (e) { } }, 400);
                }
            } catch (e) { /* bỏ qua response không phải JSON hợp lệ */ }
        };

        // Hook fetch
        const origFetch = win.fetch;
        win.fetch = async function (...args) {
            const res = await origFetch.apply(this, args);
            try {
                const url = typeof args[0] === 'string' ? args[0] : (args[0] && args[0].url) || '';
                if (url.includes(PDP_API)) {
                    res.clone().text().then(store).catch(() => { });
                }
            } catch (e) { }
            return res;
        };

        // Hook XHR
        const origOpen = win.XMLHttpRequest.prototype.open;
        win.XMLHttpRequest.prototype.open = function (method, url, ...rest) {
            const u = String(url);
            this.addEventListener('readystatechange', function () {
                if (this.readyState === 4 && u.includes(PDP_API)) {
                    store(this.responseText);
                }
            });
            return origOpen.apply(this, [method, url, ...rest]);
        };
    }

    // ============================================================
    // MAP: get_pc JSON -> item VideoAI
    // ============================================================
    function buildVideoAIItem(json, productUrl) {
        const data = json && json.data;
        if (!data || !data.item) return null;
        const item = data.item;
        const shop = data.shop_detailed || {};
        const review = data.product_review || {};

        // Ảnh: ưu tiên product_images.images, fallback item.images / item.image
        let rawImgs = (data.product_images && data.product_images.images) || item.images || [];
        if ((!rawImgs || rawImgs.length === 0) && item.image) rawImgs = [item.image];
        const images = (rawImgs || []).map(imgUrl).filter(Boolean);

        // Giá (đơn vị micro -> chia 100000); sản phẩm có biến thể dùng price_min
        let priceRaw = num(item.price);
        if (priceRaw <= 0) priceRaw = num(item.price_min);
        const price = priceRaw / 100000;

        let origRaw = num(item.price_before_discount);
        if (origRaw <= 0) origRaw = num(item.price_max_before_discount);
        const originalPrice = origRaw / 100000;

        // Tồn kho: item.stock -> normal_stock -> tổng models
        let stock = item.stock != null ? num(item.stock) : (item.normal_stock != null ? num(item.normal_stock) : 0);
        if (!stock && Array.isArray(item.models)) {
            stock = item.models.reduce((s, m) => s + num(m.stock || m.normal_stock), 0);
        }

        // Đã bán
        const soldCount = num(item.historical_sold || item.global_sold_count || review.historical_sold || item.sold || 0);

        const ratingCountArr = review.rating_count || (item.item_rating && item.item_rating.rating_count) || [];
        const ratingCount = Array.isArray(ratingCountArr) && ratingCountArr.length ? num(ratingCountArr[0]) : 0;
        const ratingStar = review.rating_star || (item.item_rating && item.item_rating.rating_star) || 0;

        const out = {
            url: productUrl,
            title: item.title || item.name || '',
            price: price,
            rating: Math.round(num(ratingStar) * 100) / 100,
            soldCount: soldCount,
            stock: stock,
            shopName: shop.name || '',
            images: images,
            reviewCount: num(review.cmt_count),
            ratingCount: ratingCount
        };

        if (originalPrice > 0) out.originalPrice = originalPrice;
        if (item.description) out.description = String(item.description);

        const feCats = item.fe_categories || [];
        if (feCats.length) {
            const names = feCats.map(c => c.display_name || c.name).filter(Boolean);
            const ids = feCats.map(c => num(c.catid)).filter(n => n > 0);
            if (names.length) out.categories = names;
            if (ids.length) out.categoryIds = ids;
        }

        const attrs = item.attributes || [];
        if (attrs.length) {
            const obj = {};
            attrs.forEach(a => { if (a && a.name) obj[a.name] = a.value != null ? String(a.value) : ''; });
            if (Object.keys(obj).length) out.attributes = obj;
        }

        return out;
    }

    // ============================================================
    // ĐẨY VIDEOAI (chia lô 200)
    // ============================================================
    async function pushToVideoAI(items, apiKey, endpoint, log) {
        let upserted = 0, skipped = 0;
        for (let i = 0; i < items.length; i += MAX_BATCH) {
            const chunk = items.slice(i, i + MAX_BATCH);
            const res = await gmRequest({
                method: 'POST',
                url: endpoint,
                headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${apiKey}` },
                data: JSON.stringify({ items: chunk })
            });
            if (res.status === 401) throw new Error('API Key không hợp lệ (401)');
            if (res.status !== 200) throw new Error(`VideoAI API lỗi: Status ${res.status}`);
            let body = {};
            try { body = JSON.parse(res.responseText); } catch (e) { }
            upserted += num(body.upserted);
            skipped += num(body.skipped);
            if (Array.isArray(body.errors)) {
                body.errors.slice(0, 10).forEach(e => log(`VideoAI skip: ${e.url} — ${e.reason}`, 'warn'));
            }
        }
        return { upserted, skipped };
    }

    // ============================================================
    // CONTROL PANEL
    // ============================================================
    function runControlPanel() {
        const mount = () => {
            if (document.getElementById(PANEL_ID)) return;
            if (!document.body) return;
            buildPanel();
        };
        if (document.readyState === 'complete' || document.readyState === 'interactive') mount();
        else window.addEventListener('DOMContentLoaded', mount);

        // Tự gắn lại nếu SPA xoá mất panel
        const obs = new MutationObserver(() => { if (!document.getElementById(PANEL_ID) && document.body) buildPanel(); });
        obs.observe(document.documentElement, { childList: true, subtree: true });

        try { GM_registerMenuCommand('Mở bảng PDP → VideoAI', () => { const p = document.getElementById(PANEL_ID); if (p) p.style.display = 'block'; }); } catch (e) { }
    }

    function buildPanel() {
        if (document.getElementById(PANEL_ID)) return;
        const savedKey = GM_getValue(`${NS}_apikey`, '');
        const savedLinks = GM_getValue(`${NS}_links`, '');
        const savedTabs = String(GM_getValue(`${NS}_tabs`, '1'));
        const savedTgToken = GM_getValue(`${NS}_tg_token`, '');
        const savedTgChat = GM_getValue(`${NS}_tg_chat`, '');
        const savedTgName = GM_getValue(`${NS}_tg_name`, '');

        const panel = document.createElement('div');
        panel.id = PANEL_ID;
        panel.style.cssText = `position: fixed; top: 80px; right: 20px; z-index: 999999; width: 320px; background: #1a1a1a; color: #fff; border: 1px solid #ee4d2d; border-radius: 8px; padding: 14px; font-family: Arial, sans-serif; font-size: 13px; box-shadow: 0 4px 15px rgba(0,0,0,0.5);`;
        panel.innerHTML = `
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">
                <strong style="color:#ee4d2d;">PDP → VideoAI</strong>
                <span id="${NS}-close" style="cursor:pointer; color:#888;">✕</span>
            </div>
            <label style="display:block; font-size:11px; margin-bottom:3px;">VideoAI API Key:</label>
            <input type="password" id="${NS}-key" placeholder="api_xxxxxxxx" value="${savedKey}" style="width:95%; padding:5px; margin-bottom:8px; border-radius:4px; border:1px solid #444; background:#2b2b2b; color:#fff; font-size:11px;">
            <label style="display:block; font-size:11px; margin-bottom:3px;">Link sản phẩm (mỗi dòng 1 link):</label>
            <textarea id="${NS}-links" rows="6" placeholder="https://shopee.vn/product/714557060/20069873884" style="width:95%; padding:6px; margin-bottom:8px; border-radius:4px; border:1px solid #444; background:#2b2b2b; color:#fff; font-size:11px; font-family:monospace; resize:vertical;">${savedLinks}</textarea>
            <label style="display:block; font-size:11px; margin-bottom:3px;">Số tab song song:</label>
            <select id="${NS}-tabs" style="width:100%; padding:5px; margin-bottom:8px; border-radius:4px; border:1px solid #444; background:#2b2b2b; color:#fff; font-size:11px;">
                <option value="1" ${savedTabs === '1' ? 'selected' : ''}>1 tab — an toàn nhất (khuyến nghị)</option>
                <option value="2" ${savedTabs === '2' ? 'selected' : ''}>2 tab — nhanh hơn</option>
                <option value="3" ${savedTabs === '3' ? 'selected' : ''}>3 tab — nhanh nhất, rủi ro captcha cao hơn</option>
            </select>
            <div style="border:1px solid #333; border-radius:4px; margin-bottom:8px;">
                <div id="${NS}-tg-toggle" style="background:#2b2b2b; padding:6px 10px; font-weight:bold; cursor:pointer; display:flex; justify-content:space-between; border-radius:4px 4px 0 0;">
                    <span>🔔 Thông báo Telegram (captcha)</span><span id="${NS}-tg-arrow">▼</span>
                </div>
                <div id="${NS}-tg-fields" style="padding:8px; display:none; background:#1f1f1f;">
                    <label style="display:block; font-size:11px; margin-bottom:3px;">Bot Token:</label>
                    <input type="password" id="${NS}-tg-token" placeholder="123456:ABC..." value="${savedTgToken}" style="width:95%; padding:5px; margin-bottom:6px; border-radius:4px; border:1px solid #444; background:#2b2b2b; color:#fff; font-size:11px;">
                    <label style="display:block; font-size:11px; margin-bottom:3px;">Chat ID (userid):</label>
                    <input type="text" id="${NS}-tg-chat" placeholder="vd: 123456789" value="${savedTgChat}" style="width:95%; padding:5px; margin-bottom:6px; border-radius:4px; border:1px solid #444; background:#2b2b2b; color:#fff; font-size:11px;">
                    <label style="display:block; font-size:11px; margin-bottom:3px;">Tên tab (nhận diện khi báo):</label>
                    <input type="text" id="${NS}-tg-name" placeholder="vd: Máy 1 - PH" value="${savedTgName}" style="width:95%; padding:5px; margin-bottom:6px; border-radius:4px; border:1px solid #444; background:#2b2b2b; color:#fff; font-size:11px;">
                    <button id="${NS}-tg-test" style="padding:4px 8px; border:none; border-radius:3px; background:#6a1b9a; color:#fff; cursor:pointer; font-size:11px; font-weight:bold;">Test Telegram</button>
                </div>
            </div>
            <div style="display:flex; gap:6px; margin-bottom:8px;">
                <button id="${NS}-start" style="flex:1; padding:8px; border:none; border-radius:4px; background:#ee4d2d; color:#fff; font-weight:bold; cursor:pointer;">Bắt đầu</button>
                <button id="${NS}-stop" style="flex:1; padding:8px; border:none; border-radius:4px; background:#555; color:#fff; font-weight:bold; cursor:pointer;" disabled>Dừng</button>
            </div>
            <div style="display:flex; justify-content:space-between; align-items:center; font-size:11px; margin-bottom:6px;">
                <span>Tiến độ: <strong id="${NS}-progress" style="color:#4caf50;">0/0</strong></span>
                <span>Đã xử lý: <strong id="${NS}-cache" style="color:#4caf50;">0</strong> <button id="${NS}-clear-cache" style="margin-left:4px; padding:2px 6px; border:none; border-radius:3px; background:#444; color:#ccc; cursor:pointer; font-size:10px;">Xóa</button></span>
            </div>
            <div id="${NS}-log" style="height:130px; overflow-y:auto; background:#0c0c0c; border:1px solid #333; padding:6px; font-family:monospace; font-size:11px; border-radius:4px; word-break:break-all;">
                <div style="color:#888;">Sẵn sàng.</div>
            </div>
        `;
        document.body.appendChild(panel);

        document.getElementById(`${NS}-close`).addEventListener('click', () => { panel.style.display = 'none'; });
        document.getElementById(`${NS}-start`).addEventListener('click', startRun);
        document.getElementById(`${NS}-stop`).addEventListener('click', () => { running = false; GM_setValue(`${NS}_session`, null); });

        updateCacheCount();
        document.getElementById(`${NS}-clear-cache`).addEventListener('click', () => {
            if (confirm('Xóa toàn bộ cache itemID đã xử lý?')) {
                doneCache = [];
                GM_setValue(`${NS}_done`, doneCache);
                updateCacheCount();
                log('Đã xóa cache itemID đã xử lý.', 'warn');
            }
        });

        // Telegram: mở/đóng, lưu cấu hình khi rời ô, nút Test
        document.getElementById(`${NS}-tg-toggle`).addEventListener('click', () => {
            const f = document.getElementById(`${NS}-tg-fields`);
            const a = document.getElementById(`${NS}-tg-arrow`);
            const open = f.style.display === 'none';
            f.style.display = open ? 'block' : 'none';
            a.innerText = open ? '▲' : '▼';
        });
        const saveTg = () => {
            GM_setValue(`${NS}_tg_token`, document.getElementById(`${NS}-tg-token`).value.trim());
            GM_setValue(`${NS}_tg_chat`, document.getElementById(`${NS}-tg-chat`).value.trim());
            GM_setValue(`${NS}_tg_name`, document.getElementById(`${NS}-tg-name`).value.trim());
        };
        ['tg-token', 'tg-chat', 'tg-name'].forEach(id => {
            document.getElementById(`${NS}-${id}`).addEventListener('blur', saveTg);
        });
        document.getElementById(`${NS}-tg-test`).addEventListener('click', async () => {
            saveTg();
            const name = (document.getElementById(`${NS}-tg-name`).value || 'PDP→VideoAI').trim();
            log('Đang gửi tin nhắn test Telegram...', 'info');
            const r = await sendTelegram(`✅ <b>[${name}]</b> Test thông báo Telegram thành công!`);
            if (r.ok) { log('Gửi Telegram thành công!', 'success'); }
            else { log(`Gửi Telegram thất bại: ${r.reason || ('HTTP ' + r.status)}`, 'error'); }
        });

        // Tự loại link trùng ngay khi dán hoặc khi rời ô nhập
        const linksTa = document.getElementById(`${NS}-links`);
        const runDedupe = () => {
            const { text, removed } = dedupeLinksText(linksTa.value);
            if (removed > 0) {
                linksTa.value = text;
                GM_setValue(`${NS}_links`, text);
                log(`Đã loại ${removed} link trùng.`, 'warn');
            }
        };
        linksTa.addEventListener('paste', () => setTimeout(runDedupe, 50));
        linksTa.addEventListener('blur', runDedupe);

        maybeResume();
    }

    // Tự khôi phục phiên nếu trang bị reload giữa chừng (session còn active).
    // Vị trí được khôi phục nhờ cache: link đã đẩy thành công sẽ bị bỏ qua, chỉ chạy tiếp phần còn lại.
    function maybeResume() {
        if (resumeChecked) return;
        resumeChecked = true;
        const session = GM_getValue(`${NS}_session`, null);
        if (!session || !session.active) return;
        if (Date.now() - (session.ts || 0) > SESSION_MAX_AGE_MS) { GM_setValue(`${NS}_session`, null); return; } // quá cũ -> bỏ
        log('Phát hiện phiên đang chạy dở (trang đã reload). Tự khôi phục sau 3s — bấm Dừng để huỷ.', 'warn');
        setTimeout(() => {
            const s = GM_getValue(`${NS}_session`, null);
            if (s && s.active && !running) startRun();
        }, RESUME_DELAY_MS);
    }

    function log(msg, type = 'info') {
        const box = document.getElementById(`${NS}-log`);
        if (!box) { console.log(`[PDP→VideoAI] ${msg}`); return; }
        const colors = { info: '#2196f3', success: '#4caf50', warn: '#ffeb3b', error: '#f44336' };
        const line = document.createElement('div');
        line.style.cssText = `margin-bottom:3px; color:${colors[type] || '#ccc'};`;
        line.innerHTML = `<span style="color:#777;">[${new Date().toLocaleTimeString()}]</span> ${msg}`;
        box.appendChild(line);
        box.scrollTop = box.scrollHeight;
    }

    function setProgress(done, total) {
        const el = document.getElementById(`${NS}-progress`);
        if (el) el.innerText = `${done}/${total}`;
    }

    function updateCacheCount() {
        const el = document.getElementById(`${NS}-cache`);
        if (el) el.innerText = doneCache.length;
    }

    // Đánh dấu 1 itemID đã xử lý xong (đẩy thành công HOẶC không đạt chất lượng) -> không làm lại
    function markDone(itemId) {
        const id = String(itemId);
        if (!doneCache.includes(id)) {
            doneCache.push(id);
            GM_setValue(`${NS}_done`, doneCache);
            updateCacheCount();
        }
    }

    function setButtons(isRunning) {
        const s = document.getElementById(`${NS}-start`);
        const st = document.getElementById(`${NS}-stop`);
        if (s) { s.disabled = isRunning; s.style.opacity = isRunning ? '0.5' : '1'; }
        if (st) { st.disabled = !isRunning; st.style.background = isRunning ? '#d32f2f' : '#555'; }
    }

    // Một lần thử: mở tab, chờ get_pc. Tự thêm/bớt itemId khỏi tập chờ (hỗ trợ chạy song song).
    async function captureOnce(link, active) {
        const itemId = link.itemId;
        GM_deleteValue(`${NS}_captured_${itemId}`);
        expectingSet.add(itemId);
        syncExpecting();

        const tab = GM_openInTab(link.productUrl, { active, insert: true });
        try {
            let waited = 0;
            const step = 300;
            while (waited < CAPTURE_TIMEOUT_MS) {
                if (!running) return null;
                const cap = GM_getValue(`${NS}_captured_${itemId}`, null);
                if (cap) {
                    GM_deleteValue(`${NS}_captured_${itemId}`);
                    return cap.responseText;
                }
                await sleep(step);
                waited += step;
            }
            return null;
        } finally {
            try { tab.close(); } catch (e) { }
            expectingSet.delete(itemId);
            syncExpecting();
        }
    }

    // Mở tab worker, chờ bắt được get_pc (lần 1 tab ngầm, timeout thì thử lại 1 lần tab nổi)
    async function captureProduct(link) {
        let text = await captureOnce(link, false);
        if (!text && running) {
            log(`Timeout SP ${link.itemId}, thử lại bằng tab nổi...`, 'warn');
            await sleep(1500);
            text = await captureOnce(link, true);
        }
        if (!text && running) log(`Bỏ qua SP ${link.itemId} (timeout sau khi thử lại).`, 'error');
        return text;
    }

    async function startRun() {
        if (running) return;

        const apiKey = (document.getElementById(`${NS}-key`).value || '').trim();
        const rawLinks = document.getElementById(`${NS}-links`).value || '';
        const concurrency = Math.max(1, Math.min(3, parseInt(document.getElementById(`${NS}-tabs`).value, 10) || 1));
        GM_setValue(`${NS}_apikey`, apiKey);
        GM_setValue(`${NS}_links`, rawLinks);
        GM_setValue(`${NS}_tabs`, String(concurrency));

        if (!apiKey) { alert('Vui lòng nhập VideoAI API Key!'); return; }

        // Parse + khử trùng + bỏ qua itemID đã xử lý (cache)
        const seen = new Set();
        const links = [];
        let invalid = 0;
        let cachedSkip = 0;
        rawLinks.split('\n').map(l => l.trim()).filter(Boolean).forEach(l => {
            const p = parseLink(l);
            if (!p) { invalid++; return; }
            if (seen.has(p.itemId)) return;
            seen.add(p.itemId);
            if (doneCache.includes(p.itemId)) { cachedSkip++; return; }
            links.push(p);
        });

        if (invalid) log(`Bỏ qua ${invalid} dòng link không hợp lệ.`, 'warn');
        if (cachedSkip) log(`Bỏ qua ${cachedSkip} link đã xử lý trước đó (cache).`, 'info');
        if (links.length === 0) {
            GM_setValue(`${NS}_session`, null);
            if (cachedSkip > 0) log('Tất cả link đã được xử lý trước đó (cache).', 'success');
            else alert('Không có link hợp lệ!');
            return;
        }

        running = true;
        setButtons(true);
        GM_setValue(`${NS}_session`, { active: true, ts: Date.now() }); // đánh dấu phiên đang chạy để resume nếu reload
        log(`Bắt đầu xử lý ${links.length} link (${concurrency} tab song song)...`, 'success');
        if (concurrency > 1) log(`Lưu ý: mở ${concurrency} tab cùng lúc tăng tốc nhưng dễ gặp captcha/giới hạn hơn.`, 'warn');
        setProgress(0, links.length);

        const collected = [];
        const collectedIds = []; // itemId song song với collected, để ghi cache sau khi đẩy thành công
        let tooFewImages = 0;
        let done = 0;
        let totalUpserted = 0;
        let totalSkipped = 0;

        // Đẩy buffer hiện có lên VideoAI; chỉ xoá buffer + ghi cache khi đẩy thành công (giữ lại để thử lại nếu lỗi)
        async function flush(isFinal = false) {
            if (collected.length === 0) return;
            const n = collected.length;
            log(`Đẩy ${n} SP lên VideoAI${isFinal ? ' (đợt cuối)' : ''}...`, 'info');
            try {
                const r = await pushToVideoAI(collected, apiKey, VIDEOAI_DEFAULT_ENDPOINT, log);
                totalUpserted += r.upserted;
                totalSkipped += r.skipped;
                collectedIds.forEach(id => { if (!doneCache.includes(id)) doneCache.push(id); });
                GM_setValue(`${NS}_done`, doneCache);
                updateCacheCount();
                collected.length = 0;
                collectedIds.length = 0;
                log(`Đã đẩy đợt: ghi ${r.upserted}, bỏ qua ${r.skipped}.`, 'success');
            } catch (e) {
                log(`Lỗi đẩy VideoAI: ${e.message}. Giữ ${n} SP để thử lại sau.`, 'error');
            }
        }

        for (let i = 0; i < links.length; i += concurrency) {
            if (!running) break;
            const batch = links.slice(i, i + concurrency);
            log(`Mở ${batch.length} tab: SP ${batch.map(b => b.itemId).join(', ')}...`, 'info');

            const results = await Promise.all(batch.map(link => captureProduct(link)));

            for (let j = 0; j < batch.length; j++) {
                done++;
                setProgress(done, links.length);
                const text = results[j];
                const link = batch[j];
                if (!text) continue; // timeout/bắt hụt -> KHÔNG cache (cho phép thử lại sau)

                let item;
                try { item = buildVideoAIItem(JSON.parse(text), link.productUrl); }
                catch (e) { markDone(link.itemId); log(`Lỗi parse dữ liệu SP ${link.itemId}: ${e.message} — đánh dấu đã xử lý.`, 'error'); continue; }

                if (!item) { markDone(link.itemId); log(`Không có dữ liệu hợp lệ cho SP ${link.itemId} — đánh dấu đã xử lý.`, 'warn'); continue; }
                if (item.images.length < MIN_IMAGES) { tooFewImages++; markDone(link.itemId); log(`SP ${link.itemId} chỉ có ${item.images.length} ảnh (<${MIN_IMAGES}) — bỏ qua & đánh dấu đã xử lý.`, 'warn'); continue; }

                collected.push(item);
                collectedIds.push(link.itemId);
                log(`✓ ${item.title.substring(0, 30)}... | ${item.price}₫ | ${item.images.length} ảnh`, 'success');
            }

            // Chu kỳ đẩy: gom đủ ngưỡng thì đẩy ngay (chống mất dữ liệu nếu gián đoạn)
            if (collected.length >= FLUSH_THRESHOLD) await flush(false);

            if (running) await sleep(800 + Math.random() * 800); // nghỉ nhẹ giữa các lô
        }

        if (tooFewImages) log(`Đã bỏ ${tooFewImages} SP do <${MIN_IMAGES} ảnh.`, 'warn');

        // Đẩy nốt phần còn lại (kể cả khi người dùng bấm Dừng giữa chừng)
        await flush(true);

        if (totalUpserted + totalSkipped > 0 || collected.length === 0) {
            log(`HOÀN TẤT! Tổng ghi: ${totalUpserted}, bỏ qua: ${totalSkipped}.`, 'success');
        }
        if (collected.length > 0) {
            log(`Còn ${collected.length} SP chưa đẩy được (lỗi mạng/API) — bấm Bắt đầu lại để thử.`, 'warn');
        }

        running = false;
        setButtons(false);
        GM_deleteValue(`${NS}_expecting`);
        // Chạy đến cuối (hoàn tất hoặc bấm Dừng) -> kết thúc phiên. Chỉ reload giữa chừng mới giữ session để auto-resume.
        GM_setValue(`${NS}_session`, null);
    }
})();
