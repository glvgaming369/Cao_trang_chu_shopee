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
    const CAPTURE_TIMEOUT_MS = 15000;
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
    // ROUTING: trang sản phẩm = worker (hook), còn lại = panel điều khiển
    // ============================================================
    const isProductPage = /\/product\/\d+\/\d+/.test(window.location.pathname) ||
                          /i\.\d+\.\d+/.test(window.location.pathname);

    if (isProductPage) {
        runWorkerHook();
    } else {
        runControlPanel();
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
            <div style="display:flex; gap:6px; margin-bottom:8px;">
                <button id="${NS}-start" style="flex:1; padding:8px; border:none; border-radius:4px; background:#ee4d2d; color:#fff; font-weight:bold; cursor:pointer;">Bắt đầu</button>
                <button id="${NS}-stop" style="flex:1; padding:8px; border:none; border-radius:4px; background:#555; color:#fff; font-weight:bold; cursor:pointer;" disabled>Dừng</button>
            </div>
            <div style="font-size:11px; margin-bottom:6px;">Tiến độ: <strong id="${NS}-progress" style="color:#4caf50;">0/0</strong></div>
            <div id="${NS}-log" style="height:130px; overflow-y:auto; background:#0c0c0c; border:1px solid #333; padding:6px; font-family:monospace; font-size:11px; border-radius:4px; word-break:break-all;">
                <div style="color:#888;">Sẵn sàng.</div>
            </div>
        `;
        document.body.appendChild(panel);

        document.getElementById(`${NS}-close`).addEventListener('click', () => { panel.style.display = 'none'; });
        document.getElementById(`${NS}-start`).addEventListener('click', startRun);
        document.getElementById(`${NS}-stop`).addEventListener('click', () => { running = false; });

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

    function setButtons(isRunning) {
        const s = document.getElementById(`${NS}-start`);
        const st = document.getElementById(`${NS}-stop`);
        if (s) { s.disabled = isRunning; s.style.opacity = isRunning ? '0.5' : '1'; }
        if (st) { st.disabled = !isRunning; st.style.background = isRunning ? '#d32f2f' : '#555'; }
    }

    let running = false;

    // Tập itemId đang chờ bắt (đồng bộ xuống GM để worker nhiều tab cùng nhận diện)
    const expectingSet = new Set();
    function syncExpecting() {
        GM_setValue(`${NS}_expecting`, Array.from(expectingSet));
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

        // Parse + khử trùng
        const seen = new Set();
        const links = [];
        let invalid = 0;
        rawLinks.split('\n').map(l => l.trim()).filter(Boolean).forEach(l => {
            const p = parseLink(l);
            if (!p) { invalid++; return; }
            if (seen.has(p.itemId)) return;
            seen.add(p.itemId);
            links.push(p);
        });

        if (invalid) log(`Bỏ qua ${invalid} dòng link không hợp lệ.`, 'warn');
        if (links.length === 0) { alert('Không có link hợp lệ!'); return; }

        running = true;
        setButtons(true);
        log(`Bắt đầu xử lý ${links.length} link (${concurrency} tab song song)...`, 'success');
        if (concurrency > 1) log(`Lưu ý: mở ${concurrency} tab cùng lúc tăng tốc nhưng dễ gặp captcha/giới hạn hơn.`, 'warn');
        setProgress(0, links.length);

        const collected = [];
        let tooFewImages = 0;
        let done = 0;

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
                if (!text) continue;

                let item;
                try { item = buildVideoAIItem(JSON.parse(text), link.productUrl); }
                catch (e) { log(`Lỗi parse dữ liệu SP ${link.itemId}: ${e.message}`, 'error'); continue; }

                if (!item) { log(`Không có dữ liệu hợp lệ cho SP ${link.itemId}.`, 'warn'); continue; }
                if (item.images.length < MIN_IMAGES) { tooFewImages++; log(`SP ${link.itemId} chỉ có ${item.images.length} ảnh (<${MIN_IMAGES}) — bỏ qua.`, 'warn'); continue; }

                collected.push(item);
                log(`✓ ${item.title.substring(0, 30)}... | ${item.price}₫ | ${item.images.length} ảnh`, 'success');
            }

            if (running) await sleep(800 + Math.random() * 800); // nghỉ nhẹ giữa các lô
        }

        if (tooFewImages) log(`Đã bỏ ${tooFewImages} SP do <${MIN_IMAGES} ảnh.`, 'warn');

        if (collected.length > 0) {
            log(`Đẩy ${collected.length} SP lên VideoAI...`, 'info');
            try {
                const r = await pushToVideoAI(collected, apiKey, VIDEOAI_DEFAULT_ENDPOINT, log);
                log(`HOÀN TẤT! Ghi: ${r.upserted}, bỏ qua: ${r.skipped}.`, 'success');
            } catch (e) {
                log(`Lỗi đẩy VideoAI: ${e.message}`, 'error');
            }
        } else {
            log('Không có sản phẩm hợp lệ để đẩy.', 'warn');
        }

        running = false;
        setButtons(false);
        GM_deleteValue(`${NS}_expecting`);
    }
})();
