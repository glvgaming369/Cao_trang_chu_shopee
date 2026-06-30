// Service worker (MV3): sở hữu toàn bộ state, mở tab, gọi VideoAI/Telegram, cache, resume.
import {
    MIN_IMAGES, MAX_BATCH, VIDEOAI_DEFAULT_ENDPOINT,
    num, parseLink, buildVideoAIItem
} from './lib/shared.js';

const CAPTURE_TIMEOUT_MS = 15000;
const FLUSH_THRESHOLD = 30;
const TG_COOLDOWN_MS = 60000;
const SESSION_MAX_AGE_MS = 6 * 3600 * 1000;
const LOG_CAP = 200;
const JOBS_ENDPOINT = 'https://videoai-api.devappnow.com/api/v1/jobs';
const BULK_STATUS_ENDPOINT = 'https://videoai-api.devappnow.com/api/v1/shopee-video/bulk-status';
const JOBS_PAGE_SIZE = 50;

const sleep = (ms) => new Promise(r => setTimeout(r, ms));

// ---------- STATE (bộ nhớ SW) ----------
let running = false;
let progress = { done: 0, total: 0 };
let logs = [];                 // ring buffer cho popup
let doneCache = [];            // itemID đã xử lý (đồng bộ storage 'done')
const pending = new Map();     // itemId -> finish(text|null)

// ---------- STORAGE helpers ----------
const store = {
    get: (k, d) => new Promise(r => chrome.storage.local.get({ [k]: d }, o => r(o[k]))),
    set: (k, v) => new Promise(r => chrome.storage.local.set({ [k]: v }, r)),
};

function log(msg, type = 'info') {
    logs.push({ ts: Date.now(), msg, type });
    if (logs.length > LOG_CAP) logs = logs.slice(-LOG_CAP);
    console.log(`[PDP→VideoAI][${type}] ${msg}`);
}

// ---------- NETWORK (chạy ở SW -> không vướng CORS) ----------
async function pushToVideoAI(items, apiKey, endpoint) {
    let upserted = 0, skipped = 0;
    for (let i = 0; i < items.length; i += MAX_BATCH) {
        const chunk = items.slice(i, i + MAX_BATCH);
        const res = await fetch(endpoint, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${apiKey}` },
            body: JSON.stringify({ items: chunk })
        });
        if (res.status === 401) throw new Error('API Key không hợp lệ (401)');
        if (res.status !== 200) throw new Error(`VideoAI API lỗi: Status ${res.status}`);
        let body = {};
        try { body = await res.json(); } catch (e) { }
        upserted += num(body.upserted);
        skipped += num(body.skipped);
        if (Array.isArray(body.errors)) {
            body.errors.slice(0, 10).forEach(e => log(`VideoAI skip: ${e.url} — ${e.reason}`, 'warn'));
        }
    }
    return { upserted, skipped };
}

async function sendTelegram(text) {
    const cfg = await store.get('config', {});
    const token = (cfg.tg_token || '').trim();
    const chat = (cfg.tg_chat || '').trim();
    if (!token || !chat) return { ok: false, reason: 'Chưa cấu hình token/chat id' };
    try {
        const res = await fetch(`https://api.telegram.org/bot${token}/sendMessage`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ chat_id: chat, text, parse_mode: 'HTML', disable_web_page_preview: true })
        });
        return { ok: res.status === 200, status: res.status };
    } catch (e) {
        return { ok: false, reason: e.message };
    }
}

// Gửi ảnh (multipart) kèm caption lên Telegram
async function sendTelegramPhoto(blob, caption) {
    const cfg = await store.get('config', {});
    const token = (cfg.tg_token || '').trim();
    const chat = (cfg.tg_chat || '').trim();
    if (!token || !chat) return { ok: false, reason: 'Chưa cấu hình token/chat id' };
    const fd = new FormData();
    fd.append('chat_id', chat);
    fd.append('caption', caption);
    fd.append('parse_mode', 'HTML');
    fd.append('photo', blob, 'captcha.jpg');
    try {
        const res = await fetch(`https://api.telegram.org/bot${token}/sendPhoto`, { method: 'POST', body: fd });
        return { ok: res.status === 200, status: res.status };
    } catch (e) {
        return { ok: false, reason: e.message };
    }
}

function dataUrlToBlob(dataUrl) {
    const [meta, b64] = dataUrl.split(',');
    const mime = (meta.match(/:(.*?);/) || [, 'image/jpeg'])[1];
    const bin = atob(b64);
    const arr = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
    return new Blob([arr], { type: mime });
}

async function onCaptcha(url, tab) {
    const cfg = await store.get('config', {});
    if (!cfg.tg_token || !cfg.tg_chat) return; // chưa cấu hình -> bỏ qua
    const last = await store.get('tg_last', 0);
    if (Date.now() - last < TG_COOLDOWN_MS) return;
    await store.set('tg_last', Date.now());
    const name = (cfg.tg_name || 'PDP→VideoAI').trim();
    const caption = `⚠️ <b>[${name}]</b> Gặp CAPTCHA trên Shopee — cần giải để tiếp tục cào.\n` +
                    `🔗 ${url}\n🕒 ${new Date().toLocaleString()}`;
    log('Phát hiện CAPTCHA — chụp màn hình gửi Telegram...', 'warn');

    // Chụp ảnh tab captcha: phải đưa tab ra trước (cũng giúp người dùng thấy để giải)
    let photoSent = false;
    try {
        if (tab && tab.id != null) {
            await chrome.tabs.update(tab.id, { active: true });
            await sleep(1200); // chờ captcha render xong
            const dataUrl = await chrome.tabs.captureVisibleTab(tab.windowId, { format: 'jpeg', quality: 70 });
            const r = await sendTelegramPhoto(dataUrlToBlob(dataUrl), caption);
            photoSent = r.ok;
            if (!r.ok) log(`Gửi ảnh Telegram thất bại (${r.reason || r.status}).`, 'warn');
        }
    } catch (e) {
        log(`Không chụp được màn hình captcha: ${e.message}`, 'warn');
    }
    if (!photoSent) await sendTelegram(caption); // fallback: chỉ gửi chữ
}

// ---------- CAPTURE (event-driven) ----------
function captureOnce(link, active) {
    return new Promise((resolve) => {
        let settled = false;
        chrome.tabs.create({ url: link.productUrl, active }).then((tab) => {
            const itemId = String(link.itemId);
            const finish = (text) => {
                if (settled) return;
                settled = true;
                clearTimeout(timer);
                pending.delete(itemId);
                if (tab && tab.id != null) chrome.tabs.remove(tab.id).catch(() => { });
                resolve(text);
            };
            const timer = setTimeout(() => finish(null), CAPTURE_TIMEOUT_MS);
            pending.set(itemId, finish);
        }).catch(() => resolve(null));
    });
}

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

function abortPending() {
    for (const finish of pending.values()) finish(null);
    pending.clear();
}

async function markDone(ids) {
    let changed = false;
    ids.forEach(id => { const s = String(id); if (!doneCache.includes(s)) { doneCache.push(s); changed = true; } });
    if (changed) await store.set('done', doneCache);
}

// ---------- JOBS SERVER (lấy link + đánh dấu used) ----------
async function fetchJobsPage(apiKey, page) {
    const url = `${JOBS_ENDPOINT}?jobType=shopee-video&status=pending&page=${page}&pageSize=${JOBS_PAGE_SIZE}`;
    const res = await fetch(url, { headers: { 'Authorization': `Bearer ${apiKey}` } });
    if (res.status === 401) throw new Error('API Key không hợp lệ (401)');
    if (res.status !== 200) throw new Error(`Jobs API lỗi: Status ${res.status}`);
    return await res.json();
}

// Đánh dấu các itemId là useStatus="used" (cờ server, chống làm lại đa máy). Trả số bản ghi cập nhật.
async function markUsed(apiKey, itemIds) {
    if (!itemIds || itemIds.length === 0) return 0;
    const res = await fetch(BULK_STATUS_ENDPOINT, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${apiKey}` },
        body: JSON.stringify({ productIds: itemIds.map(String), useStatus: 'used' })
    });
    if (res.status !== 200) throw new Error(`bulk-status Status ${res.status}`);
    const b = await res.json().catch(() => ({}));
    return num(b.updated);
}

// ---------- ORCHESTRATOR ----------
let keepAliveTimer = null;

async function startRun() {
    if (running) return;
    const cfg = await store.get('config', {});
    const apiKey = (cfg.apiKey || '').trim();
    if (!apiKey) { log('Chưa nhập VideoAI API Key!', 'error'); return; }

    const concurrency = Math.max(1, Math.min(3, parseInt(cfg.tabs, 10) || 1));
    const endpoint = cfg.endpoint || VIDEOAI_DEFAULT_ENDPOINT;
    const source = cfg.source === 'server' ? 'server' : 'manual';
    doneCache = (await store.get('done', [])).map(String);

    running = true;
    startKeepAlive();
    await store.set('session', { active: true, ts: Date.now(), source });
    progress = { done: 0, total: 0 };
    if (concurrency > 1) log(`Lưu ý: mở ${concurrency} tab cùng lúc dễ gặp captcha/giới hạn hơn.`, 'warn');

    // Buffer + tổng kết dùng chung cho cả 2 nguồn
    const buf = [];
    const bufIds = [];
    let tooFewImages = 0, totalUpserted = 0, totalSkipped = 0;

    async function flush(isFinal = false) {
        if (buf.length === 0) return;
        const n = buf.length;
        log(`Đẩy ${n} SP lên VideoAI${isFinal ? ' (đợt cuối)' : ''}...`, 'info');
        try {
            const r = await pushToVideoAI(buf, apiKey, endpoint);
            totalUpserted += r.upserted;
            totalSkipped += r.skipped;
            const ids = bufIds.slice();
            await markDone(ids);
            if (source === 'server') {
                try { const u = await markUsed(apiKey, ids); log(`Đánh dấu ${u} job = used trên server.`, 'info'); }
                catch (e) { log(`Lỗi đánh dấu bulk-status: ${e.message}`, 'warn'); }
            }
            buf.length = 0;
            bufIds.length = 0;
            log(`Đã đẩy đợt: ghi ${r.upserted}, bỏ qua ${r.skipped}.`, 'success');
        } catch (e) {
            log(`Lỗi đẩy VideoAI: ${e.message}. Giữ ${n} SP để thử lại sau.`, 'error');
        }
    }

    // Item bị loại (lỗi/không đủ ảnh): đánh dấu đã xử lý ở local + server (mode server) để không làm lại
    async function markProcessed(itemId) {
        await markDone([itemId]);
        if (source === 'server') { try { await markUsed(apiKey, [itemId]); } catch (e) { } }
    }

    async function processLinks(links) {
        for (let i = 0; i < links.length; i += concurrency) {
            if (!running) break;
            const batch = links.slice(i, i + concurrency);
            log(`Mở ${batch.length} tab: SP ${batch.map(b => b.itemId).join(', ')}...`, 'info');
            const results = await Promise.all(batch.map(link => captureProduct(link)));

            for (let j = 0; j < batch.length; j++) {
                progress.done++;
                const text = results[j];
                const link = batch[j];
                if (!text) continue; // timeout/bắt hụt -> KHÔNG đánh dấu (cho thử lại sau)

                let item = null;
                try { item = buildVideoAIItem(JSON.parse(text), link.productUrl); }
                catch (e) { await markProcessed(link.itemId); log(`Lỗi parse SP ${link.itemId}: ${e.message} — đánh dấu đã xử lý.`, 'error'); continue; }

                if (!item) { await markProcessed(link.itemId); log(`Không có dữ liệu hợp lệ cho SP ${link.itemId} — đánh dấu đã xử lý.`, 'warn'); continue; }
                if (item.images.length < MIN_IMAGES) { tooFewImages++; await markProcessed(link.itemId); log(`SP ${link.itemId} chỉ có ${item.images.length} ảnh (<${MIN_IMAGES}) — bỏ qua & đánh dấu đã xử lý.`, 'warn'); continue; }

                buf.push(item);
                bufIds.push(link.itemId);
                log(`✓ ${String(item.title).substring(0, 30)}... | ${item.price}₫ | ${item.images.length} ảnh`, 'success');
            }

            if (buf.length >= FLUSH_THRESHOLD) await flush(false);
            if (running) await sleep(800 + Math.random() * 800);
        }
    }

    if (source === 'server') {
        await runServerJobs(apiKey, processLinks, flush);
    } else {
        await runManualLinks(cfg, processLinks);
    }

    if (tooFewImages) log(`Đã bỏ ${tooFewImages} SP do <${MIN_IMAGES} ảnh.`, 'warn');
    await flush(true);
    log(`HOÀN TẤT! Tổng ghi: ${totalUpserted}, bỏ qua: ${totalSkipped}.`, 'success');
    if (buf.length > 0) log(`Còn ${buf.length} SP chưa đẩy được — bấm Bắt đầu lại để thử.`, 'warn');

    running = false;
    stopKeepAlive();
    await store.set('session', null);
}

// Nguồn DÁN TAY: parse textarea, khử trùng, bỏ qua cache local
async function runManualLinks(cfg, processLinks) {
    const seen = new Set();
    const links = [];
    let invalid = 0, cachedSkip = 0;
    (cfg.links || '').split('\n').map(l => l.trim()).filter(Boolean).forEach(l => {
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
        log(cachedSkip > 0 ? 'Tất cả link đã được xử lý trước đó (cache).' : 'Không có link hợp lệ!', cachedSkip > 0 ? 'success' : 'error');
        return;
    }
    progress.total = links.length;
    log(`Bắt đầu xử lý ${links.length} link (dán tay)...`, 'success');
    await processLinks(links);
}

// Nguồn TỪ SERVER: kéo job pending từng trang, bỏ qua job đã used, cào -> đẩy -> đánh dấu used
async function runServerJobs(apiKey, processLinks, flush) {
    log('Chế độ Từ server: lấy job pending...', 'success');
    let page = 1;
    let processedTotal = 0;
    while (running) {
        let res;
        try { res = await fetchJobsPage(apiKey, page); }
        catch (e) { log(`Lỗi lấy job (trang ${page}): ${e.message}`, 'error'); break; }
        const items = (res && res.items) || [];
        if (items.length === 0) { log('Hết job pending.', 'success'); break; }
        progress.total = res.total || 0;

        const seen = new Set();
        const links = [];
        for (const job of items) {
            const url = job && job.data && job.data.url;
            const p = url ? parseLink(url) : null;
            if (!p) continue;
            if (job.useStatus === 'used') continue;     // đã xử lý (cờ server, đa máy)
            if (doneCache.includes(p.itemId)) continue; // đã xử lý (local)
            if (seen.has(p.itemId)) continue;
            seen.add(p.itemId);
            links.push(p);
        }

        if (links.length) {
            log(`Trang ${page}: ${links.length}/${items.length} job cần cào (tổng pending ${res.total}).`, 'info');
            await processLinks(links);
            await flush(false); // đẩy + đánh dấu used ngay sau mỗi trang
            processedTotal += links.length;
        } else {
            log(`Trang ${page}: tất cả đã xử lý, sang trang sau.`, 'info');
        }

        if (items.length < JOBS_PAGE_SIZE) { log('Đã tới trang cuối.', 'success'); break; }
        page++;
    }
    log(`Phiên server: đã xử lý ${processedTotal} job.`, 'success');
}

function stopRun() {
    running = false;
    abortPending();
    stopKeepAlive();
    store.set('session', null);
    log('Đã dừng theo yêu cầu.', 'warn');
}

// ---------- KEEP-ALIVE (giảm việc SW bị ngủ giữa chừng) ----------
function startKeepAlive() {
    if (keepAliveTimer) return;
    keepAliveTimer = setInterval(() => { chrome.runtime.getPlatformInfo(() => { }); }, 20000);
}
function stopKeepAlive() {
    if (keepAliveTimer) { clearInterval(keepAliveTimer); keepAliveTimer = null; }
}

// ---------- MESSAGE ROUTER ----------
chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
    if (!msg || !msg.type) return;
    switch (msg.type) {
        case 'captured': {
            const f = pending.get(String(msg.itemId));
            if (f) f(msg.responseText);
            return;
        }
        case 'captcha':
            onCaptcha(msg.url, sender && sender.tab);
            return;
        case 'getState':
            sendResponse({ running, progress, logs: logs.slice(-120), doneCount: doneCache.length });
            return;
        case 'start':
            startRun();
            sendResponse({ ok: true });
            return;
        case 'stop':
            stopRun();
            sendResponse({ ok: true });
            return;
        case 'clearCache':
            doneCache = [];
            store.set('done', []).then(() => sendResponse({ ok: true }));
            return true;
        case 'testTelegram':
            sendTelegram(msg.text || '✅ Test thông báo Telegram thành công!').then(sendResponse);
            return true;
        case 'clearLogs':
            logs = [];
            sendResponse({ ok: true });
            return;
    }
});

// ---------- RESUME khi SW khởi động lại / cài đặt ----------
async function resumeIfNeeded() {
    if (running) return;
    const session = await store.get('session', null);
    if (!session || !session.active) return;
    if (Date.now() - (session.ts || 0) > SESSION_MAX_AGE_MS) { await store.set('session', null); return; }
    log('Phát hiện phiên đang chạy dở — tự khôi phục (bỏ qua link đã xử lý nhờ cache).', 'warn');
    startRun();
}
// Nạp cache + thử resume khi worker thức dậy
(async () => {
    doneCache = (await store.get('done', [])).map(String);
    resumeIfNeeded();
})();
chrome.runtime.onStartup && chrome.runtime.onStartup.addListener(resumeIfNeeded);
