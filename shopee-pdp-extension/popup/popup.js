import { dedupeLinksText } from '../lib/shared.js';

const $ = (id) => document.getElementById(id);
const CONFIG_FIELDS = ['apiKey', 'source', 'links', 'tabs', 'tg_token', 'tg_chat', 'tg_name'];

function applySourceUI() {
    $('links-wrap').style.display = $('source').value === 'server' ? 'none' : 'block';
}

function readConfig() {
    const cfg = {};
    CONFIG_FIELDS.forEach(f => { cfg[f] = $(f).value; });
    return cfg;
}

function saveConfig() {
    return new Promise(r => chrome.storage.local.set({ config: readConfig() }, r));
}

function send(type, extra = {}) {
    return new Promise(resolve => chrome.runtime.sendMessage({ type, ...extra }, resolve));
}

// ---------- Nạp cấu hình đã lưu ----------
chrome.storage.local.get({ config: {} }, ({ config }) => {
    CONFIG_FIELDS.forEach(f => { if (config[f] != null) $(f).value = config[f]; });
    if (!$('tabs').value) $('tabs').value = '1';
    if (!$('source').value) $('source').value = 'manual';
    applySourceUI();
});

$('source').addEventListener('change', () => { applySourceUI(); saveConfig(); });

// ---------- Lưu cấu hình khi thay đổi (debounce nhẹ) ----------
let saveTimer = null;
CONFIG_FIELDS.forEach(f => {
    $(f).addEventListener('input', () => {
        clearTimeout(saveTimer);
        saveTimer = setTimeout(saveConfig, 250);
    });
});

// ---------- Loại link trùng khi dán / rời ô ----------
function runDedupe() {
    const { text, removed } = dedupeLinksText($('links').value);
    if (removed > 0) {
        $('links').value = text;
        saveConfig();
    }
}
$('links').addEventListener('paste', () => setTimeout(runDedupe, 50));
$('links').addEventListener('blur', runDedupe);

// ---------- Telegram section ----------
$('tg-head').addEventListener('click', () => {
    const b = $('tg-body');
    const open = b.style.display === 'none' || !b.style.display;
    b.style.display = open ? 'block' : 'none';
    $('tg-arrow').textContent = open ? '▲' : '▼';
});
$('tg-test').addEventListener('click', async () => {
    await saveConfig();
    const name = ($('tg_name').value || 'PDP→VideoAI').trim();
    const r = await send('testTelegram', { text: `✅ <b>[${name}]</b> Test thông báo Telegram thành công!` });
    alert(r && r.ok ? 'Gửi Telegram thành công!' : `Gửi Telegram thất bại: ${(r && (r.reason || ('HTTP ' + r.status))) || 'không rõ'}`);
});

// ---------- Nút điều khiển ----------
$('start').addEventListener('click', async () => {
    if (!$('apiKey').value.trim()) { alert('Vui lòng nhập VideoAI API Key!'); return; }
    await saveConfig();
    await send('start');
});
$('stop').addEventListener('click', () => send('stop'));
$('clear-cache').addEventListener('click', async () => {
    if (confirm('Xóa toàn bộ cache itemID đã xử lý?')) await send('clearCache');
});

// ---------- Đồng bộ trạng thái từ background ----------
function renderLogs(items) {
    const box = $('log');
    const atBottom = box.scrollHeight - box.scrollTop - box.clientHeight < 20;
    box.innerHTML = (items || []).map(l => {
        const t = new Date(l.ts).toLocaleTimeString();
        return `<div class="ln t-${l.type}"><span class="time">[${t}]</span> ${escapeHtml(l.msg)}</div>`;
    }).join('') || '<div class="ln" style="color:#888;">Sẵn sàng.</div>';
    if (atBottom) box.scrollTop = box.scrollHeight;
}

function escapeHtml(s) {
    return String(s).replace(/[&<>]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
}

async function refresh() {
    const st = await send('getState');
    if (!st) return;
    $('progress').textContent = `${st.progress.done}/${st.progress.total}`;
    $('cache').textContent = st.doneCount;
    $('start').disabled = st.running;
    $('start').style.opacity = st.running ? '0.5' : '1';
    $('stop').disabled = !st.running;
    $('stop').style.background = st.running ? '#d32f2f' : '#555';
    renderLogs(st.logs);
}

refresh();
setInterval(refresh, 1000);
