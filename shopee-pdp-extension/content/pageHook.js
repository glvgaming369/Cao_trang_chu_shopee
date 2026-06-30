// Chạy trong MAIN world của trang Shopee (document_start) để override window.fetch / XHR
// và bắt response của API get_pc. Không truy cập được chrome.* -> chuyển dữ liệu ra qua window.postMessage,
// content/bridge.js (isolated) sẽ nhận và gửi về background.
(function () {
    'use strict';
    const PDP_API = '/api/v4/pdp/get_pc';

    function itemIdFromPath() {
        const m = location.pathname.match(/\/product\/(\d+)\/(\d+)/) ||
                  location.pathname.match(/i\.(\d+)\.(\d+)/);
        return m ? m[2] : null;
    }

    function post(text) {
        const itemId = itemIdFromPath();
        if (!itemId) return;
        try {
            const json = JSON.parse(text);
            if (json && json.data && json.data.item && json.data.item.title) {
                window.postMessage({ __pdpvai: true, kind: 'captured', itemId, responseText: text }, '*');
            }
        } catch (e) { /* response không phải JSON hợp lệ -> bỏ qua */ }
    }

    const origFetch = window.fetch;
    window.fetch = async function (...args) {
        const res = await origFetch.apply(this, args);
        try {
            const url = typeof args[0] === 'string' ? args[0] : (args[0] && args[0].url) || '';
            if (url.includes(PDP_API)) res.clone().text().then(post).catch(() => { });
        } catch (e) { }
        return res;
    };

    const origOpen = window.XMLHttpRequest.prototype.open;
    window.XMLHttpRequest.prototype.open = function (method, url, ...rest) {
        const u = String(url);
        this.addEventListener('readystatechange', function () {
            if (this.readyState === 4 && u.includes(PDP_API)) post(this.responseText);
        });
        return origOpen.apply(this, [method, url, ...rest]);
    };
})();
