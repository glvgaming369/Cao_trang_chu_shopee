// Chạy trong ISOLATED world (có chrome.*). Hai nhiệm vụ:
// 1) Nhận dữ liệu get_pc do pageHook.js (MAIN world) postMessage -> chuyển về background.
// 2) Phát hiện captcha (URL /verify hoặc DOM) -> báo background (để gửi Telegram) + hiện banner.
(function () {
    'use strict';

    // 1) Relay captured data
    window.addEventListener('message', (ev) => {
        if (ev.source !== window) return;
        const d = ev.data;
        if (d && d.__pdpvai && d.kind === 'captured') {
            try {
                chrome.runtime.sendMessage({ type: 'captured', itemId: String(d.itemId), responseText: d.responseText });
            } catch (e) { /* context invalidated khi extension reload -> bỏ qua */ }
        }
    });

    // 2) Captcha
    function isCaptcha() {
        const h = location.href;
        if (h.includes('/verify/captcha') || h.includes('/verify/traffic') || h.includes('/verify/security')) return true;
        return !!document.querySelector('.check-captcha-box') ||
               !!document.querySelector('.shopee-captcha-wrapper') ||
               !!document.getElementById('captcha-submit') ||
               !!document.querySelector('iframe[src*="captcha"]');
    }

    function showBanner() {
        if (document.getElementById('pdpvai-captcha-banner')) return;
        const b = document.createElement('div');
        b.id = 'pdpvai-captcha-banner';
        b.style.cssText = 'position:fixed;top:0;left:0;width:100%;z-index:2147483647;background:#ee4d2d;color:#fff;padding:14px;text-align:center;font:bold 15px Arial;box-shadow:0 2px 10px rgba(0,0,0,.3)';
        b.textContent = '⚠️ PHÁT HIỆN CAPTCHA! Hãy giải để tool tiếp tục cào.';
        (document.body || document.documentElement).appendChild(b);
    }

    function report() {
        try { chrome.runtime.sendMessage({ type: 'captcha', url: location.href }); } catch (e) { }
        if (document.body) showBanner();
        else window.addEventListener('DOMContentLoaded', showBanner);
    }

    if (isCaptcha()) {
        report();
    } else {
        // Captcha kiểu DOM có thể chèn sau khi tải xong
        window.addEventListener('load', () => setTimeout(() => { if (isCaptcha()) report(); }, 2500));
    }
})();
