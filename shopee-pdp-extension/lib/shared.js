// Helper thuần + mapper, dùng chung cho background (service worker) và popup.
// Không phụ thuộc DOM hay chrome.* để có thể import ở cả hai nơi.

export const MIN_IMAGES = 3;
export const MAX_BATCH = 200;
export const IMG_DOMAIN = 'down-vn';
export const PDP_API = '/api/v4/pdp/get_pc';
export const VIDEOAI_DEFAULT_ENDPOINT = 'https://videoai-api.devappnow.com/api/products/shopee-cache/batch';

export function num(v) {
    const n = Number(v);
    return isNaN(n) ? 0 : n;
}

export function imgUrl(hash) {
    if (!hash) return null;
    if (String(hash).startsWith('http')) return hash;
    return `https://${IMG_DOMAIN}.img.susercontent.com/file/${hash}`;
}

// Parse link -> { shopId, itemId, productUrl }. Hỗ trợ /product/{shop}/{item} và i.{shop}.{item}
export function parseLink(href) {
    if (!href) return null;
    let s = String(href).trim();
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

// Loại link trùng theo itemId (link hợp lệ) hoặc nội dung dòng (link lạ). Giữ thứ tự, bỏ dòng trống.
export function dedupeLinksText(text) {
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

// get_pc JSON -> item VideoAI (hoặc null nếu không hợp lệ)
export function buildVideoAIItem(json, productUrl) {
    const data = json && json.data;
    if (!data || !data.item) return null;
    const item = data.item;
    const shop = data.shop_detailed || {};
    const review = data.product_review || {};

    let rawImgs = (data.product_images && data.product_images.images) || item.images || [];
    if ((!rawImgs || rawImgs.length === 0) && item.image) rawImgs = [item.image];
    const images = (rawImgs || []).map(imgUrl).filter(Boolean);

    let priceRaw = num(item.price);
    if (priceRaw <= 0) priceRaw = num(item.price_min);
    const price = priceRaw / 100000;

    let origRaw = num(item.price_before_discount);
    if (origRaw <= 0) origRaw = num(item.price_max_before_discount);
    const originalPrice = origRaw / 100000;

    let stock = item.stock != null ? num(item.stock) : (item.normal_stock != null ? num(item.normal_stock) : 0);
    if (!stock && Array.isArray(item.models)) {
        stock = item.models.reduce((s, m) => s + num(m.stock || m.normal_stock), 0);
    }

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
