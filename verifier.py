"""
Проверка товаров и считывание внутренней пошлины прямо на карточке товара.

На страницах поиска маркетплейсы часто не показывают пошлину и полное описание,
а продавцы пишут общее название («Ноутбук Lenovo 16"») и указывают модель только в описании.
Поэтому бот параллельно по магазинам открывает карточки товаров из выдачи (все или до лимита) и берёт:
  • итоговую цену с карточки (как видит покупатель);
  • внутреннюю пошлину — ровно ту сумму, что указана на сайте (без формул);
  • полное описание и характеристики — для точной классификации и Excel.
Если карточка показала проверку «вы не робот» / блокировку — карточка пропускается (ничего не обходим).
"""
import asyncio
import inspect
import re
import time
from typing import Callable, Dict, List, Optional, Tuple

from chrome_engine import IsolatedChromeEngine
from parsers import ProductItem, _build_product, extract_duty_from_text

CARD_JS = r"""
(() => {
    const shop = %SHOP%;
    const isForeign = (el) => {
        if (!el || !el.closest) return false;
        return Boolean(el.closest(
            'article, [data-zone-name="productSnippet"], [data-auto="searchOrganic"], ' +
            '[data-zone-name*="carousel" i], [data-zone-name*="recommend" i], ' +
            '[data-zone-name*="similar" i], [data-zone-name*="analog" i], ' +
            '[data-zone-name*="moreOffers" i], [data-zone-name*="reviews" i]'
        ));
    };
    const q = (sels) => {
        for (const s of sels) {
            for (const el of document.querySelectorAll(s)) {
                if (isForeign(el)) continue;
                const txt = el.innerText ? el.innerText.trim() : '';
                if (txt) return txt;
            }
        }
        return '';
    };
    const priceSel = {
        'Ozon': ['[data-widget="webPrice"]', '[data-widget="webSale"]', '[data-widget="webOutOfStock"]'],
        'Wildberries': ['.price-block__wallet-price', '.price-block__final-price', 'ins[class*="price"]', '[class*="price-block__price"]'],
        'AliExpress': ['[class*="snow-price_SnowPrice__mainM"]', '[class*="Product_UniformBanner__uniformBannerBoxPrice"]', '[class*="price--current"]', '[class*="HazeProductPrice"]'],
        'Яндекс Маркет': ['[data-zone-name="mainOffer"] [data-auto="price-value"]', '[data-zone-name="offer"] [data-auto="price-value"]', '[data-auto="snippet-price-current"]', '[data-auto="main-price"]', '[data-auto="price-value"]', '[data-zone-name="price"]'],
        'Авито': ['[itemprop="price"]', '[data-marker="item-view/item-price"]'],
        'DNS': ['.product-buy__price', '.product-card-top__buy'],
    }[shop] || [];
    const buyboxSel = {
        'Ozon': ['[data-widget="webSale"]', '[data-widget="webPrice"]', '[data-widget="webDelivery"]'],
        'Wildberries': ['.price-block', '[class*="price-block"]', '[class*="delivery-info"]', '[class*="product-page__aside"]'],
        'AliExpress': ['[class*="Product_UniformBanner"]', '[class*="SnowPrice"]', '[class*="Delivery"]'],
        'Яндекс Маркет': ['[data-zone-name="mainOffer"]', '[data-zone-name="offer"]', '[data-auto="offer-info"]', '[data-zone-name="price"]', '[data-zone-name="delivery"]'],
        'Авито': ['[itemprop="price"]', '[data-marker="item-view/item-price"]', '[data-marker="delivery"]'],
        'DNS': ['.product-buy', '.product-card-top__buy'],
    }[shop] || [];
    const descSel = {
        'Ozon': ['[data-widget="webDescription"]', '[data-widget="webCharacteristics"]', '#section-description', '#section-characteristics'],
        'Wildberries': ['.product-params', '[class*="product-details"]', '[class*="description"]'],
        'AliExpress': ['[class*="SnowProductDescription"]', '[class*="Specifications"]', '#product-description', '[class*="description"]'],
        'Яндекс Маркет': ['[data-auto="product-full-specs"]', '[data-auto="product-description"]', '[data-zone-name="fullSpecs"]'],
        'Авито': ['[data-marker="item-view/item-description"]', '[itemprop="description"]', '[data-marker="item-view/item-params"]'],
        'DNS': ['.product-card-description-text', '.product-characteristics'],
    }[shop] || [];

    let priceText = q(priceSel);
    const metaPrice = document.querySelector('meta[itemprop="price"]');
    if (!priceText && metaPrice && metaPrice.content) priceText = metaPrice.content + ' ₽';

    const buyboxParts = [];
    for (const s of buyboxSel) {
        for (const el of document.querySelectorAll(s)) {
            if (isForeign(el)) continue;
            const t = (el.innerText || '').replace(/\n\s*₽/g, ' ₽').replace(/\s+/g, ' ').trim();
            if (t && !buyboxParts.includes(t)) buyboxParts.push(t);
        }
    }
    const buyboxText = buyboxParts.join(' | ').slice(0, 3000);

    let desc = '';
    if (shop === 'Авито') {
        const authorDesc = q(['[data-marker="item-view/item-description"]', '[itemprop="description"]']);
        const paramsEl = document.querySelector('[data-marker="item-view/item-params"]');
        const paramsTxt = paramsEl ? (paramsEl.innerText || '') : '';
        desc = (authorDesc ? authorDesc : '') + (paramsTxt ? (' | [ХАРАКТЕРИСТИКИ_АВИТО] ' + paramsTxt) : '');
    } else {
        for (const s of descSel) {
            document.querySelectorAll(s).forEach(el => {
                if (!isForeign(el)) desc += ' | ' + (el.innerText || '');
            });
        }
    }
    const rawBody = ((document.body ? document.body.innerText : '') || '').replace(/\n\s*₽/g, ' ₽');
    const stopIdx = rawBody.search(/похожие\s+товары|с\s+этим\s+товаром|вам\s+может\s+понравиться|рекомендуем/i);
    const topBody = (stopIdx > 800 ? rawBody.slice(0, stopIdx) : rawBody).slice(0, 8000);

    // Прицельно собираем блоки с пошлиной только из карточки товара (вне каруселей и чужих сниппетов)
    const dutyEls = [];
    const seenDuty = new Set();
    const dutyZoneNodes = buyboxSel.length
        ? Array.from(document.querySelectorAll(buyboxSel.join(', ') + ', span, div, p'))
        : Array.from(document.querySelectorAll('span, div, p'));
    for (const el of dutyZoneNodes) {
        if (dutyEls.length >= 12) break;
        if (isForeign(el)) continue;
        const t = (el.innerText || '').replace(/\n\s*₽/g, ' ₽').replace(/\s+/g, ' ').trim();
        if (!t || t.length > 200 || !/(?:пошлин|тамож)/i.test(t)) continue;
        if (/похожие\s+товары|рекомендуем|с\s+этим\s+товаром|рейтинг\s+товара|оценок:|купили|отзыв|вопрос/i.test(t)) continue;
        if (!seenDuty.has(t)) {
            seenDuty.add(t);
            dutyEls.push(t);
        }
    }

    const titleEl = document.querySelector('h1');
    return {
        pageUrl: location.href || '',
        title: titleEl ? titleEl.innerText.trim() : '',
        priceText: priceText.replace(/\n\s*₽/g, ' ₽').slice(0, 700),
        buyboxText: buyboxText,
        dutyText: dutyEls.slice(0, 12).join(' | '),
        desc: desc.slice(0, 6000),
        body: topBody.slice(0, 9000),
    };
})()
"""

BLOCK_MARKERS = (
    "доступ ограничен", "проблема с ip", "вы не робот", "подтвердите, что вы", "captcha", "капча",
    "проверка браузера", "access denied", "слишком много запросов",
)

_PRICE_SKIP_MARKERS = (
    "пошлин", "тамож", "/мес", "в мес", "мес.", "месяц",
    "рассроч", "кредит", "сплит", "долями",
    "балл", "бонус", "кешбэк", "кэшбэк", "выгод", "купон", "доставк",
)


def _prices_from_text(text: str) -> List[int]:
    if not text:
        return []
    prepared = re.sub(r"\n\s*₽", " ₽", text)
    prices: List[int] = []
    for line in re.split(r"[\n|]+", prepared):
        low = line.lower()
        if any(m in low for m in _PRICE_SKIP_MARKERS):
            continue
        if re.search(r"(?:/\s*мес|\bв\s+мес|\bмес\.?\b|\bмесяц|[×xх]\s*(?:2|4|6|10|12)\b)", low):
            continue
        clean = re.sub(r"[\s\u00a0\u2006\u2009\u202f]+", "", line)
        for m in re.findall(r"(\d{2,7})(?:[.,]\d+)?₽", clean):
            prices.append(int(m))
    return prices


def _pick_candidates(items: List[ProductItem], top_n: int = 0) -> List[ProductItem]:
    """
    Упорядочивает товары для обхода карточек:
    1) GREEN и YELLOW (сначала самые дешёвые),
    2) RED «Не соответствует модели» с общим/коротким заголовком (возможно, модель раскрыта в описании),
    3) остальные RED (если top_n <= 0, т.е. полная проверка всех товаров Excel).
    """
    pool = list(items) if top_n <= 0 else [
        it for it in items
        if it.status_code != "RED" or "Не соответствует модели" in (it.defects_summary or "")
    ]

    def _cand_key(it: ProductItem) -> tuple:
        p = it.base_price or it.price
        if it.status_code == "GREEN":
            return (0, p)
        if it.status_code == "YELLOW":
            return (1, p)
        t_low = (it.title or "").lower()
        has_alt_cpu = bool(re.search(r"\b\d{4}(?:hs|hx|h|u)\b", t_low))
        if "Не соответствует модели" in (it.defects_summary or "") and not has_alt_cpu:
            return (2, p)
        return (3, p)

    pool.sort(key=_cand_key)
    if top_n > 0:
        return pool[:top_n]
    by_q: Dict[str, List[ProductItem]] = {}
    for x in pool:
        by_q.setdefault((x.query or "").strip().lower(), []).append(x)
    if len(by_q) <= 1:
        return pool

    # Wave 1: гарантируем приоритетную проверку до 6 лучших GREEN/YELLOW карточек по каждому запросу
    wave1: List[ProductItem] = []
    seen_wave1 = set()
    buckets_gy = {
        qk: [x for x in lst if x.status_code in ("GREEN", "YELLOW")][:6]
        for qk, lst in by_q.items()
    }
    for idx_q in range(6):
        for lst in buckets_gy.values():
            if idx_q < len(lst):
                item_x = lst[idx_q]
                wave1.append(item_x)
                seen_wave1.add(id(item_x))

    # Wave 2: остальные карточки по round-robin запросов
    wave2: List[ProductItem] = []
    buckets_rest = [
        [x for x in lst if id(x) not in seen_wave1]
        for lst in by_q.values()
    ]
    idx_q = 0
    while True:
        added = False
        for b in buckets_rest:
            if idx_q < len(b):
                wave2.append(b[idx_q])
                added = True
        if not added:
            break
        idx_q += 1
    return wave1 + wave2


def _apply_card_data(it: ProductItem, data: dict) -> Tuple[ProductItem, bool]:
    """Пересобирает ProductItem по данным с открытой карточки товара."""
    list_price = it.base_price or it.price
    card_prices = [p for p in _prices_from_text(data.get("priceText", "")) if p >= 10]
    card_price = list_price
    if card_prices:
        cand = min(card_prices)
        # Защита от чужих цифр на странице при допуске реальных региональных и акционных колебаний
        if list_price and 0.35 * list_price <= cand <= 2.2 * list_price:
            card_price = cand

    buybox = (data.get("buyboxText") or "").strip()
    duty_zone = f"{data.get('dutyText', '')} | {data.get('priceText', '')} | {buybox}"
    if not buybox and it.shop != "Яндекс Маркет":
        duty_zone += f" | {(data.get('body') or '')[:2200]}"
    site_duty = extract_duty_from_text(duty_zone, card_price)
    full_desc = f"{it.description} | {data.get('desc', '')}".strip(" |")

    new_it = _build_product(
        shop=it.shop,
        query=it.query,
        title=it.title,
        price=card_price,
        seller_name=it.seller_name,
        seller_rating=it.seller_rating,
        reviews_count=it.reviews_count,
        location_delivery=it.location_delivery,
        description=full_desc,
        url=it.url,
        site_duty=site_duty,
        snippet_fallback=False,
    )
    changed = (
        new_it.base_price != it.base_price
        or new_it.customs_duty != it.customs_duty
        or new_it.status_code != it.status_code
    )
    new_it.duty_note = f"{new_it.duty_note} ✔ проверено на карточке"
    return new_it, changed


async def verify_top_cards(
    engine: IsolatedChromeEngine,
    items: List[ProductItem],
    top_n: int = 0,
    avito_top_n: int = 0,
    shops: tuple = ("Ozon", "Wildberries", "AliExpress", "Яндекс Маркет", "Авито", "DNS"),
    max_total_sec: float = 90.0,
    progress_cb: Optional[Callable[[Dict[str, int]], object]] = None,
    ymarket_lr: int = 213,
) -> Dict[str, int]:
    """
    Параллельно по магазинам (по 1 выделенной вкладке Chrome на магазин) открывает карточки
    товаров из выдачи и пересобирает ProductItem с точной ценой, внутренней пошлиной и описанием.
    При top_n == 0 проверяет все товары списка items (с приоритетом топ-6 GREEN/YELLOW каждого запроса).
    """
    stats = {"checked": 0, "total": len(items), "updated": 0, "blocked": 0, "duties_found": 0}
    if not items:
        return stats

    idx_by_obj = {id(it): i for i, it in enumerate(items)}
    seen_card_cache: Dict[Tuple[str, str], dict] = {}
    deadline = time.monotonic() + max(20.0, float(max_total_sec or 90.0))
    last_cb_ts = time.monotonic()
    cb_lock = asyncio.Lock()

    try:
        await engine.ensure_started()
    except Exception:
        return stats

    async def _maybe_notify(force: bool = False) -> None:
        nonlocal last_cb_ts
        if not progress_cb:
            return
        async with cb_lock:
            now = time.monotonic()
            if not force and (now - last_cb_ts < 7.5):
                return
            last_cb_ts = now
            try:
                res = progress_cb(dict(stats))
                if inspect.isawaitable(res):
                    await res
            except Exception:
                pass

    async def _shop_worker(shop: str) -> None:
        shop_items = [it for it in items if it.shop == shop]
        if not shop_items:
            return
        n = avito_top_n if shop == "Авито" else top_n
        candidates = _pick_candidates(shop_items, n)
        if not candidates:
            return

        tab_id, tab_ws = "", ""
        try:
            tab_id, tab_ws = await engine.create_tab_ws()
        except Exception:
            tab_id, tab_ws = "", ""

        blocked_streak = 0
        err_streak = 0
        js = CARD_JS.replace("%SHOP%", repr(shop))
        try:
            for it in candidates:
                if blocked_streak >= 3 or err_streak >= 3 or time.monotonic() >= deadline:
                    break
                orig_idx = idx_by_obj.get(id(it))
                if orig_idx is None:
                    continue

                url_key = (shop, (it.url or "").split("?")[0])
                data = seen_card_cache.get(url_key)
                from_cache = data is not None
                if not from_cache:
                    fetch_url = it.url
                    if shop == "Яндекс Маркет" and ymarket_lr and "lr=" not in fetch_url:
                        sep = "&" if "?" in fetch_url else "?"
                        fetch_url = f"{fetch_url}{sep}lr={int(ymarket_lr)}"
                    try:
                        data = await engine.fetch_page_js(
                            fetch_url,
                            js,
                            wait_sec=1.0,
                            scroll=False,
                            ready_selector="h1",
                            timeout_sec=10.0,
                            retries=1,
                            tab_ws=tab_ws or None,
                            allow_restart=False,
                        )
                    except Exception:
                        data = None
                    if not isinstance(data, dict) or not (data.get("body") or data.get("priceText") or data.get("title")):
                        err_streak += 1
                        continue
                    err_streak = 0
                    page_url_low = (data.get("pageUrl") or "").lower()
                    body_low = (data.get("body") or "").lower()[:3000]
                    is_punish = any(k in page_url_low for k in ("/punish", "x5secdata", "captcha", "antibot"))
                    if is_punish or (any(m in body_low for m in BLOCK_MARKERS) and not data.get("priceText")):
                        stats["blocked"] += 1
                        blocked_streak += 1
                        continue
                    blocked_streak = 0
                    seen_card_cache[url_key] = data

                new_it, changed = _apply_card_data(it, data)
                stats["checked"] += 1
                if new_it.customs_duty > 0:
                    stats["duties_found"] += 1
                if changed:
                    stats["updated"] += 1
                items[orig_idx] = new_it
                await _maybe_notify(force=False)
                if not from_cache:
                    await asyncio.sleep(0.15)
        finally:
            if tab_id:
                await engine.close_tab_ws(tab_id)

    await asyncio.gather(*(_shop_worker(sh) for sh in shops if any(it.shop == sh for it in items)))
    await _maybe_notify(force=True)
    return stats


async def verify_all_cards(
    engine: IsolatedChromeEngine,
    items: List[ProductItem],
    shops: tuple = ("Ozon", "Wildberries", "AliExpress", "Яндекс Маркет", "Авито", "DNS"),
    max_total_sec: float = 90.0,
    progress_cb: Optional[Callable[[Dict[str, int]], object]] = None,
    ymarket_lr: int = 213,
) -> Dict[str, int]:
    """Проверяет карточки всех товаров из выдачи (по каждой площадке параллельно) и забирает внутреннюю пошлину."""
    return await verify_top_cards(
        engine=engine,
        items=items,
        top_n=0,
        avito_top_n=0,
        shops=shops,
        max_total_sec=max_total_sec,
        progress_cb=progress_cb,
        ymarket_lr=ymarket_lr,
    )
