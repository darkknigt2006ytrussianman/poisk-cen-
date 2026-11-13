import asyncio
import json
import math
import re
import time
import urllib.parse
from dataclasses import dataclass
from datetime import datetime
from typing import List

import httpx

from analyzer import analyze_item, dehomoglyph_lat, dehomoglyph_ru
from chrome_engine import IsolatedChromeEngine


@dataclass
class ProductItem:
    shop: str                # Ozon / AliExpress / Яндекс Маркет / Авито / DNS / Wildberries
    query: str               # Поисковый запрос
    title: str               # Название товара
    in_stock: str            # "✅ В продаже" или "❌ Нет в продаже"
    price: int               # ИТОГОВАЯ цена в рублях (Цена витрины + Таможенная пошлина)
    base_price: int          # Цена витрины (без пошлины)
    customs_duty: int        # Таможенная пошлина в рублях
    duty_note: str           # Пояснение по пошлине
    seller_name: str         # Имя / название продавца
    seller_rating: float     # Рейтинг продавца или товара
    reviews_count: int       # Кол-во отзывов
    location_delivery: str   # Город / наличие доставки
    description: str         # Описание состояния / характеристики
    url: str                 # Прямая ссылка на товар
    status_code: str         # GREEN / YELLOW / SKIP
    status_label: str        # Читаемый вердикт состояния
    defects_summary: str     # Найденные нюансы (царапины / замена / брак)
    seller_type: str         # Оценка продавца
    score: int               # Балл от 0 до 100
    updated_at: str          # Время обнаружения
    price_change: str = "Новый"  # Статус изменения цены


def detect_stock_status(shop: str, title: str, description: str, url: str = "") -> str:
    """
    Определяет, есть ли товар в продаже (в наличии) или он уже продан / недоступен.
    Любой проданный или недоступный товар получает статус ❌ и полностью исключается из Excel.
    """
    full_text = f"{title} {description}".lower()
    out_of_stock_markers = [
        "нет в продаже",
        "нет в наличии",
        "товар закончился",
        "товара нет в наличии",
        "товар уже продан",
        "товар продан",
        "уже продан",
        "распродано",
        "разобрали",
        "отсутствует в продаже",
        "уведомить о поступлении",
        "узнать о поступлении",
        "товар недоступен",
        "снято с публикации",
        "объявление снято",
        "объявление закрыто",
        "товар зарезервирован",
        "в резерве",
        "зарезервировано",
        "не доставляется",
        "продажи прекращены",
        "данного товара прекращена",
        "продажа прекращена",
    ]
    if any(m in full_text for m in out_of_stock_markers):
        return "❌ Нет в продаже (Распродан)"

    return "✅ В продаже (В наличии)"


_DUTY_STOP_CONTEXT = (
    "беспошлин",
    "порог",
    "превышени",
    "при заказе от",
    "при заказе до",
    "при покупке от",
    "при покупке до",
    "стоимостью от",
    "стоимостью свыше",
    "свыше 200",
    "более 200",
    "без пошлин",
    "пошлина не треб",
    "пошлина не нужна",
    "пошлины нет",
    "пошлина не взимается",
    "пошлина включена",
    "уже включена",
    "включена в цену",
    "входит в стоимость",
)

DUTY_REGEX_PATTERNS = (
    r"(?:(?:доплат\w*\s+за\s+)?таможен\w*\s+(?:пошлин[аыуе]?|плат[её]ж\w*|сбор\w*)|(?<![а-яё])пошлин[аыуе]?)\s*(?:составит|около|примерно|ориентировочно|при\s+получении|на\s+таможне|оплачивается\s+отдельно|за\s+заказ|будет)?\s*[·~≈:+=-]?\s*(\d[\d\s\u00a0\u2006\u2009\u202f]{0,7})\s*(?:₽|руб\.?)",
    r"[+~≈]?\s*[~≈]?\s*(\d[\d\s\u00a0\u2006\u2009\u202f]{0,7})\s*(?:₽|руб\.?)\s*[-—–•·\s]*(?:(?:доплат\w*\s+за\s+)?таможен\w*\s+(?:пошлин\w*|плат[её]ж\w*|сбор\w*)|(?<![а-яё])пошлин\w*)",
)


def is_realistic_duty(val: int, base_price: int) -> bool:
    """Проверяет, что считанная сумма похожа на реальную пошлину сайта, а не на цену товара или порог 200 €."""
    if not (1 <= val <= 120000):
        return False
    if base_price:
        if val == base_price or val >= base_price * 0.45:
            return False
        if base_price < 90000 and val >= 19000:
            return False
    return True


def extract_duty_from_text(text: str, base_price: int) -> int:
    """Извлекает реальную сумму таможенной пошлины/доплаты из текста, отсекая пороги и чужие сборы."""
    if not text:
        return 0
    raw_segs = [re.sub(r"\s+", " ", seg.lower()).strip() for seg in re.split(r"[\n|]+", str(text)) if seg and seg.strip()]
    candidates = list(raw_segs) + [f"{raw_segs[i]} {raw_segs[i + 1]}" for i in range(len(raw_segs) - 1)]
    for low in candidates:
        if not low or ("пошлин" not in low and "тамож" not in low):
            continue
        if any(s in low for s in _DUTY_STOP_CONTEXT):
            continue
        for pat in DUTY_REGEX_PATTERNS:
            for m in re.finditer(pat, low):
                win_start = max(0, m.start() - 35)
                win_end = min(len(low), m.end() + 25)
                win = low[win_start:win_end]
                if any(s in win for s in _DUTY_STOP_CONTEXT):
                    continue
                val = int(re.sub(r"\D", "", m.group(1)) or 0)
                if is_realistic_duty(val, base_price):
                    return val
    return 0


def calculate_customs_duty(
    shop: str,
    base_price: int,
    title: str,
    description: str,
    duty_limit_rub: int = 20500,
    site_duty: int = 0,
    snippet_fallback: bool = True,
) -> tuple[int, str]:
    """
    Пошлина берётся ТОЛЬКО с сайта — никаких формул.
    Ozon, Wildberries, AliExpress и Яндекс Маркет сами показывают пошлину/таможенный платёж на карточке
    (суммы бывают любые: 4, 8, 487, 6 017 ₽ и т.д.), поэтому:
    1) сумма из открытой карточки товара (site_duty) — главный источник;
    2) при snippet_fallback=True — явная сумма «пошлина / таможенный платёж … ₽» в тексте карточки/сниппета;
    3) иначе — 0 ₽ (на сайте пошлина не указана).
    """
    if is_realistic_duty(int(site_duty or 0), base_price):
        sd = int(site_duty)
        return sd, f"+{sd:,} ₽ пошлина (указана на карточке сайта)".replace(",", " ")

    full_text = re.sub(r"\s*\|\s*", " ", f"{title} {description}").lower()
    full_text = re.sub(r"\s+", " ", full_text).strip()

    if any(p in full_text for p in ("без пошлин", "пошлина не треб", "пошлина не нужна", "пошлины нет", "пошлина включена")):
        return 0, "Без пошлины (так указано на сайте)"

    if snippet_fallback:
        explicit_duty = extract_duty_from_text(f"{title} | {description}", base_price)
        if explicit_duty > 0:
            return explicit_duty, f"+{explicit_duty:,} ₽ пошлина (указана на сайте)".replace(",", " ")

    cross_border_markers = ("из-за рубежа", "из за рубежа", "ozon global", "доставка из китая", "из китая", "cainiao")
    if shop in ("Ozon", "Wildberries", "AliExpress", "Яндекс Маркет") and any(m in full_text for m in cross_border_markers):
        return 0, "Из-за рубежа: пошлина на сайте не указана (0 ₽)"
    return 0, "Без пошлины (0 ₽)"


def extract_description_specs(query: str, title: str, description: str) -> str:
    """
    Извлекает ключевые характеристики из описания товара (процессор, ОЗУ, SSD, видеокарту, экран и др.),
    особенно когда в названии указано общее наименование, а точная конфигурация раскрыта в описании.
    """
    t_low = (title or "").lower()
    d_clean = re.sub(r"\s+", " ", (description or "")).strip()
    if not d_clean:
        return ""

    found_specs = []

    # Поиск модели процессора в описании
    cpu_patterns = [
        r"\b(ryzen\s*[3579]\s*(?:pro\s*)?\d{4}[a-z]{1,3}\d?)\b",
        r"\b(ryzen\s*ai\s*9\s*(?:hx\s*)?\d{3})\b",
        r"\b(core\s*i[3579][-\s]*\d{4,5}[a-z]{0,3})\b",
        r"\b(core\s*ultra\s*[579][-\s]*\d{3}[a-z]{0,2})\b",
        r"\b(apple\s*m[1234]\s*(?:pro|max|ultra)?)\b",
        r"\b(r[3579][-\s]*\d{4}[a-z]{1,2})\b",
        r"\b(i[3579][-\s]*\d{4,5}[a-z]{1,2})\b",
        r"\b(snapdragon\s*(?:8\s*gen\s*\d|\d+)[^\s,|;]{0,12})\b",
        r"\b(dimensity\s*\d{4}[^\s,|;]{0,8})\b",
        r"\b(helio\s*g\d{2,3}[^\s,|;]{0,6})\b",
        r"\b(tensor\s*g[12345])\b",
    ]
    for cp in cpu_patterns:
        m = re.search(cp, d_clean, re.IGNORECASE)
        if m:
            val = m.group(1).strip()
            if val.lower() not in t_low:
                found_specs.append(f"ЦП (по описанию): {val}")
            else:
                found_specs.append(f"ЦП: {val}")
            break

    # Поиск ОЗУ / Накопителя
    ram_ssd = re.search(r"\b(\d{1,2}\s*(?:гб|gb|g)\s*[/+]\s*\d{3,4}\s*(?:гб|gb|тб|tb|g))\b", d_clean, re.IGNORECASE)
    if ram_ssd:
        found_specs.append(f"Память: {ram_ssd.group(1)}")
    else:
        ram_m = re.search(r"(?:озу|оперативн\w*\s*пам\w*|ram|ddr[45x]*)\s*[:=-]?\s*(\d{1,2}\s*(?:гб|gb))", d_clean, re.IGNORECASE)
        if ram_m:
            found_specs.append(f"ОЗУ: {ram_m.group(1)}")
        ssd_m = re.search(r"(?:ssd|ссд|накопител\w*)\s*[:=-]?\s*(\d{3,4}\s*(?:гб|gb)|[124]\s*(?:тб|tb))", d_clean, re.IGNORECASE)
        if ssd_m:
            found_specs.append(f"SSD: {ssd_m.group(1)}")

    # Видеокарта
    gpu_m = re.search(r"\b((?:rtx|gtx|rx|radeon)\s*\d{3,4}\s*(?:ti|super|xt|m)?)\b", d_clean, re.IGNORECASE)
    if gpu_m:
        found_specs.append(f"Видео: {gpu_m.group(1)}")

    if found_specs:
        return " | ".join(found_specs)
    return ""


_ADJ_ENDINGS = ("ый", "ий", "ой", "ая", "яя", "ое", "ее", "ые", "ие", "ых", "их", "ую", "юю")
_HEAD_STOP = {
    "купить", "новый", "новая", "новое", "новые", "цена", "оригинал", "дешево", "дёшево", "недорого",
    "доставка", "россия", "бу", "б/у", "товар", "штука", "шт", "комплект", "набор",
}


def _ru_stem(word: str) -> str:
    """Грубая основа слова: «табуретка» → «табурет», «стулья» → «стул», «корм» → «корм»."""
    w = word.lower().strip()
    return w[: max(4, len(w) - 2)] if len(w) > 4 else w


def _query_head_noun(q_low: str) -> str:
    """Главное существительное запроса (до «для/под/на/к/с»): «корм для кошек» → «корм», «детский стул» → «стул»."""
    head_part = re.split(r"\s(?:для|под|на|к|с|из|от)\s", f" {q_low} ", maxsplit=1)[0]
    for w in re.findall(r"[а-яё]{3,}", head_part):
        if w in _HEAD_STOP or w.endswith(_ADJ_ENDINGS):
            continue
        return w
    return ""


def _normalize_cpu_text(text: str) -> str:
    """Склеивает разорванные пробелами индексы моделей и убирает кириллические гомоглифы ('Ryzеn 5 7640НS' -> 'ryzen 5 7640hs')."""
    s = dehomoglyph_lat(text or "")
    s = re.sub(r"\bryze\b", "ryzen", s)
    # Склеиваем 1 цифру + пробел + 3 цифры: '7 640' -> '7640', '8 845' -> '8845', '7 645' -> '7645'
    s = re.sub(r"\b([3456789])\s+(\d{3})(?=\s*(?:hs|hx|h|u|g|f|x3d|x|\b))", r"\1\2", s)
    # Склеиваем 4 цифры + пробел + суффикс: '7640 hs' -> '7640hs'
    s = re.sub(r"\b(\d{4})\s+(hs|hx|h|u|x3d)\b", r"\1\2", s)
    return s


def clean_product_url(url: str) -> str:
    """Очищает длинные трекинг-параметры из ссылок маркетплейсов, сохраняя путь, SKU и ID карточки."""
    raw = (url or "").strip()
    if not raw or not raw.startswith(("http://", "https://")):
        return raw
    try:
        parsed = urllib.parse.urlsplit(raw)
        host = parsed.netloc.lower()
        q_pairs = urllib.parse.parse_qsl(parsed.query, keep_blank_values=False)

        if "market.yandex." in host:
            keep = [(k, v) for k, v in q_pairs if k.lower() in ("sku", "modelid")]
            new_q = urllib.parse.urlencode(keep, doseq=True)
            return urllib.parse.urlunsplit((parsed.scheme, parsed.netloc, parsed.path, new_q, ""))

        if "ozon." in host:
            return urllib.parse.urlunsplit((parsed.scheme, parsed.netloc, parsed.path, "", ""))

        if "aliexpress." in host:
            keep = [(k, v) for k, v in q_pairs if k.lower() in ("sku_id", "skuid")]
            new_q = urllib.parse.urlencode(keep, doseq=True)
            return urllib.parse.urlunsplit((parsed.scheme, parsed.netloc, parsed.path, new_q, ""))

        if "wildberries." in host or "dns-shop." in host:
            return urllib.parse.urlunsplit((parsed.scheme, parsed.netloc, parsed.path, "", ""))

        if "avito." in host:
            drop_prefixes = ("utm_", "context", "src", "iid", "from", "s")
            keep = [(k, v) for k, v in q_pairs if not k.lower().startswith(drop_prefixes)]
            new_q = urllib.parse.urlencode(keep, doseq=True)
            return urllib.parse.urlunsplit((parsed.scheme, parsed.netloc, parsed.path, new_q, ""))

        drop_prefixes = ("utm_", "yclid", "gclid", "fbclid", "cpc", "wprid", "spm")
        keep = [(k, v) for k, v in q_pairs if not k.lower().startswith(drop_prefixes)]
        new_q = urllib.parse.urlencode(keep, doseq=True)
        return urllib.parse.urlunsplit((parsed.scheme, parsed.netloc, parsed.path, new_q, ""))
    except Exception:
        return raw


def _is_accessory_query(q_low: str) -> bool:
    """Определяет, ищет ли сам пользователь запчасть, чехол, расходник или аксессуар."""
    if not q_low:
        return False
    acc_markers = (
        "чехол", "бампер", "накладк", "обложк", "кейс", "стекло", "пленк", "гидрогел",
        "заряд", "блок питания", "кабель", "адаптер",
        "дисплей", "экран", "матриц", "аккумулятор", "батаре", "клавиатур", "тачпад", "шлейф", "кулер",
        "подшипник", "ролик", "колесик", "колёсик", "газлифт", "крестовин", "подлокотник", "механизм",
        "топ-ган", "топган", "пиастра", "заглушк", "болт", "винт", "гайк", "шайб", "фурнитур", "запчаст",
        "ремкомплект", "ремнабор", "обивк", "перетяжк", "накидк", "подушк", "подголовник",
        "основани", "поднос", "каркас", "мультиблок", "насадк", "фильтр", "картридж", "стержн", "стержен",
        "пылесборник", "шланг", "колб", "чернил", "тонер", "сменн", "запасн", "расходник",
    )
    return any(w in q_low for w in acc_markers)


def _is_aliexpress_titlecase_spam(title: str) -> bool:
    """Определяет характерный машинный Title Case AliExpress («Основание Кресла Босса Поднос Для Обеда...»)."""
    words = re.findall(r"[А-ЯЁа-яё]{3,}", title or "")
    if len(words) < 5:
        return False
    cap_cnt = sum(1 for w in words if w[0].isupper())
    return cap_cnt >= max(4, int(len(words) * 0.78))


def _effective_min_price(query: str, min_price: int) -> int:
    """
    Защита от мусорной выдачи запчастей по 50–300 ₽ (подшипники, болты, заглушки, ролики)
    при сортировке по возрастанию цены: если пользователь ищет готовый товар (кресло, стол, ноутбук и т.д.),
    а нижний порог цены остался дефолтным (≤ 300 ₽), поднимаем поисковый порог до разумного минимума категории.
    """
    q_low = (query or "").lower()
    if _is_accessory_query(q_low):
        return max(50, int(min_price or 50))
    if min_price and min_price > 500:
        return int(min_price)
    if any(w in q_low for w in ("ноутбук", "macbook", "макбук", "ryzen", "core i", "thinkpad", "ideapad", "legion", "vivobook", "matebook", "magicbook")):
        return max(int(min_price or 0), 15000)
    if any(w in q_low for w in ("телефон", "смартфон", "iphone", "айфон", "ipad", "планшет", "телевизор", "монитор", "видеокарт", "холодильник", "стиральн")):
        return max(int(min_price or 0), 3000)
    if any(w in q_low for w in ("кресл", "диван", "кроват", "шкаф", "матрас", "стол", "комод")):
        return max(int(min_price or 0), 2500)
    if any(w in q_low for w in ("табурет", "стул", "тумб", "полк", "стеллаж", "пуф", "банкетк")):
        return max(int(min_price or 0), 500)
    if any(w in q_low for w in ("шуруповерт", "шуруповёрт", "дрель", "перфоратор", "пылесос", "кофевар", "кофемашин", "микроволнов", "мультиварк", "чайник", "утюг", "фен", "блендер", "обогревател", "принтер")):
        return max(int(min_price or 0), 700)
    return max(int(min_price or 0), 200)


UNIVERSAL_PARTS_BLACKLIST = (
    # Мебель / кресла / фурнитура (строгие слова-запчасти и связки «для/на/к», не задевающие «кресло с подлокотниками и газлифтом»)
    "подшипник", "ролики для", "ролик для", "колесики для", "колёсики для", "колеса для", "колёса для", "колесо для",
    "газлифт для", "газлифт к ", "газ-лифт для", "газпатрон для", "крестовина для", "крестовина к ",
    "пятилучье для", "пятилучье к ", "подлокотник для", "подлокотники для", "механизм качания для",
    "топ-ган", "топган", "пиастра", "перманент-контакт", "мультиблок", "заглушк", "втулк",
    "основание кресл", "основание стул", "основание для", "поднос для", "каркас кресл", "каркас для",
    "шайба", "шайбы", "гайка", "гайки", "болт ", "болты ", "винт ", "винты ", "саморез", "шуруп",
    "крепеж", "крепёж", "фурнитур", "ремкомплект", "ремнабор", "запчаст", "комплектующ", "ножки для", "ножка для",
    "опора для", "опоры для", "спинка для", "сиденье для", "перетяжк", "чехол на ", "чехол для", "чехлы на ", "чехлы для",
    "накидка на ", "накидка для", "подушка для", "подголовник для", "валик под", "коврик под", "коврик для", "коробка от",
    "инструкция", "смазка для", "крепежный", "крепёжный", "накладка на ", "накладка для", "накладки на ", "накладки для", "шарнир для",
    # Бытовые товары / техника / хозтовары / канцелярия (расходники и детали)
    "насадка для", "насадки для", "насадка на ", "насадки на ", "фильтр для", "фильтры для", "пылесборник",
    "картридж для", "картриджи для", "колба для", "шланг для", "ручка для", "ручка к ",
    "держатель для", "держатель к ", "кронштейн для", "кнопка для", "крышка для", "крышка на ",
    "лоток для", "поддон для", "ремень для", "щетка для", "щётка для", "губка для",
    "лезвие для", "лезвия для", "нож для", "ножи для", "чаша для", "резервуар для",
    "уплотнитель для", "прокладка для", "шестерня для", "тэн для",
    "стержень для", "стержни для", "стержень к ", "стержни к ", "ампула для", "заправка для", "чернила для", "тонер для",
    "сменный блок", "сменная кассета", "сменные лезвия", "сменный стержень", "сменные стержни", "сменная насадка",
    "сменный фильтр", "запасная часть", "запасные части", "запасной ", "запасная ", "запасные ",
    "расходник", "для кукол", "кукольное", "миниатюр", "игрушечн",
)


def _matches_target_model(query: str, title: str, description: str, url: str = "") -> bool:
    """
    Строгая категорийная и модельная фильтрация карточек товаров по НАЗВАНИЮ и ОПИСАНИЮ:
    - Проверяет как заголовок (title), так и полное описание/характеристики (description).
    - Если в названии указано общее наименование (например «Ноутбук Lenovo 16"»), а конкретный
      процессор или характеристики (например «Ryzen 5 7640HS») указаны в описании — товар успешно проходит!
    - Отсекает чужие категории (чехлы/запчасти при поиске устройства) и конфликтующие модели в заголовке.
    """
    q_low = _normalize_cpu_text(re.sub(r"\s+", " ", (query or "").strip()))
    t_low = _normalize_cpu_text(re.sub(r"\s+", " ", (title or "").strip()))
    d_low = _normalize_cpu_text(re.sub(r"\s+", " ", (description or "").strip()))
    d_ru = dehomoglyph_ru(description or "")
    # Отделяем системный справочник площадки Авито, если продавец честно предупредил, что нужной модели нет в списке
    author_d_low = re.sub(r"\[характеристики_авито\].*$", "", d_low, flags=re.S)
    has_no_model_note = any(
        p in d_ru
        for p in ("в авито нет нужной модели", "нет нужной модели на авито", "нужной модели в списке нет", "авито не знает такую модель")
    )
    if has_no_model_note:
        # Дополнительно отрезаем блок характеристик Авито («о ноутбуке ... характеристики ... описание»), если пришёл слитно
        cleaned_author = re.sub(r"\bо\s+ноутбуке\b.*?\bописание\b", " ", author_d_low, flags=re.S)
        author_d_low = cleaned_author if cleaned_author.strip() else author_d_low
    cpu_desc_scope = author_d_low if has_no_model_note else d_low
    haystack = f"{t_low} {d_low} {dehomoglyph_ru(f'{title} {description}')} {url.lower()}"

    if not t_low or len(t_low) < 3:
        return False

    # 1. Отсекаем заведомо другие младшие процессоры, если их нет в запросе
    wrong_cpus = ["5500u", "5600u", "5625u", "5700u", "4500u", "3500u", "7520u", "7530u", "7430u"]
    wrong_scope = f"{t_low} {cpu_desc_scope}"
    for fake in wrong_cpus:
        if fake not in q_low and fake in wrong_scope:
            return False

    # Если в запросе задан 4-значный индекс процессора (например 7640hs, 7645hs, 8845hs):
    # Сравниваем 4-значное ядро номера (7640H == 7640HS, но 7645 != 7640!)
    q_cpu_digits = re.findall(r"\b(\d{4})(?:hs|h|hx|u|g|f|x3d|x)?\b", q_low)
    if q_cpu_digits:
        title_cpu_digits = re.findall(r"\b(\d{4})(?:hs|h|hx|u|g|f|x3d|x)\b", t_low)
        desc_cpu_digits = re.findall(r"\b(\d{4})(?:hs|h|hx|u|g|f|x3d|x)\b", cpu_desc_scope)
        # Также ловим 4-значные номера после слов ryzen 5/7/9 без буквенного суффикса (например «ryzen 5 7645»)
        desc_cpu_digits += re.findall(r"ryzen\s*[3579]\s*(?:pro\s*)?(\d{4})\b", cpu_desc_scope)
        title_cpu_digits += re.findall(r"ryzen\s*[3579]\s*(?:pro\s*)?(\d{4})\b", t_low)

        # Если в ЗАГОЛОВКЕ стоит другой 4-значный номер ЦП (например 7645 или 7535 вместо 7640) — отбраковываем
        if title_cpu_digits and not any(td in q_cpu_digits for td in title_cpu_digits):
            return False
        # Если в ОПИСАНИИ указан другой процессор, противоречащий запросу (например в заголовке 7640, а в описании 7645 или 7535) — отбраковываем!
        conflict_in_desc = [dd for dd in desc_cpu_digits if dd not in q_cpu_digits and dd not in ("1920", "1200", "1080", "1440", "1600", "2560", "2023", "2024", "2025", "2026", "3200", "4800", "5200", "5600", "6400")]
        if conflict_in_desc and not any(dd in q_cpu_digits for dd in desc_cpu_digits):
            return False
        if conflict_in_desc and any(dd in q_cpu_digits for dd in desc_cpu_digits):
            # В описании перечислено сразу несколько разных процессоров (список перекупщика: 7535 / 7640 / 7645)
            return False

    # Определяем, ищет ли пользователь аксессуар/запчасть или само устройство
    want_case = any(w in q_low for w in ("чехол", "бампер", "накладк", "обложк", "кейс"))
    want_glass = any(w in q_low for w in ("стекло", "пленк", "гидрогел"))
    want_charger = any(w in q_low for w in ("заряд", "блок питания", "кабель", "адаптер"))
    want_part = any(
        w in q_low
        for w in (
            "дисплей", "экран", "матриц", "аккумулятор", "батаре", "клавиатур", "тачпад", "шлейф", "кулер",
            "подшипник", "ролик", "колесик", "колёсик", "газлифт", "крестовин", "подлокотник", "механизм",
            "топ-ган", "топган", "пиастра", "заглушк", "болт", "винт", "гайк", "шайб", "фурнитур", "запчаст",
            "ремкомплект", "обивк", "перетяжк", "накидк", "подушк", "подголовник",
        )
    )
    want_accessory = want_case or want_glass or want_charger or want_part or _is_accessory_query(q_low)

    # Глобальный отсев запчастей, комплектующих, расходников, роликов, газлифтов и чехлов при поиске готового товара
    if not want_accessory:
        for part_word in UNIVERSAL_PARTS_BLACKLIST:
            pw_clean = part_word.strip()
            if pw_clean not in q_low and part_word in t_low:
                return False
        # Если в любом месте заголовка идёт запчасть/расходник + предлог цели («для/под/к/от/на») или родительный падеж цели
        part_heads = (
            r"(?:основани[ея]|поднос\w*|каркас\w*|мультиблок\w*|подшипник\w*|ролик\w*|колес\w*|колёс\w*|"
            r"газлифт\w*|крестовин\w*|подлокотник\w*|механизм\w*|заглушк\w*|чех[ое]л\w*|накидк\w*|накладк\w*|"
            r"подушк\w*|болт\w*|винт\w*|гайк\w*|шайб\w*|опор\w*|ножк\w*|спинк\w*|сидень\w*|обивк\w*|ремкомплект\w*|"
            r"ремнабор\w*|запчаст\w*|фурнитур\w*|насадк\w*|фильтр\w*|картридж\w*|стерж\w*|пылесборник\w*|"
            r"шланг\w*|колб\w*|держател\w*|кронштейн\w*|крышк\w*|лоток|лотки|поддон\w*|кнопк\w*|уплотнител\w*|прокладк\w*|шестер\w*|тэн\w*)"
        )
        m_head = re.match(rf"^(?:комплект\s+|набор\s+|универсальн\w*\s+|сменн\w*\s+|запасн\w*\s+)?({part_heads})\b", t_low)
        if m_head and m_head.group(1)[:4] not in q_low:
            return False
        if re.search(rf"\b{part_heads}\s+(?:для|под|к|от|на)\s+", t_low):
            return False
        if re.search(rf"\b{part_heads}\s+(?:кресла|стула|стола|дивана|кровати|шкафа|пылесоса|чайника|кофеварки|кофемашины|блендера|принтера|утюга|микроволновки|ручки|ноутбука|телефона|планшета)\b", t_low):
            return False

    # Специальная проверка для категории КРЕСЛО (компьютерное / офисное / игровое кресло)
    if "кресл" in q_low and not want_accessory:
        # В названии обязательно должно быть само слово «кресло» (или «кресла» не в родительном падеже после запчасти) или «компьютерный стул»
        if not any(w in t_low for w in ("кресло", "кресла", "компьютерный стул", "стул компьютерный", "стул офисный", "офисный стул", "игровой стул")):
            return False
        # Отсекаем родительный падеж («... для компьютерного кресла», «основание кресла ...») даже если продавец дописал в конце «Игровое Кресло»
        if re.search(r"(?:для|от|к|у|из|компьютерного|офисного|игрового|рабочего|руководителей|руководителя)\s+кресла\b", t_low):
            return False
        if ("butterfly" in t_low or "бабочк" in t_low) and "butterfly" not in q_low and "бабочк" not in q_low:
            return False
        if "aliexpress." in url.lower() and _is_aliexpress_titlecase_spam(title) and any(
            w in t_low for w in ("босса", "поднос", "основан", "многофункцион", "butterfly", "среднего класса", "руководителей")
        ):
            return False
        # Отсекаем несовместимые виды кресел, если ищут компьютерное / офисное / игровое кресло
        if any(w in q_low for w in ("компьютер", "офис", "игров", "геймер", "рабоч", "руководит", "оператор")):
            wrong_chairs = (
                "автокресло", "детское автокресло", "бустер", "кресло-мешок", "кресло мешок", "кресло-груша", "кресло груша",
                "кресло-мяч", "бинбэг", "кресло-качалка", "кресло качалка", "глайдер", "подвесное кресло", "кресло подвесное",
                "кресло-кокон", "кресло кокон", "гамак", "садовое", "кемпингов", "туристическ", "рыболовн", "карпов", "для рыбалки",
                "шезлонг", "пляжн", "складное туристическое", "парикмахерск", "педикюрн", "для маникюра", "мастера маникюра",
                "косметолог", "гинеколог", "стоматолог", "донорск", "медицинск", "кресло-коляска", "кресло коляска", "инвалидн",
                "кресло-туалет", "кресло туалет", "кресло-кровать", "кресло кровать", "кресло раскладное для сна",
                "для кукол", "кукольное", "миниатюр", "игрушечн",
            )
            if any(wc in t_low for wc in wrong_chairs):
                return False

    phone_brands = (
        "телефон", "смартфон", "iphone", "айфон", "xiaomi", "redmi", "poco", "samsung", "самсунг",
        "galaxy", "realme", "honor", "huawei", "pixel", "oneplus", "infinix", "tecno", "oppo",
        "vivo", "motorola", "iqoo", "zte", "nothing phone"
    )
    want_phone = any(w in q_low for w in phone_brands) and not any(w in q_low for w in ("pad", "tab", "ноутбук", "laptop", "watch", "часы", "наушник", "buds", "redmibook", "matebook", "magicbook", "vivobook", "zenbook"))

    tablet_brands = (
        "планшет", "ipad", "айпад", "matepad", "galaxy tab", "xiaomi pad", "redmi pad", "lenovo tab",
        "honor pad", "realme pad", "surface pro"
    )
    want_tablet = any(w in q_low for w in tablet_brands) or (("pad" in q_low or "tab" in q_low) and not any(w in q_low for w in ("thinkpad", "ideapad")))

    laptop_markers = (
        "ноутбук", "ноут", "laptop", "macbook", "макбук", "thinkpad", "ideapad", "vivobook",
        "zenbook", "legion", "tuf", "rog", "matebook", "magicbook", "redmibook", "machcreator", "genesis",
        "7640hs", "8845hs", "7840hs", "8945hs", "ryzen 5", "ryzen 7", "ryzen 9", "core i3", "core i5", "core i7", "core i9"
    )
    want_laptop = any(w in q_low for w in laptop_markers) and not want_phone and not want_tablet
    want_earphones = any(w in q_low for w in ("наушник", "гарнитур", "airpods", "freebuds", "buds", "qcy", "tws"))

    # 2. СТРОГИЙ КАТЕГОРИЙНЫЙ ОТБОР (по названию и описанию)
    if want_case:
        if not any(w in haystack for w in ("чехол", "бампер", "накладк", "обложк", "кейс", "книжк", "кобур", "футляр")):
            return False
        junk_cases = ("на диван", "на кресло", "на стул", "для одежды", "для гитар", "для очков", "на чемодан", "на матрас", "на подушку", "для гладильн", "для карт", "ключниц")
        if any(j in t_low for j in junk_cases):
            return False
        if not want_earphones:
            earphone_markers = ("наушник", "airpods", "freebuds", "galaxy buds", "redmi buds", "qcy", "tws", "гарнитур", "аирподс")
            if any(em in t_low for em in earphone_markers):
                return False
        if want_phone and not want_tablet:
            if any(tm in t_low for tm in ("для планшета", "на планшет", "ipad", "matepad", "для ноутбука")):
                return False
        if want_tablet:
            if any(pm in t_low for pm in ("для телефона", "на телефон", "для смартфона", "на смартфон")):
                return False
            if not any(tm in haystack for tm in ("планшет", "ipad", "matepad", "tab", "pad")):
                return False

    elif want_tablet and not want_accessory:
        tablet_junk = (
            "чехол", "обложк", "книжк", "бампер", "накладк", "кейс", "стекло", "пленк", "гидрогел",
            "стилус", "клавиатура для", "держател", "подставк", "сумк", "рюкзак", "заряд", "кабель",
            "дисплей для", "тачскрин", "матрица", "шлейф", "аккумулятор для", "батарея для",
            "для рисования", "графический планшет", "световой планшет", "планшет с зажимом", "папка-планшет",
        )
        if any(j in t_low for j in tablet_junk):
            return False
        if not any(tm in haystack for tm in ("планшет", "ipad", "matepad", "tab", "pad")):
            return False

    elif want_phone and not want_accessory:
        phone_junk = (
            "чехол", "обложк", "книжк", "бампер", "накладк", "кейс", "стекло", "пленк", "гидрогел",
            "держател", "подставк", "заряд", "кабель", "провод", "блок питания", "наушник", "гарнитур", "ремешок",
            "дисплей", "экран для", "тачскрин", "матрица", "шлейф", "аккумулятор", "батарея", "задняя крышка",
            "корпус для", "плата", "камера для", "коробка от", "муляж",
        )
        if any(j in t_low for j in phone_junk):
            return False
        if any(j in haystack for j in ("только коробка", "коробка от телефона", "чехол накладка", "чехол-книжка", "защитное стекло 3d")):
            return False

    elif want_laptop and not want_accessory:
        non_laptop_markers = (
            "mini pc", "mini-pc", "мини пк", "мини-пк", "неттоп", "gmktec", "beelink", "minisforum", "geekom",
            "моноблок", "системный блок", "сокет am4", "сокет am5", "socket am4", "socket am5",
            "вид поставки: oem", "вид поставки oem", "вид поставки: box",
            "чехол для", "сумка для", "рюкзак для", "подставка для", "столик для", "зарядное для", "блок питания для",
            "матрица для", "клавиатура для", "аккумулятор для", "батарея для", "кулер для", "вентилятор для",
            "корпус для", "петли для", "топкейс", "поддон", "шлейф для",
        )
        if any(m in t_low for m in non_laptop_markers):
            return False
        if any(m in haystack for m in ("сокет am4", "сокет am5", "socket am4", "socket am5", "вид поставки: oem", "вид поставки: box")):
            return False

    # 2б. УНИВЕРСАЛЬНАЯ ПРОВЕРКА ТИПА ТОВАРА (табуретка → табуретка, стул → стул, кресло → кресло, корм → корм)
    if not (want_phone or want_tablet or want_laptop or want_earphones):
        head_noun = _query_head_noun(q_low)
        if head_noun:
            stem = _ru_stem(head_noun)
            if stem not in t_low and not ("кресл" in q_low and "стул" in t_low):
                return False
            # «Чехол для табуретки», «ножки для табурета», «подшипник для кресла» — это не сам товар
            if not want_accessory and stem in t_low:
                before_noun = t_low.split(stem, 1)[0]
                if re.search(r"(?:^|\s)(?:для|под|к|на|от|в)\s+(?:\S+\s+){0,3}$", before_noun):
                    return False
            # Дополнительная проверка первых 220 символов описания/сниппета на скрытые чехлы/запчасти
            if not want_accessory:
                snip_head = d_low[:220]
                if any(p in snip_head for p in ("чехол на ", "чехол для ", "накидка на ", "подшипник", "газлифт для", "крестовина для", "ролики для", "колесики для", "колёсики для", "заглушки для", "ремкомплект для")):
                    return False

    # 3. Базовые слабые процессоры (если их явно не просили)
    low_end_cpus = ["celeron", "pentium", "atom", "n100", "n95", "n5095", "n4020", "n4500", "n3350", "3020e", "3250u"]
    for weak in low_end_cpus:
        if weak not in q_low and weak in haystack:
            return False

    if "ryzen" in q_low and "intel" not in q_low:
        for intel_marker in ["core i3", "core i5", "core i7", "core i9", "i3-", "i5-", "i7-", "i9-", "intel", "ultra 5", "ultra 7", "ultra 9"]:
            if intel_marker in t_low:
                return False

    # 4. ПРОВЕРКА МОДЕЛИ И ИНДЕКСОВ В НАЗВАНИИ И ОПИСАНИИ ТОВАРА (haystack = title + description + url)
    alnum_tokens = re.findall(r"\b([a-zа-яё]*\d+[a-zа-яё0-9]*)\b", q_low)
    for tok in alnum_tokens:
        if tok in ("5", "7", "9") and "ryzen" in q_low and len(alnum_tokens) > 1:
            continue
        target_text = haystack
        m_split = re.match(r"^([a-zа-яё]*)(\d+)([a-zа-яё]*)$", tok)
        if m_split:
            prefix, digits, suffix = m_split.groups()
            if prefix and not suffix:
                pat1 = rf"\b{re.escape(prefix)}[-\s]?{re.escape(digits)}\b"
                pat2 = rf"\b{re.escape(digits)}[-\s]?{re.escape(prefix)}\b"
                if not (re.search(pat1, target_text) or re.search(pat2, target_text)):
                    return False
            elif suffix and not prefix:
                # Для процессоров (например 7640hs, 8845hs) допускаем также написание 7640 hs / 7640h / 7640 в описании
                pat1 = rf"\b{re.escape(digits)}[-\s]?{re.escape(suffix)}\b"
                pat2 = rf"\b{re.escape(suffix)}[-\s]?{re.escape(digits)}\b"
                pat3 = rf"\b{re.escape(digits)}\b" if (len(digits) == 4 and suffix in ("hs", "h", "hx", "u") and "ryzen" in haystack) else pat1
                if not (re.search(pat1, target_text) or re.search(pat2, target_text) or re.search(pat3, target_text)):
                    return False
            elif prefix:
                pat = rf"\b{re.escape(prefix)}[-\s]?{re.escape(digits)}"
                if not re.search(pat, target_text):
                    return False
            else:
                pat = rf"\b{re.escape(digits)}[a-zа-яё]*\b"
                if not re.search(pat, target_text):
                    return False

    # 5. Обязательные модификаторы модели (pro, super, ti, max, ultra, x3d, xt, slim, note, plus) — в названии ИЛИ описании
    strict_modifiers = ["super", "ti", "pro", "max", "ultra", "x3d", "xtx", "xt", "slim", "oled", "note", "plus"]
    for mod in strict_modifiers:
        if re.search(rf"\b{mod}\b", q_low):
            if not re.search(rf"\b{mod}\b", haystack):
                return False

    # 6. Смысловые слова из запроса должны присутствовать в названии ИЛИ в описании карточки
    stop_words = {
        "для", "под", "или", "без", "купить", "новый", "бу", "цена", "оригинал",
        "краснодар", "москва", "доставка", "россия", "вся", "регионы",
        "телефон", "телефона", "смартфон", "смартфона", "ноутбук", "ноутбука", "ноутбуки",
        "планшет", "планшета", "чехол", "чехлы", "бампер", "накладка", "процессор", "процессором", "описанию",
    }
    chair_adj_synonyms = ("комп", "офис", "игров", "геймер", "руковод", "оператор", "рабоч", "эргоном", "ортопед", "школьн", "кресло")
    words = [w for w in re.findall(r"[a-zа-яё]{3,}", q_low) if w not in stop_words]
    for w in words:
        if w == "ryzen" and any(r in haystack for r in ("ryzen", "райзен", "amd", "амд", "r5", "r7", "r9", "7640", "8845", "7840", "7535", "6600")):
            continue
        if "кресл" in q_low and w.startswith(("компьютер", "офисн", "игров", "геймер", "рабоч")):
            if any(syn in haystack for syn in chair_adj_synonyms):
                continue
        stem = w[:4] if len(w) >= 5 else w
        if stem not in haystack:
            return False

    return True


def _build_product(
    shop: str,
    query: str,
    title: str,
    price: int,
    seller_name: str,
    seller_rating: float,
    reviews_count: int,
    location_delivery: str,
    description: str,
    url: str,
    site_duty: int = 0,
    snippet_fallback: bool = True,
) -> ProductItem:
    report = analyze_item(
        title=title,
        description=description,
        seller_name=seller_name,
        seller_rating=seller_rating,
        reviews_count=reviews_count,
        query=query,
    )
    duty_rub, duty_note = calculate_customs_duty(
        shop=shop,
        base_price=price,
        title=title,
        description=description,
        site_duty=site_duty,
        snippet_fallback=snippet_fallback,
    )
    stock_str = detect_stock_status(shop=shop, title=title, description=description, url=url)
    matches_model = _matches_target_model(query, title, description, url) if query else True
    defects_text = report.defects_summary
    score_val = report.score

    # Проверяем, найдено ли совпадение именно благодаря описанию товара
    desc_specs = extract_description_specs(query, title, description)
    t_low = _normalize_cpu_text(title or "")
    d_low = _normalize_cpu_text(description or "")
    alnum_tokens = re.findall(r"\b([a-zа-яё]*\d+[a-zа-яё0-9]*)\b", _normalize_cpu_text(query or ""))
    matched_by_desc = False
    for tok in alnum_tokens:
        if len(tok) >= 3 and tok not in t_low and (tok in d_low or tok[:4] in d_low):
            matched_by_desc = True
            break

    # Проверяем, не расходится ли заголовок объявления с описанием (чужой бренд/модель в самом названии или нет данных в тексте)
    title_desc_mismatch = False
    mismatch_reason = ""
    hay_low = f"{t_low} {dehomoglyph_ru(description or '')}"
    if any(b in t_low for b in ("rikor", "nino 200")) and not any(b in (query or "").lower() for b in ("rikor", "nino")):
        title_desc_mismatch = True
        mismatch_reason = f"⚠️ В названии указана другая модель («{title[:35]}»), а процессор указан в описании"
    elif any(p in hay_low for p in ("в авито нет нужной модели", "авито не знает такую модель", "правильные характеристики на фото", "все описания в фото")):
        if matches_model and matched_by_desc and desc_specs:
            # Продавец честно написал модель/ЦП в описании при общем названии — не штрафуем до YELLOW только за заголовок
            mismatch_reason = f"✅ Найдено в описании ({desc_specs}), общий заголовок Авито «{title.strip()}»"
        else:
            title_desc_mismatch = True
            mismatch_reason = "⚠️ В названии шаблон Авито («нет нужной модели»), характеристики только на фото"
    elif matched_by_desc and len(t_low.split()) <= 2:
        if desc_specs:
            mismatch_reason = f"✅ Модель указана в описании ({desc_specs}), общий заголовок «{title.strip()}»"
        else:
            title_desc_mismatch = True
            mismatch_reason = f"⚠️ Общее название («{title}»), модель найдена только в описании ({desc_specs})"

    if matched_by_desc and matches_model:
        prefix_note = mismatch_reason or f"✅ Найдено по описанию ({desc_specs or 'совпадение характеристик'})"
        if defects_text and "Без поломок" not in defects_text:
            defects_text = f"{prefix_note}; {defects_text}"
        else:
            defects_text = prefix_note

    if desc_specs and desc_specs not in (description or ""):
        full_desc = f"[{desc_specs}] {description or ''}".strip()
    else:
        full_desc = (description or "").strip()

    if not matches_model:
        final_status_code = "RED"
        stock_str = "❌ Отбракован (Другая модель / Категория)"
        final_status_label = "🔴 КРАСНЫЙ (Не та модель / Аксессуар)"
        q_norm = _normalize_cpu_text(query or "")
        d_norm = _normalize_cpu_text(description or "")
        q_c = re.findall(r"\b(\d{4})(?:hs|h|hx|u|g|f|x3d|x)?\b", q_norm)
        d_c = re.findall(r"(?:\b(\d{4})(?:hs|h|hx|u|g|f|x3d|x)\b|ryzen\s*[3579]\s*(?:pro\s*)?(\d{4})\b)", d_norm)
        d_digits = [a or b for a, b in d_c if (a or b) and (a or b) not in q_c]
        if q_c and d_digits:
            defects_text = f"❌ В описании указан другой ЦП ({', '.join(sorted(set(d_digits)))} вместо {q_c[0]})"
        else:
            defects_text = f"Не соответствует модели/описанию «{query}»"
        score_val = 15
    elif stock_str.startswith("❌"):
        final_status_code = "RED"
        final_status_label = "🔴 КРАСНЫЙ (Нет в наличии / Распродан)"
        score_val = 10
    elif report.status_code in ("RED", "SKIP"):
        final_status_code = "RED"
        stock_str = "❌ Отбракован (Брак / Ремонт / Резерв)"
        final_status_label = report.status_label
    elif title_desc_mismatch:
        # Если в названии одно, а в описании другое — ставим ЖЁЛТЫЙ статус (не пускаем в Зелёный Топ-10!)
        final_status_code = "YELLOW"
        stock_str = "✅ В продаже (Проверить описание)"
        final_status_label = "🟡 ЖЁЛТЫЙ (Название расходится с описанием)"
        score_val = min(score_val, 68)
    else:
        final_status_code = report.status_code
        stock_str = "✅ В продаже (В наличии)"
        final_status_label = report.status_label

    total_price = price + duty_rub
    now_str = datetime.now().strftime("%d.%m.%Y %H:%M")
    return ProductItem(
        shop=shop,
        query=query,
        title=title.strip(),
        in_stock=stock_str,
        price=total_price,
        base_price=price,
        customs_duty=duty_rub,
        duty_note=duty_note,
        seller_name=(seller_name or "Не указан").strip(),
        seller_rating=round(float(seller_rating or 0.0), 2),
        reviews_count=int(reviews_count or 0),
        location_delivery=(location_delivery or "Все регионы РФ / Доставка").strip(),
        description=full_desc[:650],
        url=clean_product_url(url),
        status_code=final_status_code,
        status_label=final_status_label,
        defects_summary=defects_text,
        seller_type=report.seller_type,
        score=score_val,
        updated_at=now_str,
    )


def filter_category_outliers(items: List[ProductItem]) -> List[ProductItem]:
    """
    Универсальный статистический фильтр аномально дешёвых деталей, расходников и мусора
    для любых категорий (мебель, бытовая техника, электроника, канцелярия и т.д.).
    """
    by_q: dict[str, List[ProductItem]] = {}
    for it in items:
        if it.status_code in ("GREEN", "YELLOW") and it.price > 0 and str(it.in_stock).startswith("✅"):
            qk = (it.query or "").strip().lower()
            by_q.setdefault(qk, []).append(it)

    for qk, grp in by_q.items():
        if _is_accessory_query(qk) or len(grp) < 4:
            continue
        non_ali_prices = sorted(x.price for x in grp if x.shop != "AliExpress")
        ref_prices = non_ali_prices if len(non_ali_prices) >= 3 else sorted(x.price for x in grp)
        median_p = ref_prices[len(ref_prices) // 2]
        if median_p < 1200:
            continue

        for it in grp:
            t_low = (it.title or "").lower()
            is_ali_tc = it.shop == "AliExpress" and _is_aliexpress_titlecase_spam(it.title)
            has_prep = bool(re.search(r"\b(?:для|под|к|от|на)\s+\w+", t_low))
            susp_word = any(
                w in t_low
                for w in (
                    "аксессуар", "деталь", "часть", "замена", "сменн", "запасн", "накладк", "поднос",
                    "основан", "каркас", "подставк", "креплени", "чехол", "накидк", "комплект для",
                    "набор для", "многофункцион", "ремонт", "butterfly", "босса", "крышка", "насадка",
                )
            )
            is_suspicious_cheap = (is_ali_tc or has_prep or susp_word) and it.price < median_p * 0.36 and (median_p - it.price) >= 1200
            is_extreme_outlier = it.price < median_p * 0.16 and (median_p - it.price) >= 2000
            if is_suspicious_cheap or is_extreme_outlier:
                it.status_code = "RED"
                it.in_stock = "❌ Отбракован (Похож на деталь/расходник)"
                it.status_label = "🔴 КРАСНЫЙ (Аномально низкая цена — деталь/хлам)"
                it.defects_summary = f"❌ Цена {it.price:,} ₽ аномально ниже рынка (~{median_p:,} ₽) — вероятен расходник/деталь"
                it.score = min(int(it.score or 15), 12)
    return items


CITY_PRESETS = {
    "все регионы рф": ("Все регионы РФ", "all", 213),
    "все регионы": ("Все регионы РФ", "all", 213),
    "вся россия": ("Все регионы РФ", "all", 213),
    "россия": ("Все регионы РФ", "all", 213),
    "рф": ("Все регионы РФ", "all", 213),
    "all": ("Все регионы РФ", "all", 213),
    # 1. Города-миллионники и крупнейшие мегаполисы
    "москва": ("Москва", "moskva", 213),
    "мск": ("Москва", "moskva", 213),
    "санкт-петербург": ("Санкт-Петербург", "sankt-peterburg", 2),
    "спб": ("Санкт-Петербург", "sankt-peterburg", 2),
    "питер": ("Санкт-Петербург", "sankt-peterburg", 2),
    "новосибирск": ("Новосибирск", "novosibirsk", 65),
    "нск": ("Новосибирск", "novosibirsk", 65),
    "екатеринбург": ("Екатеринбург", "ekaterinburg", 54),
    "екб": ("Екатеринбург", "ekaterinburg", 54),
    "казань": ("Казань", "kazan", 43),
    "нижний новгород": ("Нижний Новгород", "nizhniy_novgorod", 47),
    "нн": ("Нижний Новгород", "nizhniy_novgorod", 47),
    "красноярск": ("Красноярск", "krasnoyarsk", 62),
    "челябинск": ("Челябинск", "chelyabinsk", 56),
    "самара": ("Самара", "samara", 51),
    "уфа": ("Уфа", "ufa", 172),
    "ростов-на-дону": ("Ростов-на-Дону", "rostov-na-donu", 39),
    "ростов": ("Ростов-на-Дону", "rostov-na-donu", 39),
    "краснодар": ("Краснодар", "krasnodar", 35),
    "омск": ("Омск", "omsk", 66),
    "воронеж": ("Воронеж", "voronezh", 193),
    "пермь": ("Пермь", "perm", 50),
    "волгоград": ("Волгоград", "volgograd", 38),
    # 2. Крупные областные, краевые и республиканские центры (500к - 900к)
    "саратов": ("Саратов", "saratov", 194),
    "тюмень": ("Тюмень", "tyumen", 55),
    "тольятти": ("Тольятти", "tolyatti", 240),
    "барнаул": ("Барнаул", "barnaul", 197),
    "махачкала": ("Махачкала", "mahachkala", 28),
    "ижевск": ("Ижевск", "izhevsk", 44),
    "хабаровск": ("Хабаровск", "habarovsk", 76),
    "ульяновск": ("Ульяновск", "ulyanovsk", 195),
    "иркутск": ("Иркутск", "irkutsk", 63),
    "владивосток": ("Владивосток", "vladivostok", 75),
    "ярославль": ("Ярославль", "yaroslavl", 16),
    "севастополь": ("Севастополь", "sevastopol", 959),
    "ставрополь": ("Ставрополь", "stavropol", 36),
    "томск": ("Томск", "tomsk", 67),
    "кемерово": ("Кемерово", "kemerovo", 64),
    "набережные челны": ("Набережные Челны", "naberezhnye_chelny", 236),
    "челны": ("Набережные Челны", "naberezhnye_chelny", 236),
    "оренбург": ("Оренбург", "orenburg", 48),
    "новокузнецк": ("Новокузнецк", "novokuznetsk", 237),
    "балашиха": ("Балашиха", "balashiha", 10716),
    "рязань": ("Рязань", "ryazan", 11),
    "чебоксары": ("Чебоксары", "cheboksary", 45),
    "калининград": ("Калининград", "kaliningrad", 22),
    "пенза": ("Пенза", "penza", 49),
    "липецк": ("Липецк", "lipetsk", 9),
    "киров": ("Киров", "kirovskaya_oblast_kirov", 46),
    "астрахань": ("Астрахань", "astrahan", 37),
    "тула": ("Тула", "tula", 15),
    "сочи": ("Сочи", "sochi", 239),
    "курск": ("Курск", "kursk", 8),
    "улан-удэ": ("Улан-Удэ", "ulan-ude", 198),
    "тверь": ("Тверь", "tver", 14),
    "магнитогорск": ("Магнитогорск", "magnitogorsk", 235),
    "сургут": ("Сургут", "surgut", 973),
    "брянск": ("Брянск", "bryansk", 191),
    "иваново": ("Иваново", "ivanovo", 5),
    "якутск": ("Якутск", "yakutsk", 74),
    "владимир": ("Владимир", "vladimir", 192),
    "симферополь": ("Симферополь", "simferopol", 146),
    "белгород": ("Белгород", "belgorod", 4),
    "нижний тагил": ("Нижний Тагил", "nizhniy_tagil", 11168),
    "калуга": ("Калуга", "kaluga", 6),
    "чита": ("Чита", "chita", 68),
    "грозный": ("Грозный", "groznyy", 1106),
    "волжский": ("Волжский", "volgogradskaya_oblast_volzhskiy", 10951),
    "смоленск": ("Смоленск", "smolensk", 12),
    "подольск": ("Подольск", "podolsk", 10747),
    "саранск": ("Саранск", "saransk", 42),
    "вологда": ("Вологда", "vologda", 21),
    "курган": ("Курган", "kurgan", 53),
    "череповец": ("Череповец", "cherepovets", 968),
    "орёл": ("Орёл", "orel", 10),
    "орел": ("Орёл", "orel", 10),
    "архангельск": ("Архангельск", "arhangelsk", 20),
    "владикавказ": ("Владикавказ", "vladikavkaz", 33),
    "нижневартовск": ("Нижневартовск", "nizhnevartovsk", 1091),
    "йошкар-ола": ("Йошкар-Ола", "yoshkar-ola", 41),
    "стерлитамак": ("Стерлитамак", "sterlitamak", 11115),
    "мурманск": ("Мурманск", "murmansk", 23),
    "кострома": ("Кострома", "kostroma", 7),
    "новороссийск": ("Новороссийск", "novorossiysk", 970),
    "тамбов": ("Тамбов", "tambov", 13),
    "химки": ("Химки", "himki", 10758),
    "мытищи": ("Мытищи", "mytischi", 10740),
    "нальчик": ("Нальчик", "nalchik", 30),
    "таганрог": ("Таганрог", "taganrog", 971),
    "нижнекамск": ("Нижнекамск", "nizhnekamsk", 11127),
    "благовещенск": ("Благовещенск", "amur_blagoveschensk", 77),
    "комсомольск-на-амуре": ("Комсомольск-на-Амуре", "komsomolsk-na-amure", 11453),
    "петрозаводск": ("Петрозаводск", "petrozavodsk", 18),
    "королёв": ("Королёв", "korolev", 20728),
    "королев": ("Королёв", "korolev", 20728),
    "шахты": ("Шахты", "shahty", 11053),
    "энгельс": ("Энгельс", "engels", 11147),
    "великий новгород": ("Великий Новгород", "velikiy_novgorod", 24),
    "новгород": ("Великий Новгород", "velikiy_novgorod", 24),
    "люберцы": ("Люберцы", "lyubertsy", 10738),
    "братск": ("Братск", "bratsk", 976),
    "старый оскол": ("Старый Оскол", "staryy_oskol", 10649),
    "ангарск": ("Ангарск", "angarsk", 11256),
    "сыктывкар": ("Сыктывкар", "syktyvkar", 19),
    "дзержинск": ("Дзержинск", "dzerzhinsk", 972),
    "псков": ("Псков", "pskov", 25),
    "орск": ("Орск", "orsk", 11091),
    "красногорск": ("Красногорск", "krasnogorsk", 10735),
    "армавир": ("Армавир", "armavir", 10987),
    "абакан": ("Абакан", "abakan", 1095),
    "балаково": ("Балаково", "balakovo", 11143),
    "бийск": ("Бийск", "biysk", 975),
    "южно-сахалинск": ("Южно-Сахалинск", "yuzhno-sahalinsk", 80),
    "одинцово": ("Одинцово", "odintsovo", 10743),
    "уссурийск": ("Уссурийск", "ussuriysk", 11426),
    "прокопьевск": ("Прокопьевск", "prokopevsk", 11291),
    "рыбинск": ("Рыбинск", "rybinsk", 10839),
    "норильск": ("Норильск", "norilsk", 11311),
    "волгодонск": ("Волгодонск", "volgodonsk", 11036),
    "сызрань": ("Сызрань", "syzran", 11139),
    "петропавловск-камчатский": ("Петропавловск-Камчатский", "petropavlovsk-kamchatskiy", 78),
    "камчатский": ("Петропавловск-Камчатский", "petropavlovsk-kamchatskiy", 78),
    "каменск-уральский": ("Каменск-Уральский", "kamensk-uralskiy", 11164),
    "новочеркасск": ("Новочеркасск", "novocherkassk", 238),
    "альметьевск": ("Альметьевск", "almetevsk", 11121),
    "златоуст": ("Златоуст", "zlatoust", 11202),
    "северодвинск": ("Северодвинск", "severodvinsk", 10849),
    "хасавюрт": ("Хасавюрт", "hasavyurt", 11010),
    "керчь": ("Керчь", "kerch", 11464),
    "домодедово": ("Домодедово", "domodedovo", 10725),
    "салават": ("Салават", "salavat", 11114),
    "миасс": ("Миасс", "miass", 11212),
    "копейск": ("Копейск", "kopeysk", 11206),
    "пятигорск": ("Пятигорск", "pyatigorsk", 11067),
    "коломна": ("Коломна", "kolomna", 10734),
    "майкоп": ("Майкоп", "maykop", 1093),
    "березники": ("Березники", "berezniki", 11103),
    "ковров": ("Ковров", "kovrov", 10664),
    "кисловодск": ("Кисловодск", "kislovodsk", 11062),
    "нефтекамск": ("Нефтекамск", "neftekamsk", 11111),
    "нефтеюганск": ("Нефтеюганск", "nefteyugansk", 11188),
    "батайск": ("Батайск", "bataysk", 11033),
    "черкесск": ("Черкесск", "cherkessk", 1104),
    "дербент": ("Дербент", "derbent", 11006),
    "щелково": ("Щёлково", "schelkovo", 10765),
    "щёлково": ("Щёлково", "schelkovo", 10765),
    "невинномысск": ("Невинномысск", "nevinnomyssk", 11064),
    "серпухов": ("Серпухов", "serpuhov", 10754),
    "первоуральск": ("Первоуральск", "pervouralsk", 11171),
    "раменское": ("Раменское", "ramenskoe", 10750),
    "кызыл": ("Кызыл", "kyzyl", 11333),
    "орехово-зуево": ("Орехово-Зуево", "orehovo-zuevo", 10745),
    "новый уренгой": ("Новый Уренгой", "novyy_urengoy", 11230),
    "уренгой": ("Новый Уренгой", "novyy_urengoy", 11230),
    "обнинск": ("Обнинск", "obninsk", 967),
    "назрань": ("Назрань", "nazran", 1092),
    "каспийск": ("Каспийск", "kaspiysk", 11008),
    "долгопрудный": ("Долгопрудный", "dolgoprudnyy", 214),
    "ессентуки": ("Ессентуки", "essentukskaya", 11057),
    "димитровград": ("Димитровград", "dimitrovgrad", 11155),
    "ноябрьск": ("Ноябрьск", "noyabrsk", 11232),
    "евпатория": ("Евпатория", "evpatoriya", 11463),
    "реутов": ("Реутов", "reutov", 21621),
    "пушкино": ("Пушкино", "pushkino", 10748),
    "жуковский": ("Жуковский", "zhukovskiy", 20571),
    "ханты-мансийск": ("Ханты-Мансийск", "hanty-mansiysk", 57),
    "ногинск": ("Ногинск", "noginsk", 10742),
    "артём": ("Артём", "artem", 11409),
    "артем": ("Артём", "artem", 11409),
    "ачинск": ("Ачинск", "achinsk", 11302),
    "северск": ("Северск", "seversk", 11351),
    "бердск": ("Бердск", "berdsk", 11314),
    "элиста": ("Элиста", "elista", 1094),
    "арзамас": ("Арзамас", "arzamas", 11080),
    "зеленодольск": ("Зеленодольск", "zelenodolsk", 11125),
    "геленджик": ("Геленджик", "gelendzhik", 10990),
    "анапа": ("Анапа", "anapa", 1107),
    "туапсе": ("Туапсе", "tuapse", 1058),
    "ейск": ("Ейск", "eysk", 10993),
    "кропоткин": ("Кропоткин", "kropotkin", 10995),
    "славянск-на-кубани": ("Славянск-на-Кубани", "slavyansk-na-kubani", 10997),
    "тимашевск": ("Тимашёвск", "timashevsk", 10998),
    "белореченск": ("Белореченск", "belorechensk", 10988),
    "горячий ключ": ("Горячий Ключ", "goryachiy_klyuch", 10991),
    "ялта": ("Ялта", "yalta", 11470),
    "феодосия": ("Феодосия", "feodosiya", 11469),
    "магадан": ("Магадан", "magadan", 79),
    "биробиджан": ("Биробиджан", "birobidzhan", 11393),
    "горно-алтайск": ("Горно-Алтайск", "gorno-altaysk", 11319),
    "салехард": ("Салехард", "salehard", 58),
    "нарьян-мар": ("Нарьян-Мар", "naryan-mar", 10857),
    "анадырь": ("Анадырь", "anadyr", 11458),
    "воркута": ("Воркута", "vorkuta", 10939),
    "ухта": ("Ухта", "uhta", 10945),
    "великие луки": ("Великие Луки", "velikie_luki", 10928),
    "выборг": ("Выборг", "vyborg", 969),
    "гатчина": ("Гатчина", "gatchina", 10867),
    "всеволожск": ("Всеволожск", "vsevolozhsk", 10865),
    "мурино": ("Мурино", "murino", 2),
    "кудрово": ("Кудрово", "kudrovo", 2),
    "тобольск": ("Тобольск", "tobolsk", 11175),
    "серов": ("Серов", "serov", 11172),
    "новоуральск": ("Новоуральск", "novouralsk", 11170),
    "озёрск": ("Озёрск", "ozersk", 11214),
    "саров": ("Саров", "sarov", 11083),
    "выкса": ("Выкса", "vyksa", 11081),
    "кстово": ("Кстово", "kstovo", 11082),
    "глазов": ("Глазов", "glazov", 11150),
    "сарапул": ("Сарапул", "sarapul", 11151),
    "воткинск": ("Воткинск", "votkinsk", 11149),
    "соликамск": ("Соликамск", "solikamsk", 11108),
    "чайковский": ("Чайковский", "chaykovskiy", 11110),
    "кузнецк": ("Кузнецк", "kuznetsk", 11095),
    "мичуринск": ("Мичуринск", "michurinsk", 10802),
    "россошь": ("Россошь", "rossosh", 10683),
    "борисоглебск": ("Борисоглебск", "borisoglebsk", 10675),
    "железногорск": ("Железногорск", "kurskaya_oblast_zheleznogorsk", 10705),
    "губкин": ("Губкин", "gubkin", 10645),
    "клин": ("Клин", "klin", 10733),
    "воскресенск": ("Воскресенск", "voskresensk", 10722),
    "дубна": ("Дубна", "dubna", 215),
    "егоревск": ("Егорьевск", "egorevsk", 10728),
    "чехов": ("Чехов", "chehov", 10761),
    "дмитров": ("Дмитров", "dmitrov", 10723),
    "ступино": ("Ступино", "stupino", 10756),
    "наро-фоминск": ("Наро-Фоминск", "naro-fominsk", 10741),
    "павловский посад": ("Павловский Посад", "pavlovskiy_posad", 10746),
    "солнечногорск": ("Солнечногорск", "solnechnogorsk", 10755),
    "фрязино": ("Фрязино", "fryazino", 21619),
    "лобня": ("Лобня", "lobnya", 216),
    "видное": ("Видное", "vidnoe", 10719),
    "котельники": ("Котельники", "kotelniki", 21651),
    "людиново": ("Людиново", "lyudinovo", 10695),
    "ржев": ("Ржев", "rzhev", 10820),
    "вышний волочек": ("Вышний Волочёк", "vyshniy_volochek", 10815),
    "кинешма": ("Кинешма", "kineshma", 10688),
    "шуя": ("Шуя", "shuya", 10690),
    "муром": ("Муром", "murom", 10668),
    "александров": ("Александров", "aleksandrov", 10658),
    "гусь-хрустальный": ("Гусь-Хрустальный", "gus-hrustalnyy", 10661),
    "новомосковск": ("Новомосковск", "novomoskovsk", 10830),
    "алексин": ("Алексин", "aleksin", 10825),
    "узловая": ("Узловая", "uzlovaya", 10832),
    "щекино": ("Щёкино", "schekino", 10834),
    "елец": ("Елец", "elets", 10712),
    "клинцы": ("Клинцы", "klintsy", 10653),
    "вязьма": ("Вязьма", "vyazma", 10795),
    "рословль": ("Рославль", "roslavl", 10797),
    "ярцево": ("Ярцево", "yartsevo", 10799),
    "сафоново": ("Сафоново", "safonovo", 10798),
    "котлас": ("Котлас", "kotlas", 10844),
    "апатиты": ("Апатиты", "apatity", 10894),
    "мончегорск": ("Мончегорск", "monchegorsk", 10897),
    "канск": ("Канск", "kansk", 11306),
    "железногорск-красноярск": ("Железногорск (Красноярский край)", "krasnoyarskiy_kray_zheleznogorsk", 20086),
    "зеленоград": ("Зеленоград", "moskva_zelenograd", 216),
    "минусинск": ("Минусинск", "minusinsk", 11309),
    "лесосибирск": ("Лесосибирск", "lesosibirsk", 11308),
    "назарово": ("Назарово", "nazarovo", 11310),
    "усолье-сибирское": ("Усолье-Сибирское", "usole-sibirskoe", 11273),
    "усть-илимск": ("Усть-Илимск", "ust-ilimsk", 11274),
    "черемхово": ("Черемхово", "cheremhovo", 11275),
    "белово": ("Белово", "belovo", 11282),
    "киселевск": ("Киселёвск", "kiselevsk", 11285),
    "ленинск-кузнецкий": ("Ленинск-Кузнецкий", "leninsk-kuznetskiy", 11287),
    "междуреченск": ("Междуреченск", "mezhdurechensk", 11288),
    "юрга": ("Юрга", "yurga", 11294),
    "анжеро-судженск": ("Анжеро-Судженск", "anzhero-sudzhensk", 11281),
    "рубцовск": ("Рубцовск", "rubtsovsk", 11251),
    "новоалтайск": ("Новоалтайск", "novoaltaysk", 11250),
    "заринск": ("Заринск", "zarinsk", 11246),
    "ишим": ("Ишим", "ishim", 11173),
    "когалым": ("Когалым", "kogalym", 11184),
    "нягань": ("Нягань", "nyagan", 11189),
    "мегион": ("Мегион", "megion", 11186),
    "радужный": ("Радужный", "raduzhnyy", 11191),
    "урай": ("Урай", "uray", 11193),
    "лангепас": ("Лангепас", "langepas", 11185),
    "пыть-ях": ("Пыть-Ях", "pyt-yah", 11190),
    "надым": ("Надым", "nadym", 11229),
    "муравленко": ("Муравленко", "muravlenko", 11228),
    "губкинский": ("Губкинский", "gubkinskiy", 11226),
    "лабытнанги": ("Лабытнанги", "labytnangi", 11227),
    "бузулук": ("Бузулук", "buzuluk", 11086),
    "бугуруслан": ("Бугуруслан", "buguruslan", 11085),
    "новотроицк": ("Новотроицк", "novotroitsk", 11090),
    "бугульма": ("Бугульма", "bugulma", 11122),
    "елабуга": ("Елабуга", "elabuga", 11123),
    "лениногорск": ("Лениногорск", "leninogorsk", 11126),
    "чистополь": ("Чистополь", "chistopol", 11129),
    "заинск": ("Заинск", "zainsk", 11124),
    "азнакаево": ("Азнакаево", "aznakaevo", 11120),
    "туймазы": ("Туймазы", "tuymazy", 11116),
    "октябрьский": ("Октябрьский", "oktyabrskiy", 11112),
    "белорецк": ("Белорецк", "beloretsk", 11105),
    "ишимбай": ("Ишимбай", "ishimbay", 11107),
    "сибай": ("Сибай", "sibay", 11113),
    "кумертау": ("Кумертау", "kumertau", 11109),
    "белебей": ("Белебей", "belebey", 11104),
    "мелеуз": ("Мелеуз", "meleuz", 11110),
    "бирск": ("Бирск", "birsk", 11106),
    "учалы": ("Учалы", "uchaly", 11117),
    "новочебоксарск": ("Новочебоксарск", "novocheboksarsk", 11158),
    "канаш": ("Канаш", "kanash", 11157),
    "волжск": ("Волжск", "mariy_el_volzhsk", 11153),
    "козьмодемьянск": ("Козьмодемьянск", "kozmodemyansk", 11154),
    "рузаевка": ("Рузаевка", "ruzaevka", 11152),
    "жигулевск": ("Жигулёвск", "zhigulevsk", 11132),
    "новокуйбышевск": ("Новокуйбышевск", "novokuybyshevsk", 11135),
    "чапаевск": ("Чапаевск", "chapaevsk", 11140),
    "отрадный": ("Отрадный", "otradnyy", 11137),
    "вольск": ("Вольск", "volsk", 11144),
    "пугачев": ("Пугачёв", "pugachev", 11145),
    "ртищево": ("Ртищево", "rtischevo", 11146),
    "камышин": ("Камышин", "kamyshin", 10959),
    "михайловка": ("Михайловка", "mihaylovka", 10968),
    "урюпинск": ("Урюпинск", "uryupinsk", 10979),
    "азов": ("Азов", "azov", 11030),
    "гуково": ("Гуково", "gukovo", 11038),
    "донецк": ("Донецк (Ростовская обл.)", "rostovskaya_oblast_donetsk", 11039),
    "каменск-шахтинский": ("Каменск-Шахтинский", "kamensk-shahtinskiy", 11043),
    "красный сулин": ("Красный Сулин", "krasnyy_sulin", 11045),
    "миллерово": ("Миллерово", "millerovo", 11046),
    "новошахтинск": ("Новошахтинск", "novoshahtinsk", 11049),
    "сальск": ("Сальск", "salsk", 11051),
    "тихорецк": ("Тихорецк", "tihoretsk", 10999),
    "крымск": ("Крымск", "krymsk", 10994),
    "лабинск": ("Лабинск", "labinsk", 10996),
    "кореновск": ("Кореновск", "korenovsk", 10992),
    "темрюк": ("Темрюк", "temryuk", 11000),
    "абинск": ("Абинск", "abinsk", 10985),
    "апшеронск": ("Апшеронск", "apsheronsk", 10986),
    "гулькевичи": ("Гулькевичи", "gulkevichi", 10989),
    "приморско-ахтарск": ("Приморско-Ахтарск", "primorsko-ahtarsk", 11001),
    "усть-лабинск": ("Усть-Лабинск", "ust-labinsk", 11002),
    "георгиевск": ("Георгиевск", "georgievsk", 11056),
    "минеральные воды": ("Минеральные Воды", "mineralnye_vody", 11063),
    "минводы": ("Минеральные Воды", "mineralnye_vody", 11063),
    "михайловск": ("Михайловск", "shpakovskoe", 11068),
    "буденновск": ("Будённовск", "budennovsk", 11054),
    "изобильный": ("Изобильный", "izobilnyy", 11059),
    "светлоград": ("Светлоград", "svetlograd", 11066),
    "зеленокумск": ("Зеленокумск", "zelenokumsk", 11058),
    "благодарный": ("Благодарный", "blagodarnyy", 11053),
    "моздок": ("Моздок", "mozdok", 11022),
    "прохладный": ("Прохладный", "prohladnyy", 11014),
    "баксан": ("Баксан", "baksan", 11012),
    "гудермес": ("Гудермес", "gudermes", 11025),
    "шали": ("Шали", "shali", 11027),
    "урус-мартан": ("Урус-Мартан", "urus-martan", 11026),
    "буйнакск": ("Буйнакск", "buynaksk", 11005),
    "избербаш": ("Избербаш", "izberbash", 11007),
    "кизляр": ("Кизляр", "kizlyar", 11009),
    "кизилюрт": ("Кизилюрт", "kizilyurt", 11011),
    "находка": ("Находка", "nahodka", 974),
    "арсеньев": ("Арсеньев", "arsenev", 11408),
    "спасск-дальний": ("Спасск-Дальний", "spassk-dalniy", 11424),
    "большой камень": ("Большой Камень", "bolshoy_kamen", 11410),
    "партизанск": ("Партизанск", "partizansk", 11421),
    "лесозаводск": ("Лесозаводск", "lesozavodsk", 11417),
    "дальнегорск": ("Дальнегорск", "dalnegorsk", 11412),
    "амурск": ("Амурск", "amursk", 11449),
    "советская гавань": ("Советская Гавань", "sovetskaya_gavan", 11456),
    "николаевск-на-амуре": ("Николаевск-на-Амуре", "nikolaevsk-na-amure", 11455),
    "белогорск": ("Белогорск", "belogorsk", 11376),
    "свободный": ("Свободный", "svobodnyy", 11386),
    "тында": ("Тында", "tynda", 11390),
    "зея": ("Зея", "zeya", 11379),
    "нерюнгри": ("Нерюнгри", "neryungri", 11437),
    "мирный": ("Мирный", "mirnyy", 11435),
    "ленск": ("Ленск", "lensk", 11434),
    "алдан": ("Алдан", "aldan", 11430),
    "корсаков": ("Корсаков", "korsakov", 11450),
    "холмск": ("Холмск", "holmsk", 11452),
    "оха": ("Оха", "oha", 11451),
    "елизово": ("Елизово", "elizovo", 11397),
    "вилючинск": ("Вилючинск", "vilyuchinsk", 11396),
    "краснокаменск": ("Краснокаменск", "krasnokamensk", 11362),
    "борзя": ("Борзя", "borzya", 11358),
    "северобайкальск": ("Северобайкальск", "severobaykalsk", 11340),
    "гусиноозерск": ("Гусиноозёрск", "gusinoozersk", 11337),
    "саяногорск": ("Саяногорск", "sayanogorsk", 11343),
    "черногорск": ("Черногорск", "chernogorsk", 11344),
}


def _translit_avito_slug(text: str) -> str:
    table = {
        "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e", "ё": "e",
        "ж": "zh", "з": "z", "и": "i", "й": "y", "к": "k", "л": "l", "м": "m",
        "н": "n", "о": "o", "п": "p", "р": "r", "с": "s", "т": "t", "у": "u",
        "ф": "f", "х": "h", "ц": "ts", "ч": "ch", "ш": "sh", "щ": "sch",
        "ъ": "", "ы": "y", "ь": "", "э": "e", "ю": "yu", "я": "ya",
        " ": "_", "-": "-",
    }
    res = []
    for ch in text.lower().strip():
        if ch in table:
            res.append(table[ch])
        elif ch.isalnum() or ch in ("_", "-"):
            res.append(ch)
    slug = re.sub(r"_+", "_", "".join(res)).strip("_-")
    return slug or "all"


def resolve_city_settings(raw_city: str) -> dict:
    """
    Преобразует введённое пользователем название города (на русском или английском)
    в полный набор региональных настроек для всех маркетплейсов.
    """
    clean = (raw_city or "").strip()
    if not clean:
        return {
            "city_name": "Все регионы РФ",
            "region_slug_avito": "all",
            "ymarket_lr": 213,
        }
    key = re.sub(r"\s+", " ", clean.lower())
    if key in CITY_PRESETS:
        c_name, slug, lr = CITY_PRESETS[key]
        return {
            "city_name": c_name,
            "region_slug_avito": slug,
            "ymarket_lr": lr,
        }
    # Проверяем совпадение по avito slug
    for _, (c_name, slug, lr) in CITY_PRESETS.items():
        if key == slug:
            return {
                "city_name": c_name,
                "region_slug_avito": slug,
                "ymarket_lr": lr,
            }
    # Автоматическая транслитерация для любого другого города РФ
    pretty_name = "-".join(w.capitalize() for w in clean.split("-")) if "-" in clean else clean.title()
    return {
        "city_name": pretty_name,
        "region_slug_avito": _translit_avito_slug(clean),
        "ymarket_lr": 213,
    }


CITY_COORDS = [
    ("Москва", 55.7558, 37.6173),
    ("Санкт-Петербург", 59.9343, 30.3351),
    ("Краснодар", 45.0355, 38.9753),
    ("Новороссийск", 44.7239, 37.7687),
    ("Сочи", 43.5855, 39.7231),
    ("Геленджик", 44.5611, 38.0768),
    ("Анапа", 44.8949, 37.3169),
    ("Армавир", 44.9969, 41.1297),
    ("Туапсе", 44.0978, 39.0768),
    ("Ейск", 46.7112, 38.2738),
    ("Ростов-на-Дону", 47.2357, 39.7015),
    ("Таганрог", 47.2362, 38.8969),
    ("Новосибирск", 55.0084, 82.9357),
    ("Екатеринбург", 56.8389, 60.6057),
    ("Нижний Тагил", 57.9194, 59.9650),
    ("Казань", 55.7961, 49.1064),
    ("Набережные Челны", 55.7436, 52.3958),
    ("Нижнекамск", 55.6366, 51.8245),
    ("Нижний Новгород", 56.2965, 43.9361),
    ("Самара", 53.1959, 50.1002),
    ("Тольятти", 53.5078, 49.4204),
    ("Уфа", 54.7388, 55.9721),
    ("Стерлитамак", 53.6246, 55.9502),
    ("Челябинск", 55.1644, 61.4368),
    ("Магнитогорск", 53.4072, 58.9791),
    ("Омск", 54.9885, 73.3242),
    ("Красноярск", 56.0153, 92.8932),
    ("Пермь", 58.0105, 56.2502),
    ("Воронеж", 51.6608, 39.2003),
    ("Волгоград", 48.7080, 44.5133),
    ("Волжский", 48.7858, 44.7797),
    ("Тюмень", 57.1522, 65.5272),
    ("Сургут", 61.2540, 73.3962),
    ("Нижневартовск", 60.9344, 76.5531),
    ("Саратов", 51.5336, 46.0343),
    ("Энгельс", 51.4823, 46.1258),
    ("Барнаул", 53.3548, 83.7698),
    ("Иркутск", 52.2870, 104.3050),
    ("Братск", 56.1514, 101.6340),
    ("Хабаровск", 48.4802, 135.0719),
    ("Владивосток", 43.1155, 131.8855),
    ("Калининград", 54.7104, 20.4522),
    ("Ярославль", 57.6261, 39.8845),
    ("Ижевск", 56.8526, 53.2069),
    ("Оренбург", 51.7682, 55.0970),
    ("Орск", 51.2294, 58.4752),
    ("Тула", 54.1931, 37.6173),
    ("Тверь", 56.8587, 35.9176),
    ("Рязань", 54.6295, 39.7419),
    ("Белгород", 50.5954, 36.5873),
    ("Ставрополь", 45.0445, 41.9691),
    ("Пятигорск", 44.0486, 43.0594),
    ("Кисловодск", 43.9052, 42.7168),
    ("Кемерово", 55.3547, 86.0873),
    ("Новокузнецк", 53.7576, 87.1360),
    ("Томск", 56.4846, 84.9476),
    ("Астрахань", 46.3476, 48.0302),
    ("Киров", 58.6036, 49.6680),
    ("Чебоксары", 56.1463, 47.2511),
    ("Курск", 51.7304, 36.1926),
    ("Липецк", 52.6088, 39.5992),
    ("Брянск", 53.2436, 34.3634),
    ("Иваново", 57.0003, 40.9739),
    ("Владимир", 56.1291, 40.4066),
    ("Пенза", 53.1951, 45.0183),
    ("Ульяновск", 54.3142, 48.4031),
    ("Архангельск", 64.5399, 40.5158),
    ("Мурманск", 68.9707, 33.0749),
    ("Махачкала", 42.9831, 47.5047),
    ("Нальчик", 43.4853, 43.6071),
    ("Владикавказ", 43.0212, 44.6819),
    ("Грозный", 43.3184, 45.6986),
    ("Севастополь", 44.6167, 33.5254),
    ("Симферополь", 44.9521, 34.1024),
    ("Ялта", 44.4952, 34.1663),
    ("Чита", 52.0330, 113.5010),
    ("Улан-Удэ", 51.8348, 107.5845),
    ("Якутск", 62.0273, 129.7322),
    ("Петропавловск-Камчатский", 53.0370, 158.6559),
    ("Южно-Сахалинск", 46.9591, 142.7380),
    ("Благовещенск", 50.2907, 127.5272),
    ("Псков", 57.8194, 28.3318),
    ("Великий Новгород", 58.5228, 31.2698),
    ("Вологда", 59.2205, 39.8915),
    ("Череповец", 59.1264, 37.9092),
    ("Кострома", 57.7679, 40.9269),
    ("Калуга", 54.5138, 36.2612),
    ("Смоленск", 54.7826, 32.0453),
    ("Орёл", 52.9708, 36.0643),
    ("Тамбов", 52.7212, 41.4523),
    ("Саранск", 54.1874, 45.1839),
    ("Йошкар-Ола", 56.6316, 47.8862),
    ("Сыктывкар", 61.6688, 50.8365),
    ("Петрозаводск", 61.7850, 34.3469),
    ("Курган", 55.4410, 65.3411),
    ("Абакан", 53.7212, 91.4424),
    ("Подольск", 55.4312, 37.5458),
    ("Химки", 55.8894, 37.4449),
    ("Балашиха", 55.7963, 37.9382),
    ("Мытищи", 55.9105, 37.7363),
    ("Люберцы", 55.6787, 37.8939),
    ("Королёв", 55.9222, 37.8548),
    ("Одинцово", 55.6788, 37.2639),
    ("Домодедово", 55.4369, 37.7666),
    ("Зеленоград", 55.9919, 37.2144),
]


def _haversine_km(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    r = 6371.0
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp = math.radians(lat2 - lat1)
    dl = math.radians(lon2 - lon1)
    a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2.0 * r * math.asin(math.sqrt(max(0.0, min(1.0, a))))


def resolve_city_by_coords(lat: float, lon: float) -> dict:
    """
    Мгновенно определяет ближайший город РФ по GPS-координатам Telegram (без задержек).
    """
    best_name, best_dist = "Москва", 99999.0
    for c_name, c_lat, c_lon in CITY_COORDS:
        d = _haversine_km(lat, lon, c_lat, c_lon)
        if d < best_dist:
            best_dist = d
            best_name = c_name
    res = resolve_city_settings(best_name)
    res["distance_km"] = round(best_dist, 1)
    res["detected_place"] = f"{res['city_name']} (~{res['distance_km']} км)" if best_dist > 5 else res["city_name"]
    res["source"] = "coords"
    return res


async def resolve_city_by_coords_async(lat: float, lon: float) -> dict:
    """
    Мгновенно находит регион по координатам (локально < 1 мс), а если пользователь
    за пределами центра крупного города (> 8 км), быстро уточняет населённый пункт
    через геокодер с жёстким таймаутом 1.6 с и авто-откатом к ближайшему городу.
    """
    base = resolve_city_by_coords(lat, lon)
    if base["distance_km"] <= 8.0:
        return base
    try:
        async with httpx.AsyncClient(timeout=1.6, headers={"User-Agent": "PriceMonitorBot/1.0"}) as cli:
            resp = await cli.get(
                "https://nominatim.openstreetmap.org/reverse",
                params={
                    "format": "jsonv2",
                    "lat": f"{lat:.6f}",
                    "lon": f"{lon:.6f}",
                    "accept-language": "ru",
                    "zoom": "14",
                },
            )
            if resp.status_code == 200:
                data = resp.json() or {}
                addr = data.get("address") or {}
                country = (addr.get("country_code") or "").lower()
                raw_place = (
                    addr.get("city")
                    or addr.get("town")
                    or addr.get("village")
                    or addr.get("hamlet")
                    or addr.get("municipality")
                    or ""
                ).strip()
                state = (addr.get("state") or "").strip()
                if not raw_place and state and (not country or country == "ru"):
                    return {
                        **base,
                        "detected_place": f"{base['city_name']} (~{base['distance_km']} км, {state})",
                        "source": "geocoder+coords",
                    }
                if raw_place and (not country or country == "ru"):
                    clean_place = re.sub(r"^(?:городской округ|город|г\.|посёлок|поселок|село|станица|хутор)\s+", "", raw_place, flags=re.IGNORECASE).strip()
                    key = re.sub(r"\s+", " ", clean_place.lower())
                    if key in CITY_PRESETS:
                        exact = resolve_city_settings(clean_place)
                        exact["distance_km"] = 0.0
                        exact["detected_place"] = f"{exact['city_name']}" + (f" ({state})" if state and state != exact["city_name"] else "")
                        exact["source"] = "geocoder"
                        return exact
                    # Небольшой город/станица — используем валидный slug ближайшего центра РФ для Авито (+ точные lat/lon и радиус 5-15 км)
                    return {
                        "city_name": clean_place,
                        "region_slug_avito": base["region_slug_avito"],
                        "ymarket_lr": base["ymarket_lr"],
                        "distance_km": 0.0,
                        "detected_place": f"{clean_place}" + (f" ({state})" if state and state != clean_place else ""),
                        "source": "geocoder",
                    }
    except Exception:
        pass
    return base


_WB_GEO_CACHE: dict[str, dict] = {}


def get_city_coords(
    city_name: str = "Москва",
    lat: float | None = None,
    lon: float | None = None,
) -> tuple[float, float]:
    """Возвращает координаты для указанного города (или точные GPS-координаты пользователя)."""
    if lat is not None and lon is not None:
        return float(lat), float(lon)
    clean = re.sub(r"\s+", " ", (city_name or "").strip().lower()).replace("ё", "е")
    if clean and clean not in ("все регионы рф", "все регионы", "вся россия", "россия", "рф", "all"):
        for c_name, c_lat, c_lon in CITY_COORDS:
            c_low = c_name.lower().replace("ё", "е")
            if c_low == clean or (len(clean) >= 4 and clean in c_low):
                return c_lat, c_lon
    return 55.7558, 37.6173


async def fetch_wb_geo_payload(lat: float, lon: float) -> dict | None:
    """Получает региональный объект Wildberries (с dest и xinfo) для заданных координат города."""
    cache_key = f"{round(float(lat), 2)}_{round(float(lon), 2)}"
    if cache_key in _WB_GEO_CACHE:
        return _WB_GEO_CACHE[cache_key]
    try:
        async with httpx.AsyncClient(
            trust_env=False,
            timeout=5.5,
            headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/126.0.0.0 Safari/537.36"},
        ) as cli:
            resp = await cli.get(
                "https://user-geo-data.wildberries.ru/get-geo-info",
                params={"latitude": float(lat), "longitude": float(lon), "currency": "RUB"},
            )
            if resp.status_code == 200:
                data = resp.json() or {}
                if isinstance(data, dict) and data.get("xinfo"):
                    _WB_GEO_CACHE[cache_key] = data
                    return data
    except Exception:
        pass
    return None


def _extract_wb_dest(wb_geo: dict | None) -> str:
    xinfo = (wb_geo or {}).get("xinfo", "")
    for part in str(xinfo).split("&"):
        if part.startswith("dest="):
            return part.split("=", 1)[1].strip()
    return ""


_JS_DUTY_HELPER = r"""
        const _extractDutyRub = (txt) => {
            const s = (txt || '').replace(/\s+/g, ' ').trim();
            if (!s || !/(?:пошлин|тамож)/i.test(s)) return 0;
            const low = s.toLowerCase();
            if (/беспошлин|порог|превышени|при заказе от|при заказе до|при покупке от|свыше 200|без пошлин|пошлина не|пошлины нет|пошлина включ|уже включ/i.test(low)) return 0;
            const m = s.match(/(?:[+~≈]\s*)?(\d[\d\s\u00a0\u2006\u2009\u202f]{0,7})\s*₽\s*(?:[-—–•|\s]*)(?:(?:доплат\w*\s+за\s+)?таможен\w*\s+(?:пошлин\w*|плат[её]ж\w*|сбор\w*)|пошлин\w*)/i) ||
                      s.match(/(?:(?:доплат\w*\s+за\s+)?таможен\w*\s+(?:пошлин\w*|плат[её]ж\w*|сбор\w*)|пошлин\w*)\s*(?:около|примерно|составит|ориентировочно|при\s+получении|на\s+таможне|оплачивается\s+отдельно|за\s+заказ)?\s*[~≈:+=-]?\s*(\d[\d\s\u00a0\u2006\u2009\u202f]{0,7})\s*₽/i);
            return m ? (parseInt(m[1].replace(/\D/g, ''), 10) || 0) : 0;
        };
"""


_OZON_NOISE_PHRASES = (
    "распродажа",
    "осталось",
    "осталась",
    "цена что надо",
    "уценка",
    "уценённый",
    "уцененный",
    "скидка",
    "баллов за отзыв",
    "стало дешевле",
    "оригинал",
    "чёрная пятница",
    "черная пятница",
    "морков",
    "суперцена",
    "неделя скидок",
)


def _is_ozon_noise_title(text: str) -> bool:
    """Определяет рекламные бейджи и счётчики акций вместо названия товара Ozon."""
    s = re.sub(r"\s+", " ", str(text or "")).strip()
    if len(s) < 6:
        return True
    if "₽" in s or "%" in s:
        return True
    low = s.lower()
    if re.match(r"^[+\-−]\s*\d", low) or re.match(r"^[0-9.,\s]+$", low):
        return True
    if any(low == bw or (bw in low and len(s) < 42) for bw in _OZON_NOISE_PHRASES):
        return True
    if low in ("хит", "новинка", "рекомендуем") or "в корзину" in low or "баллов" in low:
        return True
    return False


def clean_ozon_card_title(title: str, description: str = "") -> str:
    """Возвращает очищенное название товара Ozon, отсекая бейджи вроде 'Чёрная пятница'."""
    t = re.sub(r"\s+", " ", str(title or "")).strip()
    if not _is_ozon_noise_title(t):
        return t
    for part in str(description or "").split("|"):
        cand = re.sub(r"\s+", " ", part).strip()
        if len(cand) > 10 and not _is_ozon_noise_title(cand):
            low = cand.lower()
            if not any(w in low for w in ("отзыв", "доставим", "завтра", "пошлин", "тамож", "в корзину", "баллов", "шт")):
                return cand
    return ""


async def parse_ozon(
    engine: IsolatedChromeEngine,
    query: str,
    min_price: int,
    max_price: int,
    limit: int = 100,
    city_name: str = "Все регионы РФ",
) -> List[ProductItem]:
    """
    1. Приоритет №1: Ozon (сортировка по возрастанию цены + считывание пошлины с карточки).
    """
    await engine.warmup_url(
        "https://www.ozon.ru/",
        wait_sec=8.0,
        ready_selector='a[href*="/product/"], a[href*="/category/"], input[name="text"]',
    )
    min_price = _effective_min_price(query, min_price)
    encoded_q = urllib.parse.quote(query)
    if "ноутбук" in query.lower():
        url = (
            f"https://www.ozon.ru/category/noutbuki-15692/?text={encoded_q}"
            f"&sorting=price&currency_price={min_price}.000%3B{max_price}.000&from_global=true"
        )
    else:
        url = (
            f"https://www.ozon.ru/search/?text={encoded_q}"
            f"&sorting=price&currency_price={min_price}.000%3B{max_price}.000&from_global=true"
        )

    js_extractor = """
    (async () => {
__DUTY_HELPER__
        const items = [];
        const seen = new Set();
        const badgeWords = [
            'распродажа', 'осталось', 'осталась', 'цена что надо', 'уценка', 'уценённый', 'уцененный',
            'скидка', 'баллов за отзыв', 'хит', 'стало дешевле', 'оригинал',
            'чёрная пятница', 'черная пятница', 'морков', 'суперцена', 'неделя скидок'
        ];
        const isNoiseTitle = (s) => {
            const v = (s || '').replace(/\\s+/g, ' ').trim();
            const lv = v.toLowerCase();
            if (!v || v.length < 6) return true;
            if (v.includes('₽') || v.includes('%')) return true;
            if (/^[+\\-−]\\s*\\d/.test(lv) || /^[0-9.,\\s]+$/.test(lv)) return true;
            return badgeWords.some(bw => lv === bw || lv.includes(bw));
        };

        for (let step = 0; step < 7; step++) {
            const links = Array.from(document.querySelectorAll('a[href*="/product/"]'));

            for (const a of links) {
                const href = a.href || '';
                const baseHref = href.split('?')[0];
                if (!baseHref || seen.has(baseHref)) continue;

                const linkText = (a.innerText || '').replace(/\\s+/g, ' ').trim();
                if (badgeWords.some(bw => {
                    const lowLink = linkText.toLowerCase();
                    return lowLink === bw || (lowLink.startsWith(bw) && linkText.length < 34);
                })) continue;

                let card = a;
                for (let i = 0; i < 7; i++) {
                    if (!card.parentElement) break;
                    card = card.parentElement;
                    const t = card.innerText || '';
                    if (t.includes('₽') && t.length > 20 && t.length < 2200) {
                        break;
                    }
                }

                const textBlock = (card.innerText || '').replace(/\\n\\s*₽/g, ' ₽').trim();
                if (!textBlock.includes('₽') || textBlock.includes('₸') || textBlock.includes('KZT')) continue;

                const lowBlock = textBlock.toLowerCase();
                if (lowBlock.includes('похожие')) {
                    continue;
                }

                const lines = textBlock.split('\\n').map(s => s.trim()).filter(Boolean);
                const flatBlock = lines.join(' ');

                let price = 0;
                let site_duty = 0;
                const flatDv = _extractDutyRub(flatBlock);
                if (flatDv > 0) site_duty = flatDv;
                const priceCandidates = [];
                for (const line of lines) {
                    const lowLine = line.toLowerCase();
                    const dv = _extractDutyRub(line);
                    if (dv > 0) {
                        site_duty = dv;
                        continue;
                    }
                    if (
                        lowLine.includes('пошлин') ||
                        lowLine.includes('тамож') ||
                        lowLine.includes('баллов') ||
                        lowLine.includes('балл') ||
                        lowLine.includes('бонус') ||
                        lowLine.includes('купон') ||
                        lowLine.includes('скидк') ||
                        lowLine.includes('выгод') ||
                        lowLine.includes('кешбэк') ||
                        lowLine.includes('кэшбэк') ||
                        /\\/\\s*мес/i.test(lowLine) ||
                        /(^|\\s)в\\s+мес($|\\s|\\.)/i.test(lowLine) ||
                        lowLine.includes('мес.') ||
                        lowLine.includes('месяц') ||
                        lowLine.includes('рассроч') ||
                        lowLine.includes('кредит') ||
                        lowLine.includes('сплит') ||
                        lowLine.includes('долями') ||
                        lowLine.includes('доставк') ||
                        /[×xх]\\s*(2|4|6|10|12)\\b/i.test(lowLine)
                    ) {
                        continue;
                    }
                    if (line.includes('₽')) {
                        const cleanLine = line.replace(/[\\s\\u00a0\\u2006\\u2009\\u202f]+/g, '');
                        const m = cleanLine.match(/(\\d{2,7})₽/);
                        if (m) {
                            priceCandidates.push(parseInt(m[1], 10));
                        }
                    }
                }
                if (priceCandidates.length > 0) {
                    price = Math.min(...priceCandidates);
                }
                if (!price) continue;

                let title = '';
                for (const l of lines) {
                    const ll = l.toLowerCase();
                    if (
                        !isNoiseTitle(l) &&
                        !ll.includes('бренд') &&
                        !ll.includes('вау-цен') &&
                        !ll.includes('балл') &&
                        !ll.includes('бонус') &&
                        !ll.includes('отзыв') &&
                        !ll.includes('в корзину') &&
                        !ll.includes('купить') &&
                        !ll.includes('доставим') &&
                        !ll.includes('завтра') &&
                        !ll.includes('рассроч') &&
                        !ll.includes('сплит') &&
                        !ll.includes('долями') &&
                        !ll.includes('шт.') &&
                        !ll.includes(' шт') &&
                        l.length > 10
                    ) {
                        title = l;
                        break;
                    }
                }
                if (!title && !isNoiseTitle(linkText)) title = linkText;
                if (!title || title.length < 6) continue;

                let rating = 4.7;
                let reviews = 0;
                for (const line of lines) {
                    const rm = line.match(/^([345][.,]\\d)\\s*(\\d[\\d\\s]*)?/);
                    if (rm) {
                        rating = parseFloat(rm[1].replace(',', '.'));
                        if (rm[2]) {
                            reviews = parseInt(rm[2].replace(/\\D/g, ''), 10) || 0;
                        }
                        break;
                    }
                }

                let seller = 'Маркетплейс Ozon';
                for (const line of lines) {
                    if (line.startsWith('Продавец ') || line.includes('Магазин') || line.includes('Ozon')) {
                        if (line.length < 45 && !line.includes('Карт')) {
                            seller = line;
                            break;
                        }
                    }
                }

                seen.add(baseHref);
                items.push({
                    title,
                    price,
                    seller,
                    rating,
                    reviews,
                    description: lines.join(' | '),
                    url: href,
                    site_duty: site_duty
                });
                if (items.length >= 120) break;
            }
            if (items.length >= 120) break;
            window.scrollBy(0, 900);
            await new Promise(r => setTimeout(r, 380));
        }
        return items;
    })()
    """.replace("__DUTY_HELPER__", _JS_DUTY_HELPER)

    def _collect_ozon(raw_list) -> List[ProductItem]:
        out: List[ProductItem] = []
        red_out: List[ProductItem] = []
        if not isinstance(raw_list, list):
            return out
        for r in raw_list:
            price = int(r.get("price") or 0)
            if price < min_price or (max_price > 0 and price > max_price):
                continue
            clean_title = clean_ozon_card_title(r.get("title", ""), r.get("description", ""))
            if not clean_title:
                continue
            site_duty = int(r.get("site_duty") or 0)
            prod = _build_product(
                shop="Ozon",
                query=query,
                title=clean_title,
                price=price,
                seller_name=r.get("seller", "Маркетплейс Ozon"),
                seller_rating=float(r.get("rating") or 0.0),
                reviews_count=int(r.get("reviews") or 0),
                location_delivery=f"{city_name} (Доставка Ozon)",
                description=r.get("description", ""),
                url=r.get("url", ""),
                site_duty=site_duty,
            )
            if prod.status_code == "RED":
                if len(red_out) < 15:
                    red_out.append(prod)
            else:
                out.append(prod)
                if len(out) >= limit:
                    break
        return out + red_out

    raw_items = await engine.fetch_page_js(
        url,
        js_extractor,
        wait_sec=5.5,
        scroll=False,
        ready_selector='a[href*="/product/"]',
        timeout_sec=30.0,
    )
    results = _collect_ozon(raw_items)
    if not results:
        diag = getattr(engine, "last_page_diag", None) or {}
        d_title = str(diag.get("title", "") or "")
        d_head = str(diag.get("head", "") or "")
        print(
            f"[Ozon] Диагностика: url={str(diag.get('url', ''))[:70]!r}, ready={diag.get('ready', '')!r}, "
            f"title={d_title[:60]!r}, links={diag.get('links', 0)}, head={d_head[:100]!r}",
            flush=True,
        )
        d_mix = f"{d_title} {d_head}".lower()
        if hasattr(engine, "clear_domain_cookies") and (
            not d_title
            or any(k in d_mix for k in ("antibot", "доступ ограничен", "нет соединения", "проверка", "vpn", "incident", "diag_error"))
        ):
            await engine.clear_domain_cookies("ozon")
        fallback_url = (
            f"https://www.ozon.ru/search/?text={encoded_q}"
            f"&sorting=price&currency_price={min_price}.000%3B{max_price}.000&from_global=true"
        )
        raw_items2 = await engine.fetch_page_js(
            fallback_url,
            js_extractor,
            wait_sec=5.5,
            scroll=False,
            ready_selector='a[href*="/product/"]',
            timeout_sec=30.0,
        )
        results = _collect_ozon(raw_items2)
    if not results:
        diag2 = getattr(engine, "last_page_diag", None) or {}
        if hasattr(engine, "clear_domain_cookies") and (
            not diag2.get("title")
            or any(k in f"{diag2.get('title','')} {diag2.get('head','')}".lower() for k in ("antibot", "доступ ограничен", "нет соединения", "vpn"))
        ):
            await engine.clear_domain_cookies("ozon")
        fallback_url2 = f"https://www.ozon.ru/search/?text={encoded_q}&sorting=price&from_global=true"
        raw_items3 = await engine.fetch_page_js(
            fallback_url2,
            js_extractor,
            wait_sec=5.5,
            scroll=False,
            ready_selector='a[href*="/product/"]',
            timeout_sec=30.0,
        )
        results = _collect_ozon(raw_items3)

    results.sort(key=lambda x: x.price)
    return results[:limit]


def _extract_distance_km(geo_text: str) -> float | None:
    if not geo_text:
        return None
    km_m = re.search(r"(\d+(?:[.,]\d+)?)\s*км\b", geo_text, flags=re.IGNORECASE)
    if km_m:
        try:
            return float(km_m.group(1).replace(",", "."))
        except Exception:
            return None
    m_m = re.search(r"(\d{2,4})\s*м(?:\s|$|,|\.)", geo_text, flags=re.IGNORECASE)
    if m_m:
        try:
            return round(float(m_m.group(1)) / 1000.0, 2)
        except Exception:
            return None
    return None


def _is_generic_all_russia(text: str) -> bool:
    t = re.sub(r"\s+", " ", (text or "").strip().lower())
    return t in ("", "все регионы рф", "вся россия", "россия", "рф", "все", "all")


def _format_avito_location(
    geo_str: str,
    city_name: str,
    scope_label: str,
    radius_km: int,
    has_delivery: bool = False,
    avito_mode: str = "both",
) -> tuple[str, bool]:
    """
    Классифицирует и форматирует локацию Авито:
    - 'PICKUP': строгий самовывоз в радиусе 5–15 км от города/точки пользователя;
    - 'DELIVERY': доставка в выбранный город (без фейкового «из: Все регионы РФ → Все»).
    Возвращает (formatted_location, accepted).
    """
    raw_geo = re.sub(r"\s+", " ", (geo_str or "").strip())
    if _is_generic_all_russia(raw_geo):
        raw_geo = ""

    dist_km = _extract_distance_km(raw_geo)
    city_clean = (city_name or "Все регионы РФ").strip()
    is_specific_city = not _is_generic_all_russia(city_clean)
    city_stem = re.sub(r"[^а-яa-z0-9]", "", city_clean.lower().split("-")[0].split()[0])[:5] if is_specific_city else ""
    geo_low = raw_geo.lower()
    in_selected_city = bool(city_stem and city_stem in re.sub(r"[^а-яa-z0-9]", "", geo_low))
    looks_local_street = bool(
        re.search(r"(?:р-н|район|ул\.|улица|пр-т|проспект|пер\.|бул\.|шоссе|метро|м\.\s*[а-я]|км\b|\d+\s*м\b)", geo_low)
    )

    target_city = city_clean if is_specific_city else "ваш город"

    if scope_label == "PICKUP":
        if dist_km is not None and dist_km > radius_km:
            if avito_mode != "pickup" and has_delivery:
                orig = raw_geo or "город продавца"
                return f"🚚 Авито Доставка: {orig} → {target_city}", True
            return "", False
        if is_specific_city and raw_geo and not in_selected_city and dist_km is None and not looks_local_street:
            # Карточка из рекомендаций другого города (не в пределах выбранного города)
            if avito_mode != "pickup" and has_delivery:
                return f"🚚 Авито Доставка: {raw_geo} → {target_city}", True
            return "", False
        if avito_mode == "delivery":
            if has_delivery:
                orig = raw_geo or (city_clean if is_specific_city else "город продавца")
                return f"🚚 Авито Доставка: {orig} → {target_city}", True
            return "", False
        dist_txt = f"{dist_km:g} км" if dist_km is not None else f"до {radius_km} км"
        place_txt = raw_geo or (city_clean if is_specific_city else "рядом")
        if is_specific_city and not in_selected_city and place_txt != city_clean:
            place_txt = f"{city_clean}, {place_txt}"
        return f"📍 Самовывоз ({dist_txt}) | {place_txt}", True

    # scope_label == "DELIVERY"
    if avito_mode == "pickup":
        return "", False
    orig_txt = raw_geo if raw_geo else "город продавца (РФ)"
    if is_specific_city and in_selected_city and (dist_km is None or dist_km <= radius_km):
        return f"🚚 Доставка (или самовывоз в г. {city_clean}) | {orig_txt}", True
    return f"🚚 Авито Доставка: {orig_txt} → {target_city}", True


async def parse_avito(
    engine: IsolatedChromeEngine,
    query: str,
    min_price: int,
    max_price: int,
    region_slug: str = "all",
    delivery_only: bool = True,
    all_russia_with_delivery: bool = True,
    limit: int = 100,
    city_name: str = "Все регионы РФ",
    radius_km: int = 15,
    lat: float | None = None,
    lon: float | None = None,
    avito_mode: str = "both",
) -> List[ProductItem]:
    """
    Авито (самовывоз 5–15 км + Авито Доставка):
    1) 'PICKUP': объявления строго в радиусе 5–15 км от точки/города пользователя.
    2) 'DELIVERY': объявления с Авито Доставкой (d=1) с реальным городом отправления → город пользователя.
    Режим avito_mode: 'both' | 'pickup' | 'delivery'.
    """
    min_price = _effective_min_price(query, min_price)
    encoded_q = urllib.parse.quote(query)
    loc_slug = (region_slug or "all").strip()
    cat = "/noutbuki" if "ноутбук" in query.lower() else ""
    price_q = f"&pmin={min_price}&pmax={max_price}"
    radius_km = max(5, min(int(radius_km or 15), 15))
    mode = (avito_mode or "both").strip().lower()
    if mode not in ("both", "pickup", "delivery"):
        mode = "both"

    urls_to_check = []
    if loc_slug != "all" and mode in ("both", "pickup"):
        eff_lat, eff_lon = get_city_coords(city_name, lat, lon)
        geo_q = f"&radius={radius_km}&searchRadius={radius_km}&geoCoords={eff_lat:.6f}%2C{eff_lon:.6f}"
        urls_to_check.append(
            ("PICKUP", f"https://www.avito.ru/{loc_slug}{cat}?q={encoded_q}&s=1{price_q}{geo_q}")
        )
    if mode in ("both", "delivery") and (all_russia_with_delivery or mode == "delivery" or loc_slug == "all"):
        urls_to_check.append(
            ("DELIVERY", f"https://www.avito.ru/all{cat}?q={encoded_q}&s=1&d=1{price_q}")
        )
        if cat:
            urls_to_check.append(
                ("DELIVERY", f"https://www.avito.ru/all?q={encoded_q}&s=1&d=1{price_q}")
            )
    if not urls_to_check and loc_slug == "all" and mode == "pickup":
        eff_lat, eff_lon = get_city_coords(city_name, lat, lon)
        urls_to_check.append(
            ("PICKUP", f"https://www.avito.ru/moskva{cat}?q={encoded_q}&s=1{price_q}&radius={radius_km}&searchRadius={radius_km}&geoCoords={eff_lat:.6f}%2C{eff_lon:.6f}")
        )

    js_extractor = """
    (async () => {
        for (let w = 0; w < 15; w++) {
            if (document.querySelector('[data-marker="item"], article[data-item-id]')) break;
            await new Promise(r => setTimeout(r, 300));
        }
        const items = [];
        const seen = new Set();

        for (let step = 0; step < 8; step++) {
            let cards = Array.from(document.querySelectorAll('[data-marker="item"], article[data-item-id]'));

            for (const card of cards) {
                const titleEl = card.querySelector('[itemprop="name"], [data-marker="item-title"], h3, h2');
                const linkEl = card.querySelector('a[itemprop="url"], a[data-marker="item-title"], a[href*="_"]') ||
                               Array.from(card.querySelectorAll('a[href]')).find(a => /_\\d{5,}/.test(a.href || ''));
                const priceMeta = card.querySelector('meta[itemprop="price"]');
                const priceTextEl = card.querySelector('[data-marker="item-price"]');

                const title = titleEl ? titleEl.innerText.trim() : (linkEl ? linkEl.innerText.trim() : '');
                if (!title) continue;

                const href = linkEl ? linkEl.href.split('?')[0] : '';
                const baseHref = href;
                if (!href || seen.has(href)) continue;

                let price = 0;
                if (priceMeta && priceMeta.content) {
                    price = parseInt(priceMeta.content.replace(/\\D/g, ''), 10) || 0;
                } else if (priceTextEl) {
                    price = parseInt(priceTextEl.innerText.replace(/\\D/g, ''), 10) || 0;
                } else {
                    const m = (card.innerText || '').replace(/[\\s\\u00a0]+/g, '').match(/(\\d{2,7})₽/);
                    if (m) price = parseInt(m[1], 10);
                }
                if (!price) continue;

                const descMeta = card.querySelector('meta[itemprop="description"]');
                const fullCardText = (card.innerText || '').replace(/\\n+/g, ' | ');

                const description = ((descMeta ? descMeta.content : '') + ' | ' + fullCardText).trim();

                const sellerEl = card.querySelector('[data-marker*="seller-info"] a, [class*="seller-info"] p, a[href*="/brands/"], a[href*="/user/"]');
                const seller = sellerEl ? sellerEl.innerText.trim().split('\\n')[0] : 'Продавец Авито';

                let rating = 0.0;
                let reviews = 0;
                const rMatch = (card.innerText || '').match(/([345][.,]\\d)\\s*\\n?\\s*(\\d+)\\s*отзыв/i);
                if (rMatch) {
                    rating = parseFloat(rMatch[1].replace(',', '.'));
                    reviews = parseInt(rMatch[2], 10) || 0;
                }

                const geoSelectors = [
                    '[data-marker="item-address"]',
                    '[data-marker*="geo"]',
                    '[class*="geo-root"]',
                    '[class*="geo-address"]',
                    '[class*="address-root"]',
                    '[itemprop="availableAtOrFrom"]'
                ];
                let geo = '';
                for (const sel of geoSelectors) {
                    const el = card.querySelector(sel);
                    if (el && el.innerText && el.innerText.trim()) {
                        const cand = el.innerText.replace(/\\n+/g, ', ').replace(/\\s+/g, ' ').trim();
                        if (cand) { geo = cand; break; }
                    }
                }
                if (!geo) {
                    const lines = (card.innerText || '').split('\\n').map(s => s.trim()).filter(Boolean);
                    for (let li = lines.length - 1; li >= 0; li--) {
                        const ln = lines[li];
                        if (ln.length > 2 && ln.length <= 75 && !/₽|отзыв|сегодня|вчера|час|мин|назад|компани/i.test(ln)) {
                            if (/(?:км\\b|м\\b|р-н|район|ул\\.|пр-т|шоссе|г\\.|город)/i.test(ln) || /^[А-ЯЁ][а-яё-]+(?:[\\s,-]+[А-ЯЁа-яё0-9.-]+){0,3}$/.test(ln)) {
                                geo = ln;
                                break;
                            }
                        }
                    }
                }
                const hasDelivery = !!(card.querySelector('[data-marker*="delivery"], [class*="delivery"]') || /авито доставк|с доставк/i.test(card.innerText || ''));

                seen.add(baseHref);
                items.push({
                    title,
                    price,
                    seller,
                    rating,
                    reviews,
                    geo,
                    has_delivery: hasDelivery,
                    description,
                    url: href
                });
                if (items.length >= 120) break;
            }
            if (items.length >= 120) break;
            window.scrollBy(0, 950);
            await new Promise(r => setTimeout(r, 380));
        }
        return items;
    })()
    """.replace("__DUTY_HELPER__", _JS_DUTY_HELPER)

    results: List[ProductItem] = []
    seen_urls = set()
    direct_engine: IsolatedChromeEngine | None = None
    active_engine = engine

    async def _try_avito_url(eng: IsolatedChromeEngine, target_url: str):
        raw = await eng.fetch_page_js(
            target_url,
            js_extractor,
            wait_sec=5.0,
            scroll=False,
            ready_selector='[data-marker="item"], article[data-item-id]',
            timeout_sec=26.0,
            retries=1,
        )
        diag = getattr(eng, "last_page_diag", None) or {}
        diag_txt = f"{diag.get('title', '')} {diag.get('head', '')}".lower()
        ip_block = any(m in diag_txt for m in ("доступ ограничен", "проблема с ip", "access denied"))
        blocked = ip_block or "timeouterror" in diag_txt or "diag_error" in diag_txt
        return raw, diag, ip_block, blocked

    try:
        for scope_label, url in urls_to_check:
            if scope_label == "DELIVERY" and cat and cat not in url and len([x for x in results if x.location_delivery.startswith("🚚")]) >= 3:
                continue
            t0 = time.monotonic()
            raw_items, diag, is_ip_block, is_blocked = await _try_avito_url(active_engine, url)
            elapsed = time.monotonic() - t0

            if not raw_items and not is_blocked and elapsed < 18.0:
                await asyncio.sleep(1.0)
                raw_items, diag, is_ip_block, is_blocked = await _try_avito_url(active_engine, url)

            if not raw_items:
                if getattr(engine, "bridge", None) and hasattr(engine.bridge, "mark_direct_suffixes"):
                    engine.bridge.mark_direct_suffixes("avito.ru", "avito.st")
                if hasattr(active_engine, "clear_domain_cookies"):
                    await active_engine.clear_domain_cookies("avito")
                if is_blocked:
                    print(f"[Авито] Перевод avito.ru/avito.st в прямой режим и пробуем m.avito.ru ({scope_label})", flush=True)
                    mob_url = url.replace("https://www.avito.ru/", "https://m.avito.ru/")
                    raw_items, diag, is_ip_block, is_blocked = await _try_avito_url(active_engine, mob_url)

            if not raw_items and not getattr(active_engine, "force_no_proxy", False):
                if direct_engine is None and (getattr(engine, "bridge", None) is not None or is_blocked):
                    print(f"[Авито] Резервный прямой запуск Chrome без прокси ({scope_label})", flush=True)
                    direct_engine = IsolatedChromeEngine(
                        debug_port=engine.debug_port + 2,
                        disable_images=engine.disable_images,
                        force_no_proxy=True,
                    )
                    if getattr(engine, "_region_coords", None) is not None:
                        try:
                            d_lat, d_lon = engine._region_coords
                            await direct_engine.apply_region_context(city_name=city_name, lat=d_lat, lon=d_lon)
                        except Exception:
                            pass
                if direct_engine is not None:
                    raw_items, diag, is_ip_block, is_blocked = await _try_avito_url(direct_engine, url)
                    if not raw_items:
                        mob_url = url.replace("https://www.avito.ru/", "https://m.avito.ru/")
                        raw_items, diag, is_ip_block, is_blocked = await _try_avito_url(direct_engine, mob_url)
                    if raw_items:
                        active_engine = direct_engine

            if not raw_items and diag:
                print(
                    f"[Авито] Диагностика ({scope_label}): title={str(diag.get('title', ''))[:50]!r}, "
                    f"head={str(diag.get('head', ''))[:90]!r}",
                    flush=True,
                )
            if not isinstance(raw_items, list):
                if is_ip_block and active_engine is direct_engine:
                    break
                continue
            if not raw_items:
                if is_ip_block and active_engine is direct_engine:
                    break
                continue

            red_count = 0
            for r in raw_items:
                u = r.get("url", "")
                if not u or u in seen_urls:
                    continue
                price = int(r.get("price") or 0)
                if price < min_price or (max_price > 0 and price > max_price):
                    continue

                loc_str, accepted = _format_avito_location(
                    geo_str=(r.get("geo", "") or ""),
                    city_name=city_name,
                    scope_label=scope_label,
                    radius_km=radius_km,
                    has_delivery=bool(r.get("has_delivery")),
                    avito_mode=mode,
                )
                if not accepted or not loc_str:
                    continue

                prod = _build_product(
                    shop="Авито",
                    query=query,
                    title=r.get("title", ""),
                    price=price,
                    seller_name=r.get("seller", "Продавец Авито"),
                    seller_rating=float(r.get("rating") or 0.0),
                    reviews_count=int(r.get("reviews") or 0),
                    location_delivery=loc_str,
                    description=r.get("description", ""),
                    url=u,
                )
                if prod.status_code == "RED":
                    if red_count >= 40:
                        continue
                    red_count += 1
                seen_urls.add(u)
                results.append(prod)
                if len([x for x in results if x.status_code != "RED"]) >= limit:
                    break
    finally:
        if direct_engine is not None:
            direct_engine.close()

    results.sort(key=lambda x: (0 if x.location_delivery.startswith("📍") else 1, x.price))
    return results[:limit]


async def parse_dns(
    engine: IsolatedChromeEngine,
    query: str,
    min_price: int,
    max_price: int,
    limit: int = 100,
    city_name: str = "Все регионы РФ",
) -> List[ProductItem]:
    """
    3. Приоритет №3: DNS (Все регионы РФ, сортировка по возрастанию цены).
    """
    await engine.warmup_url("https://www.dns-shop.ru/", wait_sec=5.5)
    min_price = _effective_min_price(query, min_price)
    encoded_q = urllib.parse.quote(query)
    url = (
        f"https://www.dns-shop.ru/search/?q={encoded_q}"
        f"&order=price-asc&f[price]={min_price}-{max_price}&price={min_price}-{max_price}"
    )

    js_extractor = """
    (async () => {
__DUTY_HELPER__
        const items = [];
        const seen = new Set();

        for (let step = 0; step < 8; step++) {
            const links = Array.from(document.querySelectorAll('a.catalog-product__name, a[href*="/product/"]'));

            for (const a of links) {
                const href = a.href || '';
                const baseHref = href.split('?')[0];
                if (!baseHref || seen.has(baseHref)) continue;

                let card = a.closest('[data-id="product"], .catalog-product, .markdown-product');
                if (!card) {
                    card = a;
                    for (let i = 0; i < 6; i++) {
                        if (!card.parentElement) break;
                        card = card.parentElement;
                        if ((card.innerText || '').includes('₽')) break;
                    }
                }
                if (!card) continue;

                const rawText = (card.innerText || '').replace(/\\n\\s*₽/g, ' ₽').trim();
                const title = (a.innerText || '').trim();
                if (title.length < 5) continue;

                let price = 0;
                const priceEl = card.querySelector('.product-buy__price');
                if (priceEl) {
                    const pLine = (priceEl.childNodes[0] ? priceEl.childNodes[0].textContent : priceEl.innerText || '').trim();
                    const lowP = (priceEl.innerText || '').toLowerCase();
                    if (!/\\/\\s*мес/i.test(lowP) && !/(^|\\s)в\\s+мес($|\\s|\\.)/i.test(lowP) && !lowP.includes('рассроч') && !lowP.includes('кредит')) {
                        const pm = pLine.replace(/[\\s\\u00a0\\u2006\\u2009\\u202f]+/g, '').match(/(\\d{2,7})₽/);
                        if (pm) price = parseInt(pm[1], 10);
                    }
                }
                if (!price) {
                    for (const l of rawText.split('\\n').map(s => s.trim()).filter(Boolean)) {
                        const ll = l.toLowerCase();
                        if (
                            /\\/\\s*мес/i.test(ll) || /(^|\\s)в\\s+мес($|\\s|\\.)/i.test(ll) || ll.includes('мес.') || ll.includes('месяц') ||
                            ll.includes('рассроч') || ll.includes('кредит') ||
                            ll.includes('бонус') || ll.includes('кешбэк') || ll.includes('кэшбэк') || ll.includes('выгод')
                        ) continue;
                        if (l.includes('₽')) {
                            const cleanText = l.replace(/[\\s\\u00a0\\u2006\\u2009\\u202f]+/g, '');
                            const m = cleanText.match(/(\\d{2,7})₽/);
                            if (m) {
                                price = parseInt(m[1], 10);
                                break;
                            }
                        }
                    }
                }
                if (!price) continue;

                let rating = 4.8;
                let reviews = 0;
                const ratingEl = card.querySelector('.catalog-product__rating, [data-rating]');
                if (ratingEl) {
                    const rm = ratingEl.innerText.match(/([345][.,]\\d)/);
                    if (rm) rating = parseFloat(rm[1].replace(',', '.'));
                    const revm = ratingEl.innerText.match(/\\b(\\d+)\\b/);
                    if (revm) reviews = parseInt(revm[1], 10);
                }

                seen.add(baseHref);
                items.push({
                    title,
                    price,
                    seller: 'DNS',
                    rating,
                    reviews,
                    description: rawText.replace(/\\n+/g, ' | '),
                    url: href
                });
                if (items.length >= 120) break;
            }
            if (items.length >= 120) break;
            window.scrollBy(0, 950);
            await new Promise(r => setTimeout(r, 380));
        }
        return items;
    })()
    """.replace("__DUTY_HELPER__", _JS_DUTY_HELPER)

    def _collect_dns(raw_list) -> List[ProductItem]:
        out: List[ProductItem] = []
        red_out: List[ProductItem] = []
        if not isinstance(raw_list, list):
            return out
        for r in raw_list:
            price = int(r.get("price") or 0)
            if price < min_price or (max_price > 0 and price > max_price):
                continue
            prod = _build_product(
                shop="DNS",
                query=query,
                title=r.get("title", ""),
                price=price,
                seller_name=f"DNS ({city_name})",
                seller_rating=float(r.get("rating") or 4.8),
                reviews_count=int(r.get("reviews") or 0),
                location_delivery=f"{city_name} (Магазины DNS)",
                description=r.get("description", ""),
                url=r.get("url", ""),
            )
            if prod.status_code == "RED":
                if len(red_out) < 15:
                    red_out.append(prod)
            else:
                out.append(prod)
                if len(out) >= limit:
                    break
        return out + red_out

    dns_sel = 'a.catalog-product__name, a[href*="/product/"]'
    raw_items = await engine.fetch_page_js(url, js_extractor, wait_sec=5.0, scroll=False, ready_selector=dns_sel, timeout_sec=30.0)
    results = _collect_dns(raw_items)
    if not results:
        alt_no_price = f"https://www.dns-shop.ru/search/?q={encoded_q}&order=price-asc"
        raw_items = await engine.fetch_page_js(alt_no_price, js_extractor, wait_sec=5.0, scroll=False, ready_selector=dns_sel, timeout_sec=30.0)
        results = _collect_dns(raw_items)
    if not results:
        clean_dns_q = re.sub(r"\b(ноутбук|смартфон|телефон|планшет|чехол|для|на)\w*\b", "", query.lower()).strip()
        if clean_dns_q and clean_dns_q != query.lower():
            alt_url = f"https://www.dns-shop.ru/search/?q={urllib.parse.quote(clean_dns_q)}&order=price-asc"
            raw_items = await engine.fetch_page_js(alt_url, js_extractor, wait_sec=5.0, scroll=False, ready_selector=dns_sel, timeout_sec=30.0)
            results = _collect_dns(raw_items)
    if not results:
        alnum_tokens = [t for t in re.findall(r"\b([a-zа-яё]*\d+[a-zа-яё0-9]*)\b", query.lower()) if len(t) >= 2]
        if alnum_tokens:
            best_tok = max(alnum_tokens, key=len)
            alt_url = f"https://www.dns-shop.ru/search/?q={urllib.parse.quote(best_tok)}&order=price-asc"
            raw_items = await engine.fetch_page_js(alt_url, js_extractor, wait_sec=5.0, scroll=False, ready_selector=dns_sel, timeout_sec=30.0)
            results = _collect_dns(raw_items)

    results.sort(key=lambda x: x.price)
    return results[:limit]


async def parse_wb(
    engine: IsolatedChromeEngine,
    query: str,
    min_price: int,
    max_price: int,
    limit: int = 100,
    city_name: str = "Все регионы РФ",
    lat: float | None = None,
    lon: float | None = None,
    wb_geo: dict | None = None,
) -> List[ProductItem]:
    """
    4. Приоритет №4: Wildberries (топ-100 самых дешёвых позиций с продавцом, рейтингом и пошлиной с карточки + регион ПВЗ).
    """
    await engine.ensure_started()
    min_price = _effective_min_price(query, min_price)
    results: List[ProductItem] = []
    encoded_q = urllib.parse.quote(query)
    price_u = f"{min_price * 100};{max_price * 100}"
    cat_filter = "&xsubject=2290" if "ноутбук" in query.lower() else ""

    if wb_geo is None:
        g_lat, g_lon = get_city_coords(city_name, lat, lon)
        wb_geo = await fetch_wb_geo_payload(g_lat, g_lon)
    dest_str = _extract_wb_dest(wb_geo)
    dest_q = f"&dest={urllib.parse.quote(dest_str)}" if dest_str else ""

    if wb_geo and getattr(engine, "_wb_region_applied", "") != city_name:
        geo_payload_json = json.dumps({"time": int(time.time() * 1000), "data": wb_geo}, ensure_ascii=False)
        ls_js = f"""
        (() => {{
            try {{
                localStorage.setItem('geo-data-v1-0', {json.dumps(geo_payload_json, ensure_ascii=False)});
                return true;
            }} catch(e) {{ return false; }}
        }})()
        """
        try:
            await engine.fetch_page_js("https://www.wildberries.ru/", ls_js, wait_sec=2.0, scroll=False, timeout_sec=16.0, retries=1)
            engine._wb_region_applied = city_name
        except Exception:
            pass

    wb_web_url = (
        f"https://www.wildberries.ru/catalog/0/search.aspx?page=1&sort=priceup"
        f"&search={encoded_q}&priceU={price_u}{cat_filter}{dest_q}"
    )
    js_wb = """
    (async () => {
__DUTY_HELPER__
        const items = [];
        const seen = new Set();
        for (let step = 0; step < 9; step++) {
            const cards = Array.from(document.querySelectorAll('article, .product-card'));

            for (const c of cards) {
                const a = c.querySelector('a.product-card__link, a[href*="/catalog/"]');
                if (!a) continue;
                const href = a.href.split('?')[0];
                const baseHref = href;
                if (!href || seen.has(href)) continue;

                const rawText = c.innerText || '';
                if (!rawText.includes('₽') || rawText.includes('₸')) continue;

                const lowText = rawText.toLowerCase();
                if (lowText.includes('похожие')) {
                    continue;
                }

                const lines = rawText.replace(/\\n\\s*₽/g, ' ₽').split('\\n').map(s => s.trim()).filter(Boolean);
                const flatBlock = lines.join(' ');
                let price = 0;
                let site_duty = 0;
                const flatDv = _extractDutyRub(flatBlock);
                if (flatDv > 0) site_duty = flatDv;
                const wbPrices = [];
                for (const l of lines) {
                    const ll = l.toLowerCase();
                    const dv = _extractDutyRub(l);
                    if (dv > 0) {
                        site_duty = dv;
                        continue;
                    }
                    if (
                        ll.includes('пошлин') ||
                        ll.includes('тамож') ||
                        /\\/\\s*мес/i.test(ll) || /(^|\\s)в\\s+мес($|\\s|\\.)/i.test(ll) || ll.includes('мес.') || ll.includes('месяц') ||
                        ll.includes('рассроч') ||
                        ll.includes('кредит') ||
                        ll.includes('сплит') ||
                        ll.includes('долями') ||
                        ll.includes('балл') ||
                        ll.includes('бонус') ||
                        ll.includes('кешбэк') ||
                        ll.includes('кэшбэк') ||
                        ll.includes('выгод') ||
                        ll.includes('купон') ||
                        /[×xх]\\s*(2|4|6|10|12)\\b/i.test(ll)
                    ) {
                        continue;
                    }
                    if (l.includes('₽')) {
                        const cleanText = l.replace(/[\\s\\u00a0\\u2006\\u2009\\u202f]+/g, '');
                        const m = cleanText.match(/(\\d{2,7})₽/);
                        if (m) {
                            wbPrices.push(parseInt(m[1], 10));
                        }
                    }
                }
                if (wbPrices.length > 0) {
                    price = Math.min(...wbPrices);
                }
                if (!price) continue;

                let title = '';
                const wrapEl = c.querySelector('.product-card__brand-wrap, .product-card__name');
                const nameOnlyEl = c.querySelector('.product-card__name');
                if (wrapEl && wrapEl.innerText.trim().length > 5) {
                    title = wrapEl.innerText.replace(/^\\s*\\/\\s*/, '').replace(/\\s+/g, ' ').trim();
                    if (nameOnlyEl && !title.includes(nameOnlyEl.innerText.trim())) {
                        title = (title + ' ' + nameOnlyEl.innerText.replace(/^\\s*\\/\\s*/, '').trim()).trim();
                    }
                } else if (a.getAttribute('aria-label') && a.getAttribute('aria-label').trim().length > 5) {
                    title = a.getAttribute('aria-label').trim();
                } else {
                    for (const l of lines) {
                        const ll = l.toLowerCase();
                        if (
                            !l.includes('₽') &&
                            !ll.includes('кошельк') &&
                            !ll.includes('акция') &&
                            !ll.includes('распродаж') &&
                            !ll.includes('оценк') &&
                            !ll.includes('в корзину') &&
                            !l.match(/^[−+-]?\\d+%$/) &&
                            !l.match(/^\\d+\\s+(ноября|декабря|января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|завтра|послезавтра)/i) &&
                            l.length > 8
                        ) {
                            title = l.replace(/^\\s*\\/\\s*/, '');
                            break;
                        }
                    }
                }
                if (!title) continue;

                const brandEl = c.querySelector('.product-card__brand');
                const seller = brandEl ? brandEl.innerText.trim() : 'Продавец WB';

                const ratingEl = c.querySelector('.address-rate-mini');
                const rating = ratingEl ? (parseFloat(ratingEl.innerText.replace(',', '.')) || 4.7) : 4.7;

                seen.add(baseHref);
                items.push({
                    title,
                    price,
                    seller,
                    rating,
                    desc: lines.join(' | '),
                    url: href,
                    site_duty: site_duty
                });
                if (items.length >= 120) break;
            }
            if (items.length >= 120) break;
            window.scrollBy(0, 950);
            await new Promise(r => setTimeout(r, 380));
        }
        return items;
    })()
    """.replace("__DUTY_HELPER__", _JS_DUTY_HELPER)
    def _collect_wb(raw_list) -> List[ProductItem]:
        out: List[ProductItem] = []
        red_wb: List[ProductItem] = []
        if not isinstance(raw_list, list):
            return out
        for r in raw_list:
            price = int(r.get("price") or 0)
            if price < min_price or (max_price > 0 and price > max_price):
                continue
            prod = _build_product(
                shop="Wildberries",
                query=query,
                title=r.get("title", ""),
                price=price,
                seller_name=r.get("seller", "Продавец WB"),
                seller_rating=float(r.get("rating") or 4.7),
                reviews_count=0,
                location_delivery=f"{city_name} (ПВЗ Wildberries)",
                description=r.get("desc", ""),
                url=r.get("url", ""),
                site_duty=int(r.get("site_duty") or 0),
            )
            if prod.status_code == "RED":
                if len(red_wb) < 15:
                    red_wb.append(prod)
            else:
                out.append(prod)
        return out + red_wb

    wb_sel = "article, .product-card"
    raw = await engine.fetch_page_js(wb_web_url, js_wb, wait_sec=6.0, scroll=True, ready_selector=wb_sel, timeout_sec=30.0)
    results = _collect_wb(raw)
    if not results:
        alt_wb_url = f"https://www.wildberries.ru/catalog/0/search.aspx?page=1&sort=priceup&search={encoded_q}{dest_q}"
        raw2 = await engine.fetch_page_js(alt_wb_url, js_wb, wait_sec=6.0, scroll=True, ready_selector=wb_sel, timeout_sec=30.0)
        results = _collect_wb(raw2)

    results.sort(key=lambda x: x.price)
    return results[:limit]


async def parse_aliexpress(
    engine: IsolatedChromeEngine,
    query: str,
    min_price: int,
    max_price: int,
    limit: int = 100,
    city_name: str = "Все регионы РФ",
) -> List[ProductItem]:
    """
    5. Приоритет №5: AliExpress Россия (сортировка по возрастанию цены + отсев проданных/недоступных).
    """
    min_price = _effective_min_price(query, min_price)
    encoded_q = urllib.parse.quote(query)
    url = f"https://aliexpress.ru/wholesale?SearchText={encoded_q}&sortType=price_asc&minPrice={min_price}&maxPrice={max_price}"

    js_extractor = """
    (async () => {
__DUTY_HELPER__
        const items = [];
        const seen = new Set();

        for (let step = 0; step < 8; step++) {
            const links = Array.from(document.links).filter(a => a.href && a.href.includes('/item/'));

            for (const a of links) {
                const href = a.href || '';
                const baseHref = href.split('?')[0];
                if (!baseHref || seen.has(baseHref)) continue;

                let card = a.closest('[class*="RedSnippet"], [class*="SnowSearchProductFeed"], [class*="product-snippet"]') || a;

                const rawText = (card.innerText || a.innerText || '').trim();
                if (!rawText.includes('₽') || rawText.includes('₸') || rawText.length < 12) continue;

                const lines = rawText.replace(/\\n\\s*₽/g, ' ₽').split('\\n').map(s => s.trim()).filter(Boolean);
                const flatBlock = lines.join(' ');
                let site_duty = 0;
                const flatDv = _extractDutyRub(flatBlock);
                if (flatDv > 0) site_duty = flatDv;
                const prices = [];
                for (const l of lines) {
                    const low = l.toLowerCase();
                    const dv = _extractDutyRub(l);
                    if (dv > 0) {
                        site_duty = dv;
                        continue;
                    }
                    if (
                        low.includes('пошлин') ||
                        low.includes('тамож') ||
                        low.includes('купон') ||
                        low.includes('скидк') ||
                        low.includes('баллов') ||
                        low.includes('бонус') ||
                        low.includes('кешбэк') ||
                        low.includes('кэшбэк') ||
                        low.includes('доставк') ||
                        low.includes('выгод') ||
                        /\\/\\s*мес/i.test(low) || /(^|\\s)в\\s+мес($|\\s|\\.)/i.test(low) || low.includes('мес.') || low.includes('месяц') ||
                        low.includes('рассроч') ||
                        low.includes('кредит') ||
                        low.includes('сплит') ||
                        low.includes('долями') ||
                        /[×xх]\\s*(2|4|6|10|12)\\b/i.test(low)
                    ) {
                        continue;
                    }
                    if (l.includes('₽')) {
                        const clean = l.replace(/[\\s\\u00a0\\u2006\\u2009\\u202f]+/g, '');
                        const m = clean.match(/(\\d{2,7})(?:[.,]\\d+)?₽/);
                        if (m) {
                            const pVal = parseInt(m[1], 10);
                            if (pVal >= 50) {
                                prices.push(pVal);
                            }
                        }
                    }
                }
                if (!prices.length) continue;
                let price = prices[0];
                if (prices.length > 1) {
                    const maxP = Math.max(...prices);
                    const filtered = prices.filter(p => p >= maxP * 0.35);
                    price = filtered.length ? Math.min(...filtered) : prices[0];
                }

                let title = '';
                for (const l of lines) {
                    if (
                        !l.includes('₽') &&
                        !l.includes('купили') &&
                        !l.match(/^[−+-]?\\d+%$/) &&
                        !l.match(/^[0-9.,]+$/) &&
                        l.length > 10
                    ) {
                        title = l;
                        break;
                    }
                }
                if (!title) {
                    title = (a.getAttribute('title') || a.innerText || '').trim();
                }
                if (!title || title.length < 5) continue;

                let rating = 4.8;
                let reviews = 0;
                const rm = rawText.match(/\\b([45][.,]\\d)\\b/);
                if (rm) rating = parseFloat(rm[1].replace(',', '.'));
                const boughtMatch = rawText.match(/(\\d[\\d\\s]*)\\s*купили/);
                if (boughtMatch) reviews = parseInt(boughtMatch[1].replace(/\\D/g, ''), 10) || 0;

                seen.add(baseHref);
                items.push({
                    title,
                    price,
                    seller: 'Магазин AliExpress',
                    rating,
                    reviews,
                    description: lines.join(' | '),
                    url: href,
                    site_duty: site_duty
                });
                if (items.length >= 120) break;
            }
            if (items.length >= 120) break;
            window.scrollBy(0, 950);
            await new Promise(r => setTimeout(r, 380));
        }
        return items;
    })()
    """.replace("__DUTY_HELPER__", _JS_DUTY_HELPER)

    def _collect_ali(raw_list) -> List[ProductItem]:
        out: List[ProductItem] = []
        red_out: List[ProductItem] = []
        if not isinstance(raw_list, list):
            return out
        for r in raw_list:
            price = int(r.get("price") or 0)
            if price < min_price or (max_price > 0 and price > max_price):
                continue
            site_duty = int(r.get("site_duty") or 0)
            prod = _build_product(
                shop="AliExpress",
                query=query,
                title=r.get("title", ""),
                price=price,
                seller_name=r.get("seller", "Магазин AliExpress"),
                seller_rating=float(r.get("rating") or 4.8),
                reviews_count=int(r.get("reviews") or 0),
                location_delivery=f"Доставка по РФ ({city_name}) | 🔥 Есть дешёвая скидка!",
                description=f"🔥 Есть дешёвая скидка на AliExpress! | {r.get('description', '')}",
                url=r.get("url", ""),
                site_duty=site_duty,
            )
            if prod.status_code == "RED":
                if len(red_out) < 15:
                    red_out.append(prod)
            else:
                out.append(prod)
                if len(out) >= limit:
                    break
        return out + red_out

    ali_sel = 'a[href*="/item/"]'
    raw_items = await engine.fetch_page_js(url, js_extractor, wait_sec=5.0, scroll=True, ready_selector=ali_sel, timeout_sec=30.0)
    results = _collect_ali(raw_items)
    if not results:
        alt_no_price = f"https://aliexpress.ru/wholesale?SearchText={encoded_q}&sortType=price_asc"
        raw_items = await engine.fetch_page_js(alt_no_price, js_extractor, wait_sec=5.0, scroll=True, ready_selector=ali_sel, timeout_sec=30.0)
        results = _collect_ali(raw_items)
    if not results:
        clean_ali_q = re.sub(r"\b(ноутбук|смартфон|телефон|планшет|чехол|для|на)\w*\b", "", query.lower()).strip()
        if clean_ali_q and clean_ali_q != query.lower():
            alt_url = f"https://aliexpress.ru/wholesale?SearchText={urllib.parse.quote(clean_ali_q)}&sortType=price_asc"
            raw_items = await engine.fetch_page_js(alt_url, js_extractor, wait_sec=5.0, scroll=True, ready_selector=ali_sel, timeout_sec=30.0)
            results = _collect_ali(raw_items)
    if not results:
        alnum_tokens = [t for t in re.findall(r"\b([a-zа-яё]*\d+[a-zа-яё0-9]*)\b", query.lower()) if len(t) >= 2]
        if alnum_tokens:
            best_tok = max(alnum_tokens, key=len)
            alt_url = f"https://aliexpress.ru/wholesale?SearchText={urllib.parse.quote(best_tok)}&sortType=price_asc"
            raw_items = await engine.fetch_page_js(alt_url, js_extractor, wait_sec=5.0, scroll=True, ready_selector=ali_sel, timeout_sec=30.0)
            results = _collect_ali(raw_items)

    results.sort(key=lambda x: x.price)
    return results[:limit]


async def parse_ymarket(
    engine: IsolatedChromeEngine,
    query: str,
    min_price: int,
    max_price: int,
    limit: int = 100,
    city_name: str = "Все регионы РФ",
    lr: int = 213,
) -> List[ProductItem]:
    """
    6. Приоритет №6: Яндекс Маркет (выбранный город, сортировка по возрастанию цены + полные ссылки do-waremd5).
    """
    min_price = _effective_min_price(query, min_price)
    encoded_q = urllib.parse.quote(query)
    url = (
        f"https://market.yandex.ru/search?text={encoded_q}"
        f"&how=aprice&pricefrom={min_price}&priceto={max_price}&lr={lr}"
    )

    js_extractor = """
    (async () => {
__DUTY_HELPER__
        const items = [];
        const seen = new Set();

        for (let step = 0; step < 8; step++) {
            const cards = Array.from(document.querySelectorAll('article, [data-auto="searchOrganic"], [data-autotest-id="product-snippet"]'));

            for (const c of cards) {
                const candidateLinks = Array.from(c.querySelectorAll('a[href*="/card/"], a[href*="/product--"], h3 a'));
                if (!candidateLinks.length) continue;
                // Выбираем ссылку оффера (do-waremd5 / sku / cpc)
                const a = candidateLinks.find(el => /do-waremd5=|sku=|cpc=|uniqueId=/i.test(el.href || '')) || candidateLinks[0];
                const fullHref = a.href || '';
                const baseHref = fullHref.split('?')[0];
                if (!fullHref || !baseHref || seen.has(baseHref)) continue;

                const rawText = (c.innerText || '').replace(/\\n\\s*₽/g, ' ₽').trim();
                if (!rawText.includes('₽') || rawText.includes('₸')) continue;

                const lines = rawText.split('\\n').map(s => s.trim()).filter(Boolean);
                const flatBlock = lines.join(' ');
                let price = 0;
                let site_duty = 0;
                const flatDv = _extractDutyRub(flatBlock);
                if (flatDv > 0) site_duty = flatDv;
                const ymPrices = [];
                for (const l of lines) {
                    const ll = l.toLowerCase();
                    const dv = _extractDutyRub(l);
                    if (dv > 0) {
                        site_duty = dv;
                        continue;
                    }
                    if (
                        ll.includes('пошлин') ||
                        ll.includes('тамож') ||
                        ll.includes('сплит') ||
                        ll.includes('долями') ||
                        /\\/\\s*мес/i.test(ll) || /(^|\\s)в\\s+мес($|\\s|\\.)/i.test(ll) || ll.includes('мес.') || ll.includes('месяц') ||
                        ll.includes('рассроч') ||
                        ll.includes('кредит') ||
                        ll.includes('балл') ||
                        ll.includes('бонус') ||
                        ll.includes('плюс') ||
                        ll.includes('кешбэк') ||
                        ll.includes('кэшбэк') ||
                        ll.includes('выгод') ||
                        /[×xх]\\s*(2|4|6|10|12)\\b/i.test(ll)
                    ) {
                        continue;
                    }
                    if (l.includes('₽')) {
                        const clean = l.replace(/[\\s\\u00a0\\u2006\\u2009\\u202f]+/g, '');
                        const m = clean.match(/(\\d{2,7})₽/);
                        if (m) {
                            ymPrices.push(parseInt(m[1], 10));
                        }
                    }
                }
                if (ymPrices.length > 0) {
                    price = Math.min(...ymPrices);
                }
                if (!price) continue;

                let title = '';
                const h3 = c.querySelector('h3, [data-auto="snippet-title"]');
                if (h3 && h3.innerText.trim().length > 10) {
                    title = h3.innerText.replace(/\\s+/g, ' ').trim();
                } else {
                    for (const l of lines) {
                        if (
                            !l.includes('₽') &&
                            !l.toLowerCase().includes('оригинал') &&
                            !l.toLowerCase().includes('сплит') &&
                            l.length > 14
                        ) {
                            title = l;
                            break;
                        }
                    }
                }
                if (!title) continue;

                let rating = 4.8;
                const rm = rawText.match(/([345][.,]\\d)/);
                if (rm) rating = parseFloat(rm[1].replace(',', '.'));

                let seller = 'Яндекс Маркет';
                for (const l of lines) {
                    if (l.startsWith('Продавец ') || l.startsWith('Магазин ')) {
                        seller = l.replace(/^(Продавец|Магазин)\\s+/, '');
                        break;
                    }
                }

                seen.add(baseHref);
                items.push({
                    title,
                    price,
                    seller,
                    rating,
                    description: lines.join(' | '),
                    url: fullHref,
                    site_duty: site_duty
                });
                if (items.length >= 120) break;
            }
            if (items.length >= 120) break;
            window.scrollBy(0, 950);
            await new Promise(r => setTimeout(r, 380));
        }
        return items;
    })()
    """.replace("__DUTY_HELPER__", _JS_DUTY_HELPER)
    def _collect_ym(raw_list) -> tuple[List[ProductItem], List[ProductItem]]:
        ok_out: List[ProductItem] = []
        red_out: List[ProductItem] = []
        if not isinstance(raw_list, list):
            return ok_out, red_out
        for r in raw_list:
            price = int(r.get("price") or 0)
            if price < min_price or (max_price > 0 and price > max_price):
                continue
            prod = _build_product(
                shop="Яндекс Маркет",
                query=query,
                title=r.get("title", ""),
                price=price,
                seller_name=r.get("seller", "Яндекс Маркет"),
                seller_rating=float(r.get("rating") or 4.8),
                reviews_count=0,
                location_delivery=f"{city_name} (ПВЗ Яндекс Маркет)",
                description=r.get("description", ""),
                url=r.get("url", ""),
                site_duty=int(r.get("site_duty") or 0),
            )
            if prod.status_code == "RED":
                if len(red_out) < 15:
                    red_out.append(prod)
            else:
                ok_out.append(prod)
                if len(ok_out) >= limit:
                    break
        return ok_out, red_out

    ym_sel = 'article, [data-auto="searchOrganic"], [data-autotest-id="product-snippet"]'
    raw_items = await engine.fetch_page_js(url, js_extractor, wait_sec=5.0, scroll=True, ready_selector=ym_sel, timeout_sec=30.0)
    results, red_ym = _collect_ym(raw_items)
    if not results and not red_ym:
        alt_ym_url = f"https://market.yandex.ru/search?text={encoded_q}&how=aprice&lr={lr}"
        raw_items2 = await engine.fetch_page_js(alt_ym_url, js_extractor, wait_sec=5.0, scroll=True, ready_selector=ym_sel, timeout_sec=30.0)
        results, red_ym = _collect_ym(raw_items2)
    if not results and not red_ym:
        alnum_tokens = [t for t in re.findall(r"\b([a-zа-яё]*\d+[a-zа-яё0-9]*)\b", query.lower()) if len(t) >= 2]
        if alnum_tokens:
            best_tok = max(alnum_tokens, key=len)
            alt_ym_tok = f"https://market.yandex.ru/search?text={urllib.parse.quote(best_tok)}&how=aprice&lr={lr}"
            raw_items3 = await engine.fetch_page_js(alt_ym_tok, js_extractor, wait_sec=5.0, scroll=True, ready_selector=ym_sel, timeout_sec=30.0)
            results, red_ym = _collect_ym(raw_items3)

    results.sort(key=lambda x: x.price)
    return results[:limit] + red_ym
