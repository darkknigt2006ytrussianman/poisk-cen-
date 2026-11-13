import argparse
import asyncio
import html
import json
import math
import os
import re
import subprocess
import sys
import tempfile
import time
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

sys.dont_write_bytecode = True

if sys.platform == "win32":
    import msvcrt

import httpx
from rich.console import Console
from rich.table import Table

from chrome_engine import IsolatedChromeEngine
from excel_writer import SHOP_PRIORITY_ORDER, save_to_excel
from parsers import (
    ProductItem,
    _build_product,
    clean_product_url,
    fetch_wb_geo_payload,
    filter_category_outliers,
    get_city_coords,
    parse_aliexpress,
    parse_avito,
    parse_dns,
    parse_ozon,
    parse_wb,
    parse_ymarket,
    resolve_city_by_coords_async,
    resolve_city_settings,
)
from user_store import UserStore, get_cache_dir
from verifier import verify_all_cards, verify_top_cards

BASE_DIR = Path(__file__).parent
_CACHE_DIR = get_cache_dir()
CONFIG_PATH = BASE_DIR / "config.json"
HISTORY_PATH = _CACHE_DIR / "price_history.json"
VERSION_STATE_PATH = _CACHE_DIR / "poisk_bot_version.json"
EXCEL_PATH = BASE_DIR / "Мониторинг_Товаров.xlsx"
LOG_PATH = _CACHE_DIR / "bot_runtime.log"

# Версионирование X.Y.Z: +0.0.1 — мелкие правки/коррекции, +0.1.0 — заметный набор функций, +1.0.0 — глобальное обновление
BOT_VERSION = "2.8.2"
BOT_CHANGELOG = {
    "version": BOT_VERSION,
    "kind": "patch",  # patch (+0.0.1) | minor (+0.1.0) | major (+1.0.0)
    "added": [
        "Фильтр «📍 Самовывоз 5–15 км / 🚚 Доставка / 🌐 Всё» прямо под отчётом Авито и выбор радиуса (5 / 10 / 15 км) в меню города",
        "Шифрованный перенос настроек, выбранного города и таймера сканирования между перезапусками облачных смен",
    ],
    "changed": [
        "Интервал автообновления по умолчанию установлен строго каждые 3 часа (без самопроизвольных 5-минутных циклов или перехода на 1 день)",
        "Топ-лист, рейтинг «Цена / Качество» и Excel строго привязаны к выбранному пользователем региону и сегодняшней дате — старые/чужие данные не подмешиваются",
    ],
    "fixed": [
        "Убрана заглушка Авито «Доставка из: Все регионы РФ → Все» — теперь извлекается реальный город продавца и передаются координаты для точного самовывоза 5–15 км",
        "Исправлен старт облачной смены (не засыпает на старом закоммиченном Excel и не повторяет старые команды Telegram после рестарта)",
    ],
}

CHANGELOG_KIND_LABELS = {
    "patch": "🧩 Точечное обновление",
    "minor": "✨ Функциональное обновление",
    "major": "🌟 Глобальное обновление",
}


def format_update_news_tg() -> str:
    kind = str(BOT_CHANGELOG.get("kind") or "patch").strip().lower()
    kind_label = CHANGELOG_KIND_LABELS.get(kind, "🧩 Обновление")
    lines = [
        f"🚀 <b>Бот обновлён до версии {_h(BOT_VERSION)}</b>",
        f"<i>{_h(kind_label)}</i>\n",
    ]
    sections = (
        ("➕ <b>Добавлено:</b>", BOT_CHANGELOG.get("added", [])),
        ("🛠 <b>Скорректировано:</b>", BOT_CHANGELOG.get("changed", [])),
        ("🐛 <b>Исправлено:</b>", BOT_CHANGELOG.get("fixed", [])),
    )
    for hdr, items_list in sections:
        if items_list:
            lines.append(hdr)
            for point in items_list:
                lines.append(f"• {_h(point)}")
            lines.append("")
    lines.append("<i>Нажмите «🔍 Поиск новых цен» или отправьте /changelog, чтобы снова открыть список изменений.</i>")
    return "\n".join(lines).strip()


def _get_announced_version() -> str:
    meta_ver = str((rt.user_store.meta or {}).get("announced_version") or "").strip()
    if meta_ver:
        return meta_ver
    try:
        if VERSION_STATE_PATH.exists():
            data = json.loads(VERSION_STATE_PATH.read_text(encoding="utf-8"))
            return str(data.get("announced_version") or "").strip()
    except Exception:
        pass
    return ""


def _set_announced_version(ver: str) -> None:
    rt.user_store.meta["announced_version"] = ver
    rt.user_store.save()
    try:
        VERSION_STATE_PATH.write_text(
            json.dumps({"announced_version": ver, "updated_at": datetime.now().isoformat()}, ensure_ascii=False),
            encoding="utf-8",
        )
    except Exception:
        pass


def _collect_broadcast_recipients(cfg: dict) -> List[str]:
    rec: List[str] = []
    seen = set()
    tg_sec = cfg.get("telegram", {}) if isinstance(cfg, dict) else {}
    raw_candidates: List[str] = list(rt.active_chat_ids)
    init_cid = str(tg_sec.get("chat_id") or "").strip()
    if init_cid:
        raw_candidates.append(init_cid)
    for ac in (tg_sec.get("allowed_chat_ids") or []):
        if str(ac).strip():
            raw_candidates.append(str(ac).strip())
    for uk in (rt.user_store.users or {}).keys():
        s_uk = str(uk).strip()
        if s_uk.lstrip("-").isdigit():
            raw_candidates.append(s_uk)
    for c in raw_candidates:
        if c and c not in seen and c.lstrip("-").isdigit():
            seen.add(c)
            rec.append(c)
    return rec

if sys.stdout is None or sys.stderr is None:
    _log_fh = open(LOG_PATH, "a", encoding="utf-8", buffering=1)
    sys.stdout = _log_fh
    sys.stderr = _log_fh
    console = Console(file=_log_fh, force_terminal=False)
else:
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace", write_through=True)
        sys.stderr.reconfigure(encoding="utf-8", errors="replace", write_through=True)
    except Exception:
        pass
    console = Console()


class RuntimeState:
    def __init__(self):
        self.items: List[ProductItem] = []
        self.is_scanning: bool = False
        self.current_step: str = "Готов к работе"
        self.last_updated: str = "—"
        self.scan_event = asyncio.Event()
        self.current_scan_task: Optional[asyncio.Task] = None
        self.last_update_id: int = 0
        self.tg_dialog_mode: Dict[str, str] = {}
        self.active_chat_ids: set[str] = set()
        self.user_store = UserStore()
        self.user_items: Dict[str, List[ProductItem]] = {}
        self.user_scan_meta: Dict[str, dict] = {}
        self.global_scan_meta: dict = {}
        self.user_excel_paths: Dict[str, Path] = {}
        self.avito_view_filter: Dict[str, str] = {}
        self.scan_queue: List[str] = []


rt = RuntimeState()

SUBSCRIBERS_PATH = _CACHE_DIR / "tg_bot_subscribers.json"


def _today_str() -> str:
    return datetime.now().strftime("%d.%m.%Y")


def _norm_reg(val: object) -> str:
    return re.sub(r"\s+", " ", str(val or "Все регионы РФ").strip().lower())


def _build_scan_meta(
    city_name: str,
    queries: List[str],
    cid: str = "",
    avito_mode: str = "both",
    avito_radius_km: int = 15,
    ts: Optional[float] = None,
) -> dict:
    return {
        "region": str(city_name or "Все регионы РФ").strip(),
        "queries": [str(q).strip() for q in (queries or []) if str(q).strip()],
        "date": _today_str(),
        "cid": str(cid or ""),
        "avito_mode": str(avito_mode or "both"),
        "avito_radius_km": int(avito_radius_km or 15),
        "ts": float(ts if ts is not None else time.time()),
    }


def _invalidate_user_cache(cid: str, clear_global: bool = False) -> None:
    s_cid = str(cid or "")
    if s_cid:
        rt.user_items.pop(s_cid, None)
        rt.user_scan_meta.pop(s_cid, None)
        rt.user_excel_paths.pop(s_cid, None)
        rt.avito_view_filter.pop(s_cid, None)
    if clear_global or (s_cid and rt.global_scan_meta.get("cid") == s_cid) or len(rt.user_store.users) <= 1:
        rt.items = []
        rt.global_scan_meta = {}


def _meta_matches_user(
    meta: Optional[dict],
    u_cfg: Optional[dict],
    items_list: Optional[List[ProductItem]] = None,
    require_today: bool = True,
) -> tuple[bool, str]:
    today_s = _today_str()
    cfg_eff = u_cfg or {}
    want_reg = _norm_reg(cfg_eff.get("region", "Все регионы РФ"))
    want_queries = [q.strip().lower() for q in _active_queries(cfg_eff) if q.strip()]

    if meta:
        meta_date = str(meta.get("date") or "").strip()
        if require_today and meta_date and meta_date != today_s:
            return False, "stale_date"
        meta_reg = _norm_reg(meta.get("region", ""))
        if meta_reg and meta_reg != want_reg:
            return False, "wrong_region"
        meta_q = [str(q).strip().lower() for q in (meta.get("queries") or []) if str(q).strip()]
        if want_queries and meta_q and set(meta_q) != set(want_queries):
            return False, "wrong_queries"
        want_mode = str(cfg_eff.get("avito_mode") or "both").strip().lower()
        meta_mode = str(meta.get("avito_mode") or "both").strip().lower()
        if meta.get("avito_mode") and meta_mode != want_mode:
            return False, "wrong_avito_mode"
        want_rad = int(cfg_eff.get("avito_radius_km") or 15)
        meta_rad = int(meta.get("avito_radius_km") or 15)
        if meta.get("avito_radius_km") and meta_rad != want_rad:
            return False, "wrong_avito_radius"
    elif require_today and items_list:
        dated = [str(x.updated_at or "").strip() for x in items_list if str(x.updated_at or "").strip()]
        if dated and not any(d.startswith(today_s) for d in dated):
            return False, "stale_date"

    return True, "ok"


def _get_scoped_items(cid: str = "", u_cfg: Optional[dict] = None) -> tuple[List[ProductItem], str]:
    s_cid = str(cid or "")
    eff_cfg = u_cfg
    if eff_cfg is None:
        if s_cid and rt.user_store.exists(s_cid):
            eff_cfg = rt.user_store.get(s_cid)
        else:
            eff_cfg = {}
    cur_reg = str((eff_cfg or {}).get("region") or "Все регионы РФ")
    today_s = _today_str()

    if s_cid and s_cid in rt.user_items:
        u_items = rt.user_items.get(s_cid) or []
        if u_items:
            ok, _ = _meta_matches_user(rt.user_scan_meta.get(s_cid), eff_cfg, u_items, require_today=True)
            if ok:
                return u_items, ""

    if rt.items and rt.global_scan_meta:
        ok, _ = _meta_matches_user(rt.global_scan_meta, eff_cfg, rt.items, require_today=True)
        if ok:
            return rt.items, ""

    if not s_cid and rt.items and not eff_cfg:
        ok, _ = _meta_matches_user(rt.global_scan_meta or None, {}, rt.items, require_today=True)
        if ok:
            return rt.items, ""

    if rt.is_scanning:
        return [], (
            f"⏳ <b>Данные за сегодня ({today_s}) для региона «{_h(cur_reg)}» прямо сейчас обновляются.</b>\n"
            f"Дождитесь окончания текущего поиска — актуальный список придёт автоматически."
        )
    return [], (
        f"⏳ <b>Для региона «{_h(cur_reg)}» ещё нет свежих данных за сегодня ({today_s}).</b>\n"
        f"Нажмите «🔍 Поиск новых цен», чтобы загрузить актуальные предложения именно для вашего региона."
    )


def _save_items_snapshot_to_store(items: List[ProductItem], meta: dict) -> None:
    if not items or not meta:
        return
    try:
        payload = {
            "meta": meta,
            "items": [asdict(it) for it in items],
        }
        rt.user_store.save_items_snapshot(payload)
    except Exception:
        pass


def _restore_items_snapshot_from_store(target_cfg: dict, target_cid: str = "") -> bool:
    snap = rt.user_store.load_items_snapshot()
    if not isinstance(snap, dict):
        return False
    meta = snap.get("meta")
    raw_items = snap.get("items")
    if not isinstance(meta, dict) or not isinstance(raw_items, list) or not raw_items:
        return False
    fields = set(ProductItem.__dataclass_fields__.keys())
    restored: List[ProductItem] = []
    for row in raw_items:
        if not isinstance(row, dict):
            continue
        clean_row = {k: v for k, v in row.items() if k in fields}
        try:
            restored.append(ProductItem(**clean_row))
        except Exception:
            continue
    if not restored:
        return False
    ok, _ = _meta_matches_user(meta, target_cfg, restored, require_today=True)
    if not ok:
        return False
    rt.items = restored
    rt.global_scan_meta = meta
    rt.last_updated = restored[0].updated_at or _today_str()
    use_cid = str(target_cid or meta.get("cid") or "")
    if use_cid:
        rt.user_items[use_cid] = restored
        rt.user_scan_meta[use_cid] = meta
    return True


def load_subscribers() -> set[str]:
    try:
        if SUBSCRIBERS_PATH.exists():
            data = json.loads(SUBSCRIBERS_PATH.read_text(encoding="utf-8"))
            if isinstance(data, list):
                return {str(x).strip() for x in data if str(x).strip()}
    except Exception:
        pass
    return set()


def save_subscribers(subs: set[str]) -> None:
    try:
        SUBSCRIBERS_PATH.write_text(json.dumps(sorted(list(subs)), ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


rt.active_chat_ids = load_subscribers()


def trigger_immediate_scan(cid: str = "") -> None:
    if cid and cid not in rt.scan_queue:
        rt.scan_queue.append(cid)
    if rt.current_scan_task and not rt.current_scan_task.done():
        rt.current_scan_task.cancel()
    rt.scan_event.set()


def _auto_price_bounds(queries: List[str]) -> tuple[int, int]:
    q_joined = " ".join(queries).lower()
    if any(w in q_joined for w in ("чехол", "стекло", "пленк", "бампер", "кабель", "заряд", "подшипник", "ролик", "колесик", "газлифт", "запчаст")):
        return 50, 5000
    if any(w in q_joined for w in ("ноутбук", "macbook", "макбук", "ryzen", "core i", "thinkpad", "ideapad", "legion", "vivobook", "matebook", "magicbook")):
        return 15000, 250000
    if any(w in q_joined for w in ("телефон", "смартфон", "iphone", "айфон", "ipad", "планшет", "телевизор", "монитор", "видеокарт", "холодильник", "стиральн")):
        return 3000, 200000
    if any(w in q_joined for w in ("кресл", "диван", "кроват", "шкаф", "матрас", "стол", "комод")):
        return 2000, 100000
    if any(w in q_joined for w in ("табурет", "стул", "тумб", "полк", "стеллаж", "пуф", "банкетк")):
        return 400, 40000
    if any(w in q_joined for w in ("шуруповерт", "шуруповёрт", "дрель", "перфоратор", "пылесос", "кофевар", "кофемашин", "микроволнов", "мультиварк", "чайник", "утюг", "фен", "блендер", "обогревател")):
        return 600, 70000
    return 300, 250000


MAX_PRODUCTS = 10

TG_MAIN_KEYBOARD = {
    "keyboard": [
        [{"text": "🔍 Поиск новых цен"}],
        [{"text": "📦 Мои товары"}, {"text": "🔄 Заменить один товар"}],
        [{"text": "⚖️ Сравнение Цена/Качество"}, {"text": "📋 Лучшие цены по магазинам"}],
        [{"text": "📥 Скачать Excel (.xlsx)"}, {"text": "💰 Сменить цену (ОТ и ДО)"}],
        [{"text": "📍 Сменить город"}, {"text": "🔔 Уведомления и интервал"}],
        [{"text": "📌 Отправить мою геолокацию (5–15 км)", "request_location": True}],
    ],
    "resize_keyboard": True,
    "is_persistent": True,
}

TG_CITIES_PAGES = {
    1: [
        ("📍 Москва", "city:Москва"),
        ("📍 Санкт-Петербург", "city:Санкт-Петербург"),
        ("📍 Краснодар", "city:Краснодар"),
        ("📍 Новороссийск", "city:Новороссийск"),
        ("📍 Ростов-на-Дону", "city:Ростов-на-Дону"),
        ("📍 Сочи", "city:Сочи"),
        ("📍 Казань", "city:Казань"),
        ("📍 Екатеринбург", "city:Екатеринбург"),
        ("📍 Новосибирск", "city:Новосибирск"),
        ("📍 Нижний Новгород", "city:Нижний Новгород"),
        ("📍 Самара", "city:Самара"),
        ("📍 Челябинск", "city:Челябинск"),
    ],
    2: [
        ("📍 Уфа", "city:Уфа"),
        ("📍 Красноярск", "city:Красноярск"),
        ("📍 Пермь", "city:Пермь"),
        ("📍 Воронеж", "city:Воронеж"),
        ("📍 Волгоград", "city:Волгоград"),
        ("📍 Саратов", "city:Саратов"),
        ("📍 Тюмень", "city:Тюмень"),
        ("📍 Тольятти", "city:Тольятти"),
        ("📍 Барнаул", "city:Барнаул"),
        ("📍 Ижевск", "city:Ижевск"),
        ("📍 Ульяновск", "city:Ульяновск"),
        ("📍 Иркутск", "city:Иркутск"),
    ],
    3: [
        ("📍 Хабаровск", "city:Хабаровск"),
        ("📍 Владивосток", "city:Владивосток"),
        ("📍 Ярославль", "city:Ярославль"),
        ("📍 Махачкала", "city:Махачкала"),
        ("📍 Томск", "city:Томск"),
        ("📍 Оренбург", "city:Оренбург"),
        ("📍 Кемерово", "city:Кемерово"),
        ("📍 Новокузнецк", "city:Новокузнецк"),
        ("📍 Рязань", "city:Рязань"),
        ("📍 Набережные Челны", "city:Набережные Челны"),
        ("📍 Астрахань", "city:Астрахань"),
        ("📍 Калининград", "city:Калининград"),
    ],
}


def find_city_page(city_name: str) -> int:
    target = (city_name or "").strip().lower()
    for p, lst in TG_CITIES_PAGES.items():
        for _, cb in lst:
            if cb.split(":", 1)[1].lower() == target:
                return p
    return 1


def build_cities_markup(
    page: int = 1,
    selected_city: str = "Все регионы РФ",
    avito_mode: str = "both",
    avito_radius_km: int = 15,
) -> dict:
    page = max(1, min(page, 3))
    sel_low = (selected_city or "Все регионы РФ").strip().lower()
    all_label = "✅ Все регионы РФ (Вся Россия)" if sel_low == "все регионы рф" else "🌍 Все регионы РФ (Вся Россия)"
    buttons = [[{"text": all_label, "callback_data": "city:Все регионы РФ"}]]
    in_presets = sel_low == "все регионы рф"
    for _, lst in TG_CITIES_PAGES.items():
        if any(cb.split(":", 1)[1].lower() == sel_low for _, cb in lst):
            in_presets = True
            break
    if not in_presets and selected_city:
        buttons.append([{"text": f"✅ Выбран город: {selected_city}", "callback_data": "noop"}])

    city_list = TG_CITIES_PAGES.get(page, TG_CITIES_PAGES[1])
    for i in range(0, len(city_list), 2):
        row = []
        for idx in (i, i + 1):
            if idx < len(city_list):
                raw_lbl, cb = city_list[idx]
                c_val = cb.split(":", 1)[1]
                lbl = f"✅ {c_val}" if c_val.lower() == sel_low else raw_lbl
                row.append({"text": lbl, "callback_data": cb})
        buttons.append(row)
    nav_row = []
    if page > 1:
        nav_row.append({"text": "⬅️ Назад", "callback_data": f"citypage:{page - 1}"})
    nav_row.append({"text": f"📄 Стр. {page}/3", "callback_data": "noop"})
    if page < 3:
        nav_row.append({"text": "Вперёд ➡️", "callback_data": f"citypage:{page + 1}"})
    buttons.append(nav_row)

    cur_rad = max(5, min(int(avito_radius_km or 15), 15))
    cur_mode = (avito_mode or "both").strip().lower()
    buttons.append([
        {"text": f"{'✅ ' if cur_rad == 5 else ''}📍 5 км", "callback_data": "avitorad:5"},
        {"text": f"{'✅ ' if cur_rad == 10 else ''}📍 10 км", "callback_data": "avitorad:10"},
        {"text": f"{'✅ ' if cur_rad == 15 else ''}📍 15 км", "callback_data": "avitorad:15"},
    ])
    buttons.append([
        {"text": f"{'✅ ' if cur_mode == 'pickup' else ''}📍 Самовывоз", "callback_data": "avitomode:pickup"},
        {"text": f"{'✅ ' if cur_mode == 'delivery' else ''}🚚 Доставка", "callback_data": "avitomode:delivery"},
        {"text": f"{'✅ ' if cur_mode == 'both' else ''}🌐 Оба", "callback_data": "avitomode:both"},
    ])
    return {"inline_keyboard": buttons}


TG_CITIES_INLINE = build_cities_markup(1)


LADDER_INTERVALS = [180, 1440]  # Циклический авторежим уведомлений: 3 часа ⇄ 1 день (24 ч)
AUTO_INTERVAL_TOKENS = ("ladder", "smart_ladder", "auto", "-1")


def is_auto_interval(cfg_or_val: object, use_ladder: bool = False) -> bool:
    if isinstance(cfg_or_val, dict):
        raw = str(cfg_or_val.get("interval_minutes", "ladder")).strip()
        return raw in AUTO_INTERVAL_TOKENS or bool(cfg_or_val.get("use_ladder"))
    return str(cfg_or_val if cfg_or_val is not None else "").strip() in AUTO_INTERVAL_TOKENS or bool(use_ladder)


def format_interval_str(val: object, use_ladder: bool = False) -> str:
    s = str(val if val is not None else "").strip()
    if is_auto_interval(s, use_ladder):
        return "Авторежим: 3ч ⇄ 1д (циклично)"
    try:
        m = int(s)
    except Exception:
        return f"{s} мин."
    if m < 60:
        return f"{m} мин."
    if m < 1440:
        h = m // 60
        rem = m % 60
        return f"{h} ч." if rem == 0 else f"{h} ч {rem} мин."
    if m < 10080:
        d = m // 1440
        rem_h = (m % 1440) // 60
        rem_str = f" {rem_h} ч." if rem_h else ""
        if d == 1:
            return f"1 день{rem_str}"
        if d in (2, 3, 4):
            return f"{d} дня{rem_str}"
        return f"{d} дней{rem_str}"
    w = m // 10080
    if w == 1:
        return "1 неделя"
    if w in (2, 3, 4):
        return f"{w} недели"
    return f"{w} недель"


def get_effective_interval_minutes(cfg: dict) -> int:
    if is_auto_interval(cfg):
        step = int(cfg.get("ladder_step", 0)) % len(LADDER_INTERVALS)
        return LADDER_INTERVALS[step]
    try:
        return max(5, int(str(cfg.get("interval_minutes", 180)).strip()))
    except Exception:
        return 180


def advance_ladder_step(cfg: dict) -> None:
    if is_auto_interval(cfg):
        cur_step = int(cfg.get("ladder_step", 0))
        cfg["ladder_step"] = (cur_step + 1) % len(LADDER_INTERVALS)


def build_notify_markup(cfg: dict) -> dict:
    cur_raw = str(cfg.get("interval_minutes", "ladder")).strip()
    is_ladder = is_auto_interval(cfg)
    cur_interval = 0 if is_ladder else (int(cur_raw) if cur_raw.isdigit() else 180)
    muted_q = set(cfg.get("muted_queries", []))
    notify_mode = cfg.get("alert_filter", "GREEN_ONLY")  # GREEN_ONLY, ALL, OFF

    mode_row = [
        {"text": ("🟢 Только зелёные ✅" if notify_mode == "GREEN_ONLY" else "🟢 Зелёные"), "callback_data": "notify_mode:GREEN_ONLY"},
        {"text": ("🔔 Все (🟢+🟡) ✅" if notify_mode == "ALL" else "🔔 Все (🟢+🟡)"), "callback_data": "notify_mode:ALL"},
        {"text": ("🔕 Без звука ✅" if notify_mode == "OFF" else "🔕 Без звука"), "callback_data": "notify_mode:OFF"},
    ]

    ladder_row = [
        {
            "text": ("🔄 Авторежим: 3ч ⇄ 1д (циклично) ✅" if is_ladder else "🔄 Авторежим (3ч ⇄ 1д, циклично)"),
            "callback_data": "interval:ladder",
        }
    ]

    row_quick = []
    for iv, lbl in [(15, "15 мин"), (30, "30 мин"), (60, "1 час"), (180, "3 часа")]:
        mark = "✓ " if (not is_ladder and iv == cur_interval) else ""
        row_quick.append({"text": f"{mark}{lbl}", "callback_data": f"interval:{iv}"})

    row_long = []
    for iv, lbl in [(1440, "1 день"), (2880, "2 дня"), (4320, "3 дня"), (10080, "1 неделя")]:
        mark = "✓ " if (not is_ladder and iv == cur_interval) else ""
        row_long.append({"text": f"{mark}{lbl}", "callback_data": f"interval:{iv}"})

    queries = cfg.get("queries", [])[:MAX_PRODUCTS]
    q_rows = []
    for idx, q in enumerate(queries):
        is_muted = q in muted_q
        icon = "🔕" if is_muted else "🔔"
        short_q = q if len(q) <= 24 else (q[:22] + "…")
        q_rows.append([{"text": f"{icon} Поз. {idx + 1}: {short_q}", "callback_data": f"toggle_q:{idx}"}])

    keyboard = [
        mode_row,
        ladder_row,
        row_quick,
        row_long,
        *q_rows,
        [
            {"text": "📦 Управление товарами (✅/❌)", "callback_data": "qact:open"},
            {"text": "🔄 Найти сейчас", "callback_data": "action:scan"},
        ],
    ]
    return {"inline_keyboard": keyboard}


def format_notify_settings_text(cfg: dict) -> str:
    cur_raw = cfg.get("interval_minutes", "ladder")
    is_ladder = is_auto_interval(cfg)
    step = int(cfg.get("ladder_step", 0)) % len(LADDER_INTERVALS)
    muted_q = set(cfg.get("muted_queries", []))
    notify_mode = cfg.get("alert_filter", "GREEN_ONLY")
    mode_descr = {
        "GREEN_ONLY": "🟢 Только самые выгодные (Зелёные позиции)",
        "ALL": "🔔 Все подходящие позиции (Зелёные + Жёлтые)",
        "OFF": "🔕 Отключение звуковых алертов (Тихий режим в фоне)",
    }.get(notify_mode, notify_mode)

    if is_ladder:
        ladder_names = ["3 часа", "1 день (24 ч)"]
        cur_step_str = ladder_names[step]
        interval_info = (
            f"<b>Авторежим (цикл):</b> <b>3 часа ⇄ 1 день ⇄ 3 часа ⇄ 1 день</b>.\n"
            f"  • <i>Текущий этап цикла: {cur_step_str} (шаг {step + 1} из {len(LADDER_INTERVALS)})</i>"
        )
    else:
        interval_info = f"<b>каждые {format_interval_str(cur_raw)}</b>"

    q_lines = []
    for idx, q in enumerate(cfg.get("queries", [])[:MAX_PRODUCTS], 1):
        st = "🔕 Отключено" if q in muted_q else "🔔 Включено"
        q_lines.append(f"  {idx}. <b>{_h(q)}</b> — <i>{st}</i>")

    return (
        f"🔔 <b>Настройки уведомлений и интервала обновления</b>\n\n"
        f"⏱ <b>Интервал проверки цен:</b> {interval_info}\n"
        f"🎯 <b>Режим оповещений:</b> {mode_descr}\n\n"
        f"📋 <b>Уведомления по каждой позиции (до {MAX_PRODUCTS}):</b>\n"
        + ("\n".join(q_lines) if q_lines else "  • <i>(список пуст)</i>")
        + f"\n\n👇 <b>Нажмите на любую позицию ниже, чтобы включить/убрать уведомление 🔔/🔕, или выберите интервал:</b>"
    )


SHOPS_META = [
    ("ozon", "Ozon", "1. Ozon"),
    ("aliexpress", "AliExpress", "2. AliExpress"),
    ("ymarket", "Яндекс Маркет", "3. Яндекс Маркет"),
    ("avito", "Авито", "4. Авито"),
    ("dns", "DNS", "5. DNS"),
    ("wb", "Wildberries", "6. Wildberries"),
]


def _get_shops_enabled_map(cfg: dict) -> dict:
    raw = cfg.get("shops_priority") or {}
    return {k: bool(raw.get(k, True)) for k, _, _ in SHOPS_META}


def build_shops_top_markup(cfg: Optional[dict] = None, active_shop: str = "ALL", avito_view: str = "ALL") -> dict:
    enabled = _get_shops_enabled_map(cfg or {})
    rows = [
        [
            {"text": ("✅ Лучшие (Все магазины)" if active_shop == "ALL" else "🔥 Лучшие (Все магазины)"), "callback_data": "shop:ALL"},
            {"text": "⚖️ Цена / Качество", "callback_data": "qact:compare"},
        ]
    ]
    pair = []
    for key, name, label in SHOPS_META:
        is_sel = active_shop == name
        is_on = enabled.get(key, True)
        if is_sel:
            txt = f"✅ {label}"
        elif not is_on:
            txt = f"⬜ {label} (выкл)"
        else:
            txt = label
        pair.append({"text": txt, "callback_data": f"shop:{name}"})
        if len(pair) == 2:
            rows.append(pair)
            pair = []
    if pair:
        rows.append(pair)
    if active_shop == "Авито":
        av_rad = max(5, min(int((cfg or {}).get("avito_radius_km", 15) or 15), 15))
        av_cur = (avito_view or "ALL").upper()
        rows.append([
            {"text": f"{'✅ ' if av_cur == 'PICKUP' else ''}📍 Самовывоз ≤{av_rad} км", "callback_data": "avitoview:PICKUP"},
            {"text": f"{'✅ ' if av_cur == 'DELIVERY' else ''}🚚 Доставка", "callback_data": "avitoview:DELIVERY"},
            {"text": f"{'✅ ' if av_cur == 'ALL' else ''}🌐 Всё", "callback_data": "avitoview:ALL"},
        ])
    rows.append([
        {"text": "⚙️ Магазины (✅/⬜)", "callback_data": "menu:shops_filter"},
        {"text": "🔄 Обновить цены (регион)", "callback_data": "qact:refresh_now"},
    ])
    return {"inline_keyboard": rows}


def build_shops_filter_markup(cfg: dict) -> dict:
    enabled = _get_shops_enabled_map(cfg)
    rows = []
    pair = []
    for key, name, label in SHOPS_META:
        mark = "✅" if enabled.get(key, True) else "⬜"
        pair.append({"text": f"{mark} {label}", "callback_data": f"toggleshop:{key}"})
        if len(pair) == 2:
            rows.append(pair)
            pair = []
    if pair:
        rows.append(pair)
    rows.append([
        {"text": "✅ Включить все 6", "callback_data": "toggleshop:ALL"},
        {"text": "🔍 Искать по отмеченным", "callback_data": "action:scan"},
    ])
    rows.append([{"text": "⬅️ Назад ко всем магазинам", "callback_data": "shop:ALL"}])
    return {"inline_keyboard": rows}


TG_SHOPS_INLINE = build_shops_top_markup()


def _active_queries(cfg: dict) -> List[str]:
    queries = [str(x).strip() for x in (cfg.get("queries") or []) if str(x).strip()][:MAX_PRODUCTS]
    if not queries:
        return []
    selected = cfg.get("selected_queries")
    if isinstance(selected, list) and selected:
        sel_set = set(selected)
        chosen = [q for q in queries if q in sel_set]
        if chosen:
            return chosen[:MAX_PRODUCTS]
    return queries


def _sync_query_lists(cfg: dict) -> None:
    queries = [str(x).strip() for x in (cfg.get("queries") or []) if str(x).strip()][:MAX_PRODUCTS]
    cfg["queries"] = queries
    q_set = set(queries)
    sel = cfg.get("selected_queries")
    if not isinstance(sel, list):
        cfg["selected_queries"] = list(queries)
    else:
        cfg["selected_queries"] = [q for q in queries if q in set(sel)]
    muted = [q for q in (cfg.get("muted_queries") or []) if q in q_set]
    cfg["muted_queries"] = muted


def build_products_markup(cfg: dict) -> dict:
    _sync_query_lists(cfg)
    queries = cfg.get("queries", [])
    sel_set = set(cfg.get("selected_queries", []))
    muted_set = set(cfg.get("muted_queries", []))
    rows = []
    for idx, q in enumerate(queries):
        is_sel = q in sel_set
        is_muted = q in muted_set
        short_q = q if len(q) <= 18 else (q[:16] + "…")
        rows.append([
            {"text": f"{'✅' if is_sel else '⬜'} {idx + 1}. {short_q}", "callback_data": f"qsel:{idx}"},
            {"text": ("🔕" if is_muted else "🔔"), "callback_data": f"qtog:{idx}"},
            {"text": "✏️", "callback_data": f"qrep:{idx}"},
            {"text": "❌", "callback_data": f"qdel:{idx}"},
        ])
    rows.append([
        {"text": "➕ Добавить товар", "callback_data": "qact:add"},
        {"text": "🔄 Заменить 1 товар", "callback_data": "qact:replace_menu"},
    ])
    rows.append([
        {"text": "✅ Выбрать все", "callback_data": "qact:all"},
        {"text": "🗑 Убрать снятые ⬜", "callback_data": "qact:clear_off"},
    ])
    rows.append([
        {"text": "🔔 Вкл. уведомления", "callback_data": "qact:unmute_all"},
        {"text": "🆕 Новый список", "callback_data": "qact:replace_all"},
    ])
    rows.append([
        {"text": "⚖️ Цена / Качество", "callback_data": "qact:compare"},
        {"text": "🔍 Найти выбранные ✅", "callback_data": "action:scan"},
    ])
    return {"inline_keyboard": rows}


def build_replace_pick_markup(cfg: dict) -> dict:
    _sync_query_lists(cfg)
    queries = cfg.get("queries", [])
    rows = []
    for idx, q in enumerate(queries):
        short_q = q if len(q) <= 30 else (q[:28] + "…")
        rows.append([{"text": f"✏️ Заменить {idx + 1}: {short_q}", "callback_data": f"qrep:{idx}"}])
    rows.append([{"text": "🆕 Заменить весь список целиком", "callback_data": "qact:replace_all"}])
    rows.append([{"text": "⬅️ Назад в «📦 Мои товары»", "callback_data": "qact:open"}])
    return {"inline_keyboard": rows}


def build_pq_markup(cfg: Optional[dict] = None, active_q: str = "ALL") -> dict:
    c = cfg or {}
    _sync_query_lists(c)
    queries = c.get("queries", [])
    rows = [[{"text": ("✅ Все товары (сводка)" if active_q == "ALL" else "🌐 Все товары (сводка)"), "callback_data": "pq:ALL"}]]
    if len(queries) > 1:
        pair = []
        for idx, q in enumerate(queries[:MAX_PRODUCTS]):
            short_q = q if len(q) <= 18 else (q[:16] + "…")
            mark = "✅ " if active_q == q else ""
            pair.append({"text": f"{mark}{idx + 1}. {short_q}", "callback_data": f"pqq:{idx}"})
            if len(pair) == 2:
                rows.append(pair)
                pair = []
        if pair:
            rows.append(pair)
    rows.append([
        {"text": "🔄 Заменить 1 товар", "callback_data": "qact:replace_menu"},
        {"text": "📋 Лучшие по магазинам", "callback_data": "shop:ALL"},
    ])
    rows.append([{"text": "🔄 Обновить цены (регион)", "callback_data": "qact:refresh_now"}])
    return {"inline_keyboard": rows}


def format_products_menu_text(cfg: dict) -> str:
    _sync_query_lists(cfg)
    queries = cfg.get("queries", [])
    sel_set = set(cfg.get("selected_queries", []))
    muted_set = set(cfg.get("muted_queries", []))
    lines = [
        f"📦 <b>Ваши отслеживаемые товары ({len(queries)} шт.)</b>",
        "<i>Нажмите на название — ✅/⬜ включить в поиск; 🔔/🔕 — уведомления; ✏️ — заменить товар; ❌ — удалить.</i>\n",
    ]
    if not queries:
        lines.append("• <i>Список пуст. Нажмите «➕ Добавить товар» или напишите название в чат.</i>")
    else:
        for idx, q in enumerate(queries, 1):
            s_mark = "✅ В поиске" if q in sel_set else "⬜ Пауза"
            n_mark = "🔕 Без уведомл." if q in muted_set else "🔔 Уведомл. вкл"
            lines.append(f"{idx}. <b>{_h(q)}</b> — {s_mark} | {n_mark}")
    return "\n".join(lines)


def compute_price_quality_score(it: ProductItem, ref_price: int) -> float:
    base_ref = max(1, ref_price or it.price)
    ratio = base_ref / max(1, it.price)
    price_pts = 50.0 * min(1.0, ratio ** 0.72)
    rating = float(it.seller_rating or 0.0)
    if rating > 0:
        rating_pts = 16.0 * max(0.0, min(1.5, rating - 3.5)) / 1.5
    else:
        rating_pts = 8.0
    reviews_pts = min(9.0, math.log1p(max(0, int(it.reviews_count or 0))) * 2.2)
    cond_pts = float(it.score or 50) * 0.20
    if it.status_code == "GREEN":
        cond_pts += 5.0
    if "проверено на карточке" in (it.duty_note or ""):
        cond_pts += 2.0
    return round(min(99.9, price_pts + rating_pts + reviews_pts + cond_pts), 1)


def _format_single_query_pq(
    q_label: str,
    q_items: List[ProductItem],
    max_rows: int = 5,
    compact: bool = False,
    index_num: Optional[int] = None,
) -> str:
    greens = [i for i in q_items if i.status_code == "GREEN"]
    yellows = [i for i in q_items if i.status_code == "YELLOW"]
    if len(greens) >= 3:
        min_g = min(x.price for x in greens)
        good_yellows = [x for x in yellows if x.price < min_g * 1.15 and int(x.score or 0) >= 55]
        pool = greens + good_yellows
    else:
        pool = q_items
    ref_p = min(x.price for x in pool)
    scored = sorted(
        [(compute_price_quality_score(x, ref_p), x) for x in pool],
        key=lambda pair: (-pair[0], pair[1].price),
    )
    cheapest = min(pool, key=lambda x: x.price)
    best_pq_val, best_pq_item = scored[0]

    prefix = f"{index_num}. " if index_num else ""
    lines = ["━━━━━━━━━━━━━━━━━━━━"]
    if q_label:
        lines.append(f"📦 <b>{prefix}{_h(q_label)}</b> <i>({len(q_items)} в наличии)</i>")

    ch_url = clean_product_url(cheapest.url)
    bp_url = clean_product_url(best_pq_item.url)
    ch_price_txt = (
        f"Итого <b>{cheapest.price:,} ₽</b> <i>({cheapest.base_price:,} ₽ + пошлина {cheapest.customs_duty:,} ₽)</i>"
        if cheapest.customs_duty > 0
        else f"<b>{cheapest.price:,} ₽</b> <i>(пошлина 0 ₽)</i>"
    )
    lines.append(
        f"💎 <b>Минимум цены:</b> {ch_price_txt} · <b>{_h(cheapest.shop)}</b>\n"
        f"   └ <a href=\"{_h(ch_url)}\">{_h(cheapest.title[:48])} ↗</a>"
    )
    bp_price_txt = (
        f"Итого <b>{best_pq_item.price:,} ₽</b> <i>({best_pq_item.base_price:,} ₽ + пошлина {best_pq_item.customs_duty:,} ₽)</i>"
        if best_pq_item.customs_duty > 0
        else f"<b>{best_pq_item.price:,} ₽</b> <i>(пошлина 0 ₽)</i>"
    )
    r_str = f"⭐{best_pq_item.seller_rating:.1f}" if best_pq_item.seller_rating else "⭐—"
    rev_str = f", {best_pq_item.reviews_count} отз." if best_pq_item.reviews_count else ""
    lines.append(
        f"🏆 <b>Лучший баланс ({best_pq_val}/100):</b> {bp_price_txt} · {r_str}{rev_str} · <b>{_h(best_pq_item.shop)}</b>\n"
        f"   └ <a href=\"{_h(bp_url)}\">{_h(best_pq_item.title[:48])} ↗</a>"
    )

    lines.append("\n🏬 <b>Лучшее по магазинам:</b>")
    for sh_name in ("Ozon", "AliExpress", "Яндекс Маркет", "Авито", "DNS", "Wildberries"):
        sh_scored = [(pq, x) for pq, x in scored if x.shop == sh_name]
        if sh_scored:
            pq_v, x = sh_scored[0]
            sh_p_txt = (
                f"Итого <b>{x.price:,} ₽</b> <i>({x.base_price:,} + пошлина {x.customs_duty:,} ₽)</i>"
                if x.customs_duty > 0
                else f"<b>{x.price:,} ₽</b>"
            )
            rs = f"⭐{x.seller_rating:.1f}" if x.seller_rating else "⭐—"
            sh_url = clean_product_url(x.url)
            lines.append(
                f"• <b>{_h(sh_name)}</b> — {sh_p_txt} | {rs} | Ц/К <b>{pq_v}</b> · "
                f"<a href=\"{_h(sh_url)}\">{_h(x.title[:30])} ↗</a>"
            )
        elif not compact:
            lines.append(f"• <b>{_h(sh_name)}</b> — <i>нет подходящих</i>")

    show_n = min(max_rows, len(scored))
    if show_n > 0 and not compact:
        lines.append(f"\n📊 <b>Топ-{show_n} по рейтингу Цена / Качество:</b>")
        for idx, (pq_v, it) in enumerate(scored[:show_n], 1):
            st_icon = "🟢" if it.status_code == "GREEN" else "🟡"
            it_p_txt = (
                f"Итого {it.price:,} ₽</b> <i>({it.base_price:,} + пошлина {it.customs_duty:,} ₽)</i>"
                if it.customs_duty > 0
                else f"{it.price:,} ₽</b>"
            )
            rs = f" · ⭐{it.seller_rating:.1f}" if it.seller_rating else ""
            rev = f" ({it.reviews_count} отз.)" if it.reviews_count else ""
            it_url = clean_product_url(it.url)
            lines.append(
                f"{idx}. {st_icon} <b>{_h(it.shop)} · {it_p_txt} — <b>{pq_v}/100</b>{rs}{rev}\n"
                f"   └ <a href=\"{_h(it_url)}\">{_h(it.title[:52])} ↗</a>"
            )
    return "\n".join(lines)


def format_price_quality_comparison_tg(
    items: Optional[List[ProductItem]] = None,
    cid: str = "",
    max_rows: int = 10,
    query_filter: str = "ALL",
    u_cfg: Optional[dict] = None,
) -> str:
    eff_cfg = u_cfg if u_cfg is not None else (rt.user_store.get(str(cid)) if (cid and rt.user_store.exists(str(cid))) else {})
    cur_reg = str((eff_cfg or {}).get("region") or "Все регионы РФ")
    if items is not None:
        source_items = items
        empty_reason = ""
    else:
        source_items, empty_reason = _get_scoped_items(cid, eff_cfg)
    if not source_items and empty_reason:
        return empty_reason

    avail = [i for i in (source_items or []) if str(i.in_stock).startswith("✅") and i.price > 0 and i.status_code in ("GREEN", "YELLOW")]
    if not avail:
        return (
            f"⚖️ <b>Сравнение Цена / Качество ({_h(cur_reg)}) пока недоступно.</b>\n"
            f"Нажмите «🔍 Поиск новых цен», чтобы собрать актуальные предложения со всех 6 магазинов."
        )

    by_q: Dict[str, List[ProductItem]] = {}
    for x in avail:
        q_name = (x.query or "Товар").strip()
        by_q.setdefault(q_name, []).append(x)

    if query_filter and query_filter != "ALL":
        matched_q = next((qk for qk in by_q if qk.lower() == query_filter.strip().lower()), None)
        if not matched_q:
            return (
                f"⚖️ По товару <b>«{_h(query_filter)}»</b> в регионе <b>{_h(cur_reg)}</b> пока нет доступных предложений.\n"
                "Нажмите «🌐 Все товары (сводка)» или запустите новый поиск."
            )
        header = (
            f"⚖️ <b>Сравнение «Цена / Качество» — по 6 магазинам</b>\n"
            f"📍 <b>Регион:</b> {_h(cur_reg)} · <i>Рейтинг 0–100 (цена, пошлина, рейтинг, отзывы)</i>"
        )
        return header + "\n\n" + _format_single_query_pq(matched_q, by_q[matched_q], max_rows=max_rows, compact=False)

    header = (
        f"⚖️ <b>Сравнение «Цена / Качество» — {len(by_q)} поз. × 6 магазинов</b>\n"
        f"📍 <b>Регион:</b> {_h(cur_reg)} · <i>Рейтинг 0–100 (цена, пошлина, рейтинг, отзывы)</i>"
    )
    if len(by_q) == 1:
        q_only, q_list = next(iter(by_q.items()))
        return header + "\n\n" + _format_single_query_pq(q_only, q_list, max_rows=max_rows, compact=False)

    per_q_rows = 3 if len(by_q) <= 3 else 2
    compact_mode = len(by_q) > 2
    blocks = []
    for pos_i, (q_name, q_list) in enumerate(by_q.items(), 1):
        blocks.append(_format_single_query_pq(q_name, q_list, max_rows=per_q_rows, compact=compact_mode, index_num=pos_i))
    return header + "\n\n" + "\n\n".join(blocks)


TOKEN_FILE_PATH = BASE_DIR / ".tg_token"
LOCAL_AUTH_PATH = BASE_DIR / ".tg_auth.json"


def load_config() -> dict:
    if not CONFIG_PATH.exists():
        raise FileNotFoundError(f"Файл настроек не найден: {CONFIG_PATH}")
    cfg = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))

    # Загружаем локальные учётные данные из .tg_auth.json (не попадает в Git)
    if LOCAL_AUTH_PATH.exists():
        try:
            local_auth = json.loads(LOCAL_AUTH_PATH.read_text(encoding="utf-8-sig"))
            tg_sec = cfg.setdefault("telegram", {})
            if local_auth.get("bot_token") and not tg_sec.get("bot_token"):
                tg_sec["bot_token"] = str(local_auth["bot_token"]).strip()
                tg_sec["enabled"] = True
            if local_auth.get("chat_id") and not tg_sec.get("chat_id"):
                tg_sec["chat_id"] = str(local_auth["chat_id"]).strip()
            if local_auth.get("allowed_chat_ids") and not tg_sec.get("allowed_chat_ids"):
                tg_sec["allowed_chat_ids"] = list(local_auth["allowed_chat_ids"])
        except Exception:
            pass

    env_token = os.environ.get("TG_BOT_TOKEN", "").strip()
    if not env_token and TOKEN_FILE_PATH.exists():
        try:
            env_token = TOKEN_FILE_PATH.read_text(encoding="utf-8").strip()
        except Exception:
            pass
    env_chat = os.environ.get("TG_CHAT_ID", "").strip()
    env_allowed = os.environ.get("TG_ALLOWED_CHATS", "").strip()

    if env_token:
        cfg.setdefault("telegram", {})["bot_token"] = env_token
        cfg["telegram"]["enabled"] = True
    if env_chat:
        cfg.setdefault("telegram", {})["chat_id"] = env_chat
    if env_allowed:
        cfg.setdefault("telegram", {})["allowed_chat_ids"] = [x.strip() for x in env_allowed.split(",") if x.strip()]
    elif env_chat and not cfg.get("telegram", {}).get("allowed_chat_ids"):
        cfg.setdefault("telegram", {})["allowed_chat_ids"] = [env_chat]

    return cfg


async def _sync_desc_to_tg(token_val: str, state_str: str) -> None:
    try:
        async with httpx.AsyncClient(timeout=5.0) as cl:
            await cl.post(
                f"https://api.telegram.org/bot{token_val}/setMyDescription",
                json={"description": state_str[:500]},
            )
    except Exception:
        pass


def save_config(cfg: dict) -> None:
    tg_data = cfg.get("telegram", {})
    if isinstance(tg_data, dict) and (tg_data.get("bot_token") or tg_data.get("chat_id")):
        try:
            existing_auth = {}
            if LOCAL_AUTH_PATH.exists():
                existing_auth = json.loads(LOCAL_AUTH_PATH.read_text(encoding="utf-8-sig"))
            if tg_data.get("bot_token"):
                existing_auth["bot_token"] = tg_data["bot_token"]
            if tg_data.get("chat_id"):
                existing_auth["chat_id"] = tg_data["chat_id"]
            if tg_data.get("allowed_chat_ids"):
                existing_auth["allowed_chat_ids"] = tg_data["allowed_chat_ids"]
            LOCAL_AUTH_PATH.write_text(json.dumps(existing_auth, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    safe_cfg = json.loads(json.dumps(cfg))
    for priv_key in ("lat", "lon", "detected_place"):
        safe_cfg.pop(priv_key, None)
    if "telegram" in safe_cfg and isinstance(safe_cfg["telegram"], dict):
        safe_cfg["telegram"]["bot_token"] = ""
        safe_cfg["telegram"]["chat_id"] = ""
        safe_cfg["telegram"]["allowed_chat_ids"] = []
    CONFIG_PATH.write_text(json.dumps(safe_cfg, ensure_ascii=False, indent=2), encoding="utf-8")


def load_history() -> Dict[str, int]:
    if HISTORY_PATH.exists():
        try:
            return json.loads(HISTORY_PATH.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_history(history: Dict[str, int]) -> None:
    HISTORY_PATH.write_text(json.dumps(history, ensure_ascii=False, indent=2), encoding="utf-8")


def load_items_from_existing_excel() -> List[ProductItem]:
    if not EXCEL_PATH.exists():
        return []
    today_s = _today_str()
    try:
        file_day = datetime.fromtimestamp(EXCEL_PATH.stat().st_mtime).strftime("%d.%m.%Y")
        if file_day != today_s:
            return []
    except Exception:
        pass

    items: List[ProductItem] = []
    seen = set()
    q_char = chr(34)
    detected_city = ""
    try:
        with zipfile.ZipFile(EXCEL_PATH, "r") as z:
            try:
                if "xl/workbook.xml" in z.namelist():
                    wb_xml = z.read("xl/workbook.xml").decode("utf-8", errors="ignore")
                    m_city = re.search(r'name="5_DNS\s*\(([^)]+)\)"', wb_xml)
                    if m_city:
                        c_raw = m_city.group(1).strip()
                        detected_city = "Все регионы РФ" if "все регион" in c_raw.lower() else c_raw
            except Exception:
                pass

            ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
            for name in sorted(z.namelist()):
                if not name.startswith("xl/worksheets/sheet"):
                    continue
                root = ET.fromstring(z.read(name).decode("utf-8"))
                for r in root.findall(".//s:row", ns)[1:]:
                    cells = r.findall("s:c", ns)
                    if len(cells) < 21:
                        continue
                    vals = ["".join(c.itertext()) for c in cells]
                    row_updated = str(vals[20] or "").strip()
                    if row_updated and not row_updated.startswith(today_s):
                        continue
                    f_el = cells[18].find("s:f", ns)
                    url = (
                        f_el.text.split(q_char)[1]
                        if (f_el is not None and f_el.text and q_char in f_el.text)
                        else vals[18]
                    )
                    if not url or url in seen:
                        continue
                    seen.add(url)
                    base_p = int(vals[9] or vals[8] or 0)
                    it = _build_product(
                        shop=vals[1],
                        query=vals[2],
                        title=vals[7],
                        price=base_p,
                        seller_name=vals[13],
                        seller_rating=float(vals[15] or 0),
                        reviews_count=int(vals[16] or 0),
                        location_delivery=vals[17],
                        description=vals[19],
                        url=url,
                    )
                    it.price_change = vals[12]
                    it.updated_at = row_updated
                    if "Нет в продаже" in vals[3] or "Распродан" in vals[4]:
                        it.in_stock = "❌ Нет в продаже (Распродан)"
                        it.status_code = "RED"
                        it.status_label = "🔴 КРАСНЫЙ (Нет в наличии / Распродан)"
                    items.append(it)
        if items:
            rt.last_updated = items[0].updated_at or today_s
            stored_meta = dict((rt.user_store.meta or {}).get("last_global_meta") or {})
            q_seen = list(dict.fromkeys([x.query for x in items if x.query]))
            reg_final = detected_city or str(stored_meta.get("region") or "")
            if reg_final:
                rt.global_scan_meta = _build_scan_meta(
                    city_name=reg_final,
                    queries=stored_meta.get("queries") or q_seen,
                    cid=str(stored_meta.get("cid") or ""),
                    avito_mode=str(stored_meta.get("avito_mode") or "both"),
                    avito_radius_km=int(stored_meta.get("avito_radius_km") or 15),
                    ts=float(stored_meta.get("ts") or EXCEL_PATH.stat().st_mtime),
                )
    except Exception:
        pass
    return items


def _h(val: object) -> str:
    return html.escape(str(val if val is not None else ""), quote=True)


def _safe_truncate_html(text: str, limit: int = 3980) -> str:
    """Безопасно обрезает Telegram HTML по границе строки и закрывает незакрытые теги."""
    raw = str(text or "")
    if len(raw) <= limit:
        return raw
    cut_limit = max(200, limit - 48)
    cut = raw[:cut_limit]
    last_nl = cut.rfind("\n")
    if last_nl > cut_limit // 2:
        cut = cut[:last_nl]
    cut = re.sub(r"<[^>]*$", "", cut).rstrip()

    stack: List[str] = []
    for m in re.finditer(r"<(/?)(a|b|i|code|pre|u|s)(?:\s+[^>]*)?>", cut, flags=re.I):
        is_close = bool(m.group(1))
        tag = m.group(2).lower()
        if not is_close:
            stack.append(tag)
        else:
            for idx in range(len(stack) - 1, -1, -1):
                if stack[idx] == tag:
                    del stack[idx]
                    break
    if stack:
        closing = "".join(f"</{t}>" for t in reversed(stack))
        return f"{cut}{closing}\n…"
    return f"{cut}\n…"


async def tg_send_message(
    client: httpx.AsyncClient,
    token: str,
    chat_id: str,
    text: str,
    reply_markup: Optional[dict] = None,
) -> Optional[int]:
    safe_text = _safe_truncate_html(text, 3980)
    payload = {
        "chat_id": chat_id,
        "text": safe_text,
        "parse_mode": "HTML",
        "disable_web_page_preview": True,
    }
    if reply_markup is not None:
        payload["reply_markup"] = reply_markup
    try:
        r = await client.post(f"https://api.telegram.org/bot{token}/sendMessage", json=payload, timeout=15.0)
        data = r.json()
        if not data.get("ok"):
            plain_txt = re.sub(r"</?(?:b|i|code|pre|a|u|s)(?:\s+[^>]*)?>", "", safe_text)
            payload_plain = dict(payload)
            payload_plain["text"] = plain_txt[:4000]
            payload_plain.pop("parse_mode", None)
            r = await client.post(f"https://api.telegram.org/bot{token}/sendMessage", json=payload_plain, timeout=15.0)
            data = r.json()
        return data.get("result", {}).get("message_id")
    except Exception:
        return None


async def tg_send_html_chunks(
    client: httpx.AsyncClient,
    token: str,
    chat_id: str,
    text: str,
    reply_markup: Optional[dict] = None,
    chunk_limit: int = 3650,
) -> Optional[int]:
    """Разбивает длинный HTML-отчёт строго по границам блоков, чтобы не рвать ссылки и теги."""
    raw = str(text or "")
    if len(raw) <= chunk_limit:
        return await tg_send_message(client, token, chat_id, raw, reply_markup=reply_markup)

    blocks = raw.split("\n\n")
    chunks: List[str] = []
    cur = ""
    for blk in blocks:
        candidate = f"{cur}\n\n{blk}".strip() if cur else blk
        if len(candidate) <= chunk_limit:
            cur = candidate
        else:
            if cur:
                chunks.append(cur)
                cur = ""
            if len(blk) <= chunk_limit:
                cur = blk
            else:
                sub = ""
                for ln in blk.split("\n"):
                    cand_l = f"{sub}\n{ln}".strip() if sub else ln
                    if len(cand_l) <= chunk_limit:
                        sub = cand_l
                    else:
                        if sub:
                            chunks.append(sub)
                        sub = ln
                if sub:
                    cur = sub
    if cur:
        chunks.append(cur)
    if not chunks:
        chunks = [raw[:chunk_limit]]

    last_mid: Optional[int] = None
    last_idx = len(chunks) - 1
    for i, ch in enumerate(chunks):
        rm = reply_markup if i == last_idx else None
        last_mid = await tg_send_message(client, token, chat_id, ch, reply_markup=rm)
        if i < last_idx:
            await asyncio.sleep(0.18)
    return last_mid


async def tg_edit_message(
    client: httpx.AsyncClient,
    token: str,
    chat_id: str,
    message_id: int,
    text: str,
    reply_markup: Optional[dict] = None,
) -> None:
    safe_text = _safe_truncate_html(text, 3980)
    try:
        payload = {
            "chat_id": chat_id,
            "message_id": message_id,
            "text": safe_text,
            "parse_mode": "HTML",
            "disable_web_page_preview": True,
        }
        if reply_markup is not None:
            payload["reply_markup"] = reply_markup
        r = await client.post(
            f"https://api.telegram.org/bot{token}/editMessageText",
            json=payload,
            timeout=10.0,
        )
        if not r.json().get("ok"):
            plain = {
                "chat_id": chat_id,
                "message_id": message_id,
                "text": re.sub(r"</?(?:b|i|code|pre|a|u|s)(?:\s+[^>]*)?>", "", safe_text)[:4000],
                "disable_web_page_preview": True,
            }
            if reply_markup is not None:
                plain["reply_markup"] = reply_markup
            await client.post(
                f"https://api.telegram.org/bot{token}/editMessageText",
                json=plain,
                timeout=10.0,
            )
    except Exception:
        pass


async def send_telegram_excel(
    client: Optional[httpx.AsyncClient],
    token: str,
    chat_id: str,
    caption: str = "",
    excel_file: Optional[Path] = None,
) -> bool:
    target_path = excel_file or rt.user_excel_paths.get(str(chat_id)) or EXCEL_PATH
    if not target_path.exists():
        if EXCEL_PATH.exists():
            target_path = EXCEL_PATH
        else:
            if client is not None:
                await tg_send_message(client, token, chat_id, "⚠️ Файл Мониторинг_Товаров.xlsx ещё создаётся, подождите минуту.", TG_MAIN_KEYBOARD)
            return False

    try:
        file_bytes = target_path.read_bytes()
    except Exception as e:
        console.print(f"[red]❌ [TG Excel] Не удалось прочитать {target_path}: {e}[/red]")
        return False
    if not file_bytes:
        console.print(f"[red]❌ [TG Excel] Пустой файл {target_path}[/red]")
        return False

    rm_json = json.dumps(TG_MAIN_KEYBOARD, ensure_ascii=False)
    last_err = ""
    for attempt in range(1, 4):
        cap_val = (caption[:950] if attempt == 1 else (caption.splitlines()[0][:200] if caption else "Таблица Мониторинг_Товаров.xlsx"))
        data_form = {"chat_id": str(chat_id), "caption": cap_val}
        if attempt <= 2:
            data_form["reply_markup"] = rm_json
        files_form = {
            "document": (
                "Мониторинг_Товаров.xlsx",
                file_bytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        }
        try:
            if attempt == 1 and client is not None:
                resp = await client.post(
                    f"https://api.telegram.org/bot{token}/sendDocument",
                    data=data_form,
                    files=files_form,
                    timeout=45.0,
                )
            else:
                async with httpx.AsyncClient(timeout=45.0, trust_env=False) as fresh_cl:
                    resp = await fresh_cl.post(
                        f"https://api.telegram.org/bot{token}/sendDocument",
                        data=data_form,
                        files=files_form,
                        timeout=45.0,
                    )
            js_resp = resp.json() if resp.content else {}
            if resp.status_code == 200 and js_resp.get("ok"):
                console.print(f"[green]✅ [TG Excel] Отправлен ({len(file_bytes):,} байт, попытка {attempt}) -> chat_id={chat_id}[/green]")
                return True
            last_err = f"HTTP {resp.status_code}: {js_resp.get('description') or resp.text[:180]}"
        except Exception as e:
            last_err = str(e)
        if attempt < 3:
            await asyncio.sleep(1.5 * attempt)

    console.print(f"[red]❌ [TG Excel] Не удалось отправить после 3 попыток (chat_id={chat_id}): {last_err}[/red]")
    return False


def format_top_items_tg(shop_filter: str = "ALL", cid: str = "", u_cfg: Optional[dict] = None) -> str:
    s_cid = str(cid or "")
    eff_cfg = u_cfg if u_cfg is not None else (rt.user_store.get(s_cid) if (s_cid and rt.user_store.exists(s_cid)) else {})
    cur_reg = str((eff_cfg or {}).get("region") or "Все регионы РФ")
    av_rad = max(5, min(int((eff_cfg or {}).get("avito_radius_km", 15) or 15), 15))
    source_items, empty_reason = _get_scoped_items(s_cid, eff_cfg)
    if not source_items:
        return empty_reason or f"📋 <b>{_h(cur_reg)}:</b> пока нет доступных позиций в наличии.\nНажмите «🔍 Поиск новых цен»."

    greens = [i for i in source_items if i.status_code == "GREEN" and str(i.in_stock).startswith("✅")]
    yellows = [i for i in source_items if i.status_code == "YELLOW" and str(i.in_stock).startswith("✅")]

    if shop_filter != "ALL":
        greens = [i for i in greens if i.shop == shop_filter]
        yellows = [i for i in yellows if i.shop == shop_filter]

    avito_view = rt.avito_view_filter.get(s_cid, "ALL") if shop_filter == "Авито" else "ALL"
    if shop_filter == "Авито" and avito_view in ("PICKUP", "DELIVERY"):
        def _match_av_view(it: ProductItem) -> bool:
            ld = (it.location_delivery or "").strip()
            if avito_view == "PICKUP":
                return ld.startswith("📍") or "самовывоз" in ld.lower()
            return ld.startswith("🚚") or "доставк" in ld.lower()

        greens = [i for i in greens if _match_av_view(i)]
        yellows = [i for i in yellows if _match_av_view(i)]

    if greens:
        min_g = min(x.price for x in greens)
        cheaper_y = [x for x in yellows if x.price < min_g * 1.05 and int(x.score or 0) >= 55]
        pool = greens + cheaper_y
    else:
        pool = yellows
    pool.sort(key=lambda x: x.price)

    title_hdr = "Все 6 магазинов" if shop_filter == "ALL" else _h(shop_filter)
    av_note = ""
    if shop_filter == "Авито":
        if avito_view == "PICKUP":
            av_note = f" · 📍 Самовывоз ≤{av_rad} км"
        elif avito_view == "DELIVERY":
            av_note = " · 🚚 Доставка"
        else:
            av_note = f" · 📍 Самовывоз ≤{av_rad} км + 🚚 Доставка"

    if not pool:
        extra_tip = "\nПопробуйте переключить «📍 Самовывоз / 🚚 Доставка / 🌐 Всё» или запустить новый поиск." if shop_filter == "Авито" else "\nНажмите «🔍 Поиск новых цен»."
        return f"📋 <b>{title_hdr}</b> (📍 <b>{_h(cur_reg)}</b>{av_note}): пока нет доступных позиций.{extra_tip}"

    has_y = any(x.status_code == "YELLOW" for x in pool)
    mode_label = "🟢 + 🟡 выгодные с нюансом" if (greens and has_y) else ("🟢 только ЗЕЛЁНЫЕ" if greens else "🟡 проверенные предложения")
    lines = [
        f"🏆 <b>Лучшие предложения ({mode_label}) — {title_hdr}</b>\n"
        f"📍 <b>Регион:</b> {_h(cur_reg)}{av_note} · найдено: <b>{len(pool)} шт.</b>\n"
    ]
    for idx, it in enumerate(pool[:12], start=1):
        price_head = f"Итого {it.price:,} ₽" if it.customs_duty > 0 else f"{it.price:,} ₽"
        duty_info = f" <i>({it.base_price:,} ₽ + пошлина {it.customs_duty:,} ₽)</i>" if it.customs_duty > 0 else " <i>(пошлина 0 ₽)</i>"
        st_icon = "🟢" if it.status_code == "GREEN" else "🟡"
        link_text = (
            "Открыть на AliExpress (🔥 скидка!) ↗"
            if it.shop == "AliExpress"
            else f"Открыть карточку ({_h(it.shop)}) ↗"
        )
        loc_raw = (it.location_delivery or "").strip()
        loc_clean = _h(loc_raw[:68])
        loc_line = loc_clean if loc_raw.startswith(("📍", "🚚")) else f"📍 {loc_clean}"
        lines.append(
            f"{idx}. {st_icon} <b>[{_h(it.shop)}] {price_head}</b>{duty_info}\n"
            f"   {_h(it.title[:68])}\n"
            f"   {loc_line}\n"
            f"   👤 {_h(it.seller_name[:28])} | 🔗 <a href=\"{_h(clean_product_url(it.url))}\">{link_text}</a>"
        )
    return "\n\n".join(lines)


def format_settings_tg(cfg: dict, cid: str = "") -> str:
    """Компактный блок текущих персональных настроек и статуса."""
    _sync_query_lists(cfg)
    queries = cfg.get("queries") or []
    sel_set = set(cfg.get("selected_queries") or [])
    muted_set = set(cfg.get("muted_queries") or [])
    if queries:
        q_lines = []
        for idx, q in enumerate(queries, 1):
            s_ic = "✅" if q in sel_set else "⬜"
            n_ic = "🔕" if q in muted_set else "🔔"
            q_lines.append(f"  {idx}. {s_ic}{n_ic} <code>{_h(q)}</code>")
        q_list = "\n".join(q_lines)
    else:
        q_list = "  • <i>Пока не выбран — просто напишите название любого товара в чат!</i>"

    enabled_shops = _get_shops_enabled_map(cfg)
    shops_str = ", ".join(
        f"{'✅' if enabled_shops.get(k, True) else '⬜'} {name}"
        for k, name, _ in SHOPS_META
    )
    city_str = _h(cfg.get("detected_place") or cfg.get("region", "Все регионы РФ"))
    has_geo = cfg.get("lat") is not None and cfg.get("lon") is not None
    av_rad = max(5, min(int(cfg.get("avito_radius_km", 15) or 15), 15))
    av_mode = str(cfg.get("avito_mode", "both") or "both").lower()
    if av_mode == "pickup":
        avito_rule = f"только самовывоз ≤{av_rad} км"
    elif av_mode == "delivery":
        avito_rule = "только Авито Доставка"
    else:
        avito_rule = f"самовывоз ≤{av_rad} км {'📌' if has_geo else ''} + Доставка"
    status_str = f"🔄 Идёт поиск ({_h(rt.current_step)})" if rt.is_scanning else "✅ Готов к поиску"
    source_items, _ = _get_scoped_items(cid, cfg)
    g_cnt = sum(1 for i in source_items if i.status_code == "GREEN" and str(i.in_stock).startswith("✅"))
    y_cnt = sum(1 for i in source_items if i.status_code == "YELLOW" and str(i.in_stock).startswith("✅"))
    r_cnt = sum(1 for i in source_items if i.status_code == "RED" or str(i.in_stock).startswith("❌"))
    price_note = "ручной фильтр" if cfg.get("price_is_custom") else "авто-подбор по товару"
    return (
        f"⚙️ <b>Текущие настройки (v{_h(BOT_VERSION)})</b>\n\n"
        f"📍 <b>Регион поиска:</b> ✅ <b>{city_str}</b> <i>({avito_rule.strip()})</i>\n"
        f"🏬 <b>Площадки (до {len(SHOPS_META)} на 1 позицию):</b> {shops_str}\n"
        f"🎯 <b>Мои товары ({len(queries)} шт.):</b>\n{q_list}\n\n"
        f"💰 <b>Диапазон цен:</b> от <b>{cfg.get('min_price', 100):,}</b> до <b>{cfg.get('max_price', 250000):,} ₽</b> <i>({price_note})</i>\n"
        f"⏱ <b>Интервал автопроверки:</b> {format_interval_str(cfg.get('interval_minutes', 180), bool(cfg.get('use_ladder', False)))}\n"
        f"📊 <b>Позиций в Excel:</b> {len(source_items)} шт. (🟢 {g_cnt} / 🟡 {y_cnt} / 🔴 {r_cnt}) — <i>{rt.last_updated if source_items else 'ожидание скана'}</i>\n"
        f"📡 <b>Статус:</b> {status_str} <i>(/changelog — что нового)</i>\n\n"
        f"✍️ <b>Напишите название товара</b> или нажмите <b>«🔍 Поиск новых цен»</b> внизу 👇"
    )


def format_welcome_blocks_tg(cfg: dict, cid: str = "") -> List[str]:
    """Разбивает приветствие и описание бота на 4 отдельные смысловые карточки."""
    _sync_query_lists(cfg)
    enabled_shops = _get_shops_enabled_map(cfg)
    shop_lines = [
        f"  {'✅' if enabled_shops.get(k, True) else '⬜'} {label}"
        for k, _, label in SHOPS_META
    ]
    block_1 = (
        f"👋 <b>Добро пожаловать! Мониторинг минимальных цен 24/7 (v{_h(BOT_VERSION)})</b>\n\n"
        f"Бот автоматически ищет минимальные реальные цены на нужные товары и сравнивает предложения по схеме:\n"
        f"🏬 <b>На 1 позицию — все {len(SHOPS_META)} магазинов:</b>\n"
        + "\n".join(shop_lines)
    )
    block_2 = (
        f"🔎 <b>Блок 1. Как работает поиск и сверка карточек</b>\n\n"
        f"• При нажатии <b>«🔍 Поиск новых цен»</b> каждая позиция по очереди проходит все {len(SHOPS_META)} площадок.\n"
        f"• В живом окне виден точный прогресс: <code>🔄</code> идёт поиск → <code>✅ N поз.</code> найдено → <code>❌ 0 поз.</code> пусто, а возле номера позиции ставится общая отметка.\n"
        f"• Все найденные товары перепроверяются по реальным карточкам: считывается точная <b>внутренняя пошлина с сайта</b>, отсекаются рассрочки и фейки.\n"
        f"• По итогам приходит рейтинг <b>⚖️ Цена / Качество</b> и полная таблица <code>Мониторинг_Товаров.xlsx</code>."
    )
    block_3 = (
        f"🧭 <b>Блок 2. Товары, геолокация и автоуведомления</b>\n\n"
        f"• <b>Гибкий список товаров:</b> отмечайте нужные ✅/⬜ для поиска, включайте алерты 🔔/🔕 по любой позиции и удаляйте ❌ лишнее.\n"
        f"• <b>Мгновенная геолокация 📱 и Авито:</b> определяет ваш город, позволяет выбрать самовывоз <b>5–15 км</b> или доставку в ваш город.\n"
        f"• <b>Умный ценовой фильтр и расписание:</b> авто-подбор ОТ/ДО под категорию и стабильное автообновление <b>каждые 3 часа</b>."
    )
    block_4 = format_settings_tg(cfg, cid)
    return [block_1, block_2, block_3, block_4]


async def tg_send_welcome_blocks(client: httpx.AsyncClient, token: str, cid: str, cfg: dict) -> None:
    """Отправляет разделённое описание в виде нескольких отдельных сообщений с клавиатурой на последнем."""
    blocks = format_welcome_blocks_tg(cfg, cid)
    last_idx = len(blocks) - 1
    for idx, txt in enumerate(blocks):
        rm = TG_MAIN_KEYBOARD if idx == last_idx else None
        await tg_send_message(client, token, cid, txt, reply_markup=rm)
        if idx < last_idx:
            await asyncio.sleep(0.15)


def format_scan_progress_tg(
    *,
    step_label: str,
    city_name: str,
    min_price: int,
    max_price: int,
    price_is_custom: bool,
    pos_states: List[dict],
    done: bool = False,
) -> str:
    """Формирует живой отчёт «1 позиция → 6 магазинов» с символами 🔄, ⏳, ✅, ❌."""
    header = (
        "✅ <b>Поиск завершён! Карточки проверены, таблица Excel обновлена.</b>"
        if done
        else f"🔄 <b>Идёт поиск минимальных цен и сверка карточек:</b>\n<i>{_h(step_label)}</i>"
    )
    price_str = f"{min_price:,} – {max_price:,} ₽" if price_is_custom else "авто по категории товара"
    lines = [
        header,
        "",
        f"📍 <b>Регион:</b> {_h(city_name)} <i>(Авито: самовывоз ≤15 км + Доставка)</i>",
        f"🏬 <b>Схема:</b> по <b>{len(SHOPS_META)} магазинов</b> на 1 позицию",
        f"💰 <b>Диапазон:</b> {_h(price_str)}",
        "",
    ]
    detailed_all = len(pos_states) <= 3
    for idx, p_st in enumerate(pos_states, 1):
        q = p_st["query"]
        states = p_st["states"]
        counts = p_st["counts"]
        active_vals = [v for v in states.values() if v != "off"]
        has_running = any(v == "running" for v in active_vals)
        all_finished = all(v == "done" for v in active_vals)
        pos_total = sum(counts.values())

        if all_finished:
            q_badge = f"✅ <b>{_h(q)}</b> — <i>всего {pos_total} поз.</i>" if pos_total > 0 else f"❌ <b>{_h(q)}</b> — <i>0 поз.</i>"
        elif has_running:
            q_badge = f"🔄 <b>{_h(q)}</b> — <i>идёт поиск…</i>"
        else:
            q_badge = f"⏳ <b>{_h(q)}</b> — <i>в очереди</i>"

        lines.append(f"<b>{idx}.</b> {q_badge}")
        if detailed_all or has_running or done:
            for s_idx, (sk_key, sh_name, _) in enumerate(SHOPS_META, 1):
                st = states.get(sk_key, "queued")
                cnt = counts.get(sk_key, 0)
                if st == "off":
                    st_str = "⬜ отключён"
                elif st == "running":
                    st_str = "🔄 <b>идёт поиск…</b>"
                elif st == "done":
                    st_str = f"✅ <b>{cnt} поз.</b>" if cnt > 0 else "❌ <b>0 поз.</b>"
                else:
                    st_str = "⏳ ожидание"
                lines.append(f"  {s_idx}) <b>{_h(sh_name)}:</b> {st_str}")
        else:
            compact = []
            for sk_key, sh_name, _ in SHOPS_META:
                st = states.get(sk_key, "queued")
                cnt = counts.get(sk_key, 0)
                short_n = {"Яндекс Маркет": "ЯМ", "Wildberries": "WB", "AliExpress": "Ali"}.get(sh_name, sh_name)
                if st == "off":
                    compact.append(f"⬜ {short_n}")
                elif st == "done":
                    compact.append(f"{'✅' if cnt > 0 else '❌'} {short_n} {cnt}")
                else:
                    compact.append(f"⏳ {short_n}")
            lines.append("  " + " | ".join(compact))
    return "\n".join(lines)


def _get_user_effective_cfg(base_cfg: dict, cid: str) -> dict:
    """Возвращает персональные настройки пользователя cid (с заполнением из base_cfg при первом входе)."""
    is_new = not rt.user_store.exists(cid)
    u = rt.user_store.touch(cid)
    changed = False
    if is_new and not u.get("queries"):
        init_q = [str(x).strip() for x in (base_cfg.get("queries") or ["кресло компьютерное"]) if str(x).strip()][:MAX_PRODUCTS]
        u["queries"] = init_q
        init_sel = base_cfg.get("selected_queries")
        u["selected_queries"] = [q for q in init_q if q in set(init_sel)] if isinstance(init_sel, list) and init_sel else list(init_q)
        u["min_price"] = int(base_cfg.get("min_price", 100))
        u["max_price"] = int(base_cfg.get("max_price", 250000))
        u["price_is_custom"] = bool(base_cfg.get("price_is_custom", False))
        u["region"] = base_cfg.get("region", "Все регионы РФ")
        u["region_slug_avito"] = base_cfg.get("region_slug_avito", "all")
        u["ymarket_lr"] = int(base_cfg.get("ymarket_lr", 213))
        u["avito_mode"] = str(base_cfg.get("avito_mode", "both") or "both")
        u["avito_radius_km"] = int(base_cfg.get("avito_radius_km", 15) or 15)
        raw_iv = base_cfg.get("interval_minutes", 180)
        u["interval_minutes"] = raw_iv if is_auto_interval(raw_iv) else int(raw_iv)
        u["use_ladder"] = is_auto_interval(raw_iv, bool(base_cfg.get("use_ladder", False)))
        u["ladder_step"] = int(base_cfg.get("ladder_step", 0)) % len(LADDER_INTERVALS)
        base_shops = base_cfg.get("shops_priority") or {}
        u["shops_priority"] = {k: bool(base_shops.get(k, True)) for k, _, _ in SHOPS_META}
        changed = True
    if not isinstance(u.get("shops_priority"), dict):
        u["shops_priority"] = {k: True for k, _, _ in SHOPS_META}
        changed = True
    _sync_query_lists(u)
    if changed:
        rt.user_store.save()
    return u


async def _ensure_user_update_news(
    client: httpx.AsyncClient,
    token: str,
    cid: str,
    u_cfg: dict,
    *,
    skip_if_changelog: bool = False,
) -> bool:
    """Гарантированно отправляет новости текущей версии BOT_VERSION пользователю ровно 1 раз."""
    if not cid or not token:
        return False
    prev_ver = str(u_cfg.get("last_announced_version") or "").strip()
    if prev_ver == BOT_VERSION:
        return False
    u_cfg["last_announced_version"] = BOT_VERSION
    if skip_if_changelog:
        rt.user_store.save()
        return False
    mid_sent = await tg_send_message(client, token, cid, format_update_news_tg(), reply_markup=TG_MAIN_KEYBOARD)
    if mid_sent is None:
        u_cfg["last_announced_version"] = prev_ver
        return False
    _set_announced_version(BOT_VERSION)
    return True


async def telegram_bot_loop() -> None:
    """
    Постоянный фоновый обработчик публичного Telegram-бота.
    У каждого пользователя (chat_id) — свои настройки товара, цены, города и уведомлений.
    """
    await asyncio.sleep(0.3)
    await rt.user_store.load()
    saved_upd_id = int((rt.user_store.meta or {}).get("last_update_id") or 0)
    if saved_upd_id > 0:
        rt.last_update_id = max(rt.last_update_id, saved_upd_id)
    cfg = load_config()
    token = str(cfg.get("telegram", {}).get("bot_token") or "").strip()
    init_cid = str(cfg.get("telegram", {}).get("chat_id") or "").strip()

    if token:
        async with httpx.AsyncClient(timeout=15.0, trust_env=False) as client:
            try:
                await client.post(
                    f"https://api.telegram.org/bot{token}/setMyCommands",
                    json={
                        "commands": [
                            {"command": "start", "description": "👋 Меню и настройки"},
                            {"command": "items", "description": "📦 Мои товары (✅/🔔/✏️/❌)"},
                            {"command": "replace", "description": "🔄 Заменить один товар в списке"},
                            {"command": "compare", "description": "⚖️ Сравнение Цена / Качество"},
                            {"command": "city", "description": "📍 Сменить город / геолокация"},
                            {"command": "shops", "description": "🏬 Выбор магазинов (✅/⬜)"},
                            {"command": "price", "description": "💰 Задать цену (ОТ и ДО)"},
                            {"command": "notify", "description": "🔔 Уведомления и авторежим"},
                            {"command": "scan", "description": "🔍 Запустить поиск сейчас"},
                            {"command": "top", "description": "📋 Лучшие цены по магазинам"},
                            {"command": "excel", "description": "📥 Скачать таблицу Excel (.xlsx)"},
                            {"command": "changelog", "description": "🆕 Новости обновления бота"},
                        ]
                    },
                )
            except Exception:
                pass
            try:
                await client.post(
                    f"https://api.telegram.org/bot{token}/setMyShortDescription",
                    json={
                        "short_description": "Мониторинг минимальных цен 24/7: 6 маркетплейсов на позицию, анти-хлам фильтр, пошлины и Excel.",
                    },
                )
                await client.post(
                    f"https://api.telegram.org/bot{token}/setMyDescription",
                    json={
                        "description": (
                            "🔎 Поиск минимальных цен 24/7 — 6 маркетплейсов на каждую позицию:\n"
                            "Ozon, AliExpress, Яндекс Маркет, Авито, DNS и Wildberries.\n\n"
                            "• Живой трекер по каждой позиции (🔄 → ✅ / ❌) и сверка карточек.\n"
                            "• Умный анти-хлам отсев деталей/расходников + пошлины прямо с карточек.\n"
                            "• Замена 1 товара или списка, геолокация Авито (5–15 км) и Excel (.xlsx)."
                        ),
                    },
                )
            except Exception:
                pass
            has_pending = False
            try:
                start_params = {"timeout": 1}
                if rt.last_update_id > 0:
                    start_params["offset"] = rt.last_update_id + 1
                upds = (await client.get(f"https://api.telegram.org/bot{token}/getUpdates", params=start_params)).json().get("result", [])
                if upds:
                    fresh_cutoff = time.time() - 90.0
                    for u in upds:
                        uid_val = int(u.get("update_id", 0))
                        msg = u.get("message") or (u.get("callback_query") or {}).get("message") or {}
                        cid = str(msg.get("chat", {}).get("id") or "")
                        if cid:
                            rt.active_chat_ids.add(cid)
                            save_subscribers(rt.active_chat_ids)
                            _get_user_effective_cfg(cfg, cid)
                        u_ts = float((u.get("message") or {}).get("date") or 0.0)
                        if rt.last_update_id > 0:
                            has_pending = True
                        elif u.get("callback_query") or u_ts < fresh_cutoff:
                            rt.last_update_id = max(rt.last_update_id, uid_val)
                        else:
                            has_pending = True
                    if rt.last_update_id > saved_upd_id:
                        rt.user_store.meta["last_update_id"] = rt.last_update_id
                        rt.user_store.save()
            except Exception:
                pass

            if _get_announced_version() != BOT_VERSION:
                b_recipients = _collect_broadcast_recipients(cfg)
                if b_recipients:
                    news_msg = format_update_news_tg()
                    sent_any = False
                    for r_cid in b_recipients:
                        u_r = _get_user_effective_cfg(cfg, r_cid)
                        if str(u_r.get("last_announced_version") or "").strip() == BOT_VERSION:
                            continue
                        mid_sent = await tg_send_message(client, token, r_cid, news_msg, reply_markup=TG_MAIN_KEYBOARD)
                        if mid_sent is not None:
                            u_r["last_announced_version"] = BOT_VERSION
                            sent_any = True
                        await asyncio.sleep(0.2)
                    if sent_any:
                        _set_announced_version(BOT_VERSION)

            if init_cid and not has_pending:
                u_init = _get_user_effective_cfg(cfg, init_cid)
                if not u_init.get("welcomed"):
                    u_init["welcomed"] = True
                    rt.user_store.save()
                    await tg_send_welcome_blocks(client, token, init_cid, u_init)

    async with httpx.AsyncClient(
        timeout=20.0,
        limits=httpx.Limits(max_keepalive_connections=10, max_connections=20, keepalive_expiry=60.0),
    ) as client:
        while True:
            cfg = load_config()
            token = str(cfg.get("telegram", {}).get("bot_token") or "").strip()
            if not token:
                await asyncio.sleep(2.0)
                continue

            try:
                params = {"timeout": 8}
                if rt.last_update_id:
                    params["offset"] = rt.last_update_id + 1
                r = await client.get(f"https://api.telegram.org/bot{token}/getUpdates", params=params)
                updates = r.json().get("result", [])
                prev_upd_id = rt.last_update_id

                for upd in updates:
                    rt.last_update_id = max(rt.last_update_id, int(upd.get("update_id", 0)))

                    # 1. Обработка нажатий на инлайн-кнопки
                    cb = upd.get("callback_query")
                    if cb:
                        cb_id = cb.get("id")
                        cb_data = cb.get("data") or ""
                        cid = str(cb.get("message", {}).get("chat", {}).get("id") or "")
                        mid = cb.get("message", {}).get("message_id")
                        if cid:
                            rt.active_chat_ids.add(cid)
                            save_subscribers(rt.active_chat_ids)
                        u_cfg = _get_user_effective_cfg(cfg, cid) if cid else cfg
                        if cid:
                            await _ensure_user_update_news(client, token, cid, u_cfg)
                        _cb_answered = False

                        async def _ack_cb(toast: str = "") -> None:
                            nonlocal _cb_answered
                            if _cb_answered or not cb_id:
                                return
                            _cb_answered = True
                            payload_ack = {"callback_query_id": cb_id}
                            if toast:
                                payload_ack["text"] = toast[:190]
                            try:
                                await client.post(
                                    f"https://api.telegram.org/bot{token}/answerCallbackQuery",
                                    json=payload_ack,
                                    timeout=4.0,
                                )
                            except Exception:
                                pass

                        if cb_data == "noop":
                            await _ack_cb()
                            continue

                        elif cb_data.startswith("citypage:"):
                            await _ack_cb()
                            new_p = int(cb_data.split(":", 1)[1])
                            if mid:
                                try:
                                    await client.post(
                                        f"https://api.telegram.org/bot{token}/editMessageReplyMarkup",
                                        json={
                                            "chat_id": cid,
                                            "message_id": mid,
                                            "reply_markup": build_cities_markup(
                                                new_p,
                                                u_cfg.get("region", "Все регионы РФ"),
                                                str(u_cfg.get("avito_mode", "both")),
                                                int(u_cfg.get("avito_radius_km", 15) or 15),
                                            ),
                                        },
                                        timeout=6.0,
                                    )
                                except Exception:
                                    pass
                            continue

                        elif cb_data.startswith("qsel:"):
                            q_idx = int(cb_data.split(":", 1)[1])
                            _sync_query_lists(u_cfg)
                            queries = u_cfg.get("queries", [])
                            if 0 <= q_idx < len(queries):
                                t_q = queries[q_idx]
                                sel = list(u_cfg.get("selected_queries", []))
                                if t_q in sel:
                                    sel = [x for x in sel if x != t_q]
                                    await _ack_cb(f"⬜ Снят с поиска: {t_q[:30]}")
                                else:
                                    sel.append(t_q)
                                    await _ack_cb(f"✅ Выбран для поиска: {t_q[:30]}")
                                u_cfg["selected_queries"] = [q for q in queries if q in set(sel)]
                                rt.user_store.save()
                                if mid:
                                    await tg_edit_message(client, token, cid, mid, format_products_menu_text(u_cfg), reply_markup=build_products_markup(u_cfg))
                            else:
                                await _ack_cb()
                            continue

                        elif cb_data.startswith("qtog:"):
                            q_idx = int(cb_data.split(":", 1)[1])
                            _sync_query_lists(u_cfg)
                            queries = u_cfg.get("queries", [])
                            if 0 <= q_idx < len(queries):
                                t_q = queries[q_idx]
                                muted = set(u_cfg.get("muted_queries", []))
                                if t_q in muted:
                                    muted.remove(t_q)
                                    await _ack_cb(f"🔔 Уведомления включены: {t_q[:28]}")
                                else:
                                    muted.add(t_q)
                                    await _ack_cb(f"🔕 Уведомления отключены: {t_q[:28]}")
                                u_cfg["muted_queries"] = sorted(list(muted))
                                cfg["muted_queries"] = u_cfg["muted_queries"]
                                rt.user_store.save()
                                save_config(cfg)
                                if mid:
                                    await tg_edit_message(client, token, cid, mid, format_products_menu_text(u_cfg), reply_markup=build_products_markup(u_cfg))
                            else:
                                await _ack_cb()
                            continue

                        elif cb_data.startswith("qdel:"):
                            q_idx = int(cb_data.split(":", 1)[1])
                            _sync_query_lists(u_cfg)
                            queries = list(u_cfg.get("queries", []))
                            if 0 <= q_idx < len(queries):
                                removed = queries.pop(q_idx)
                                u_cfg["queries"] = queries
                                u_cfg["selected_queries"] = [x for x in u_cfg.get("selected_queries", []) if x != removed]
                                u_cfg["muted_queries"] = [x for x in u_cfg.get("muted_queries", []) if x != removed]
                                cfg["queries"] = queries
                                cfg["selected_queries"] = u_cfg["selected_queries"]
                                cfg["muted_queries"] = u_cfg["muted_queries"]
                                _invalidate_user_cache(cid, clear_global=True)
                                rt.user_store.save()
                                save_config(cfg)
                                await _ack_cb(f"❌ Удалён: {removed[:30]}")
                                if mid:
                                    await tg_edit_message(client, token, cid, mid, format_products_menu_text(u_cfg), reply_markup=build_products_markup(u_cfg))
                            else:
                                await _ack_cb()
                            continue

                        elif cb_data.startswith("qrep:"):
                            q_idx = int(cb_data.split(":", 1)[1])
                            _sync_query_lists(u_cfg)
                            queries = list(u_cfg.get("queries", []))
                            if 0 <= q_idx < len(queries):
                                old_q = queries[q_idx]
                                await _ack_cb(f"✏️ Заменяем #{q_idx + 1}")
                                rt.tg_dialog_mode[cid] = f"WAIT_REPLACE_PRODUCT:{q_idx}"
                                await tg_send_message(
                                    client,
                                    token,
                                    cid,
                                    f"✏️ <b>Замена позиции #{q_idx + 1}:</b> <code>{_h(old_q)}</code>\n\n"
                                    f"✍️ Напишите новое название товара (или несколько через запятую):\n"
                                    f"<i>Остальные позиции в списке и Excel сохранятся, обновится только выбранный слот, и сразу запустится поиск.</i>",
                                    reply_markup=TG_MAIN_KEYBOARD,
                                )
                            else:
                                await _ack_cb("⚠️ Позиция не найдена")
                            continue

                        elif cb_data == "pq:ALL":
                            await _ack_cb("🌐 Показаны все товары")
                            pq_txt = format_price_quality_comparison_tg(None, cid, max_rows=10, query_filter="ALL")
                            pq_kb = build_pq_markup(u_cfg, "ALL")
                            if mid and len(pq_txt) <= 3850:
                                await tg_edit_message(client, token, cid, mid, pq_txt, reply_markup=pq_kb)
                            else:
                                await tg_send_html_chunks(client, token, cid, pq_txt, reply_markup=pq_kb)
                            continue

                        elif cb_data.startswith("pqq:"):
                            q_idx = int(cb_data.split(":", 1)[1])
                            _sync_query_lists(u_cfg)
                            queries = list(u_cfg.get("queries", []))
                            if 0 <= q_idx < len(queries):
                                target_q = queries[q_idx]
                                await _ack_cb(f"🎯 {target_q[:28]}")
                                pq_txt = format_price_quality_comparison_tg(None, cid, max_rows=10, query_filter=target_q)
                                pq_kb = build_pq_markup(u_cfg, target_q)
                                if mid and len(pq_txt) <= 3850:
                                    await tg_edit_message(client, token, cid, mid, pq_txt, reply_markup=pq_kb)
                                else:
                                    await tg_send_html_chunks(client, token, cid, pq_txt, reply_markup=pq_kb)
                            else:
                                await _ack_cb()
                            continue

                        elif cb_data.startswith("qact:"):
                            act = cb_data.split(":", 1)[1]
                            if act == "open":
                                await _ack_cb()
                                await tg_send_message(client, token, cid, format_products_menu_text(u_cfg), reply_markup=build_products_markup(u_cfg))
                            elif act == "add":
                                await _ack_cb()
                                _sync_query_lists(u_cfg)
                                if len(u_cfg.get("queries", [])) >= MAX_PRODUCTS:
                                    await tg_send_message(
                                        client,
                                        token,
                                        cid,
                                        f"⚠️ Достигнут максимум <b>{MAX_PRODUCTS} товаров</b>. Нажмите <b>✏️</b> чтобы заменить любой товар или <b>🗑 Убрать снятые ⬜</b>.",
                                        reply_markup=build_products_markup(u_cfg),
                                    )
                                else:
                                    rt.tg_dialog_mode[cid] = "WAIT_ADD_PRODUCT"
                                    await tg_send_message(
                                        client,
                                        token,
                                        cid,
                                        f"➕ <b>Напишите название товара, который хотите добавить в список</b> (можно несколько через запятую):\n"
                                        f"<i>Товар добавится, отметится ✅ для поиска и 🔔 для уведомлений, и сразу запустится проверка цен.</i>",
                                        reply_markup=TG_MAIN_KEYBOARD,
                                    )
                            elif act == "replace_menu":
                                await _ack_cb()
                                _sync_query_lists(u_cfg)
                                queries = list(u_cfg.get("queries", []))
                                if len(queries) >= 2:
                                    await tg_send_message(
                                        client,
                                        token,
                                        cid,
                                        f"🔄 <b>Какой именно товар заменить?</b> <i>(сейчас в списке {len(queries)} шт.)</i>\n"
                                        f"Нажмите на нужную позицию — поменяется только она, а остальные останутся в Excel:",
                                        reply_markup=build_replace_pick_markup(u_cfg),
                                    )
                                else:
                                    rt.tg_dialog_mode[cid] = "WAIT_PRODUCT"
                                    await tg_send_message(
                                        client,
                                        token,
                                        cid,
                                        "✍️ <b>Напишите новое название товара</b> (или несколько через запятую — список обновится и запустится поиск):",
                                        reply_markup=TG_MAIN_KEYBOARD,
                                    )
                            elif act == "replace_all":
                                await _ack_cb()
                                rt.tg_dialog_mode[cid] = "WAIT_PRODUCT"
                                await tg_send_message(
                                    client,
                                    token,
                                    cid,
                                    "🆕 <b>Напишите новый список товаров</b> (один или несколько через запятую / с новой строки):\n"
                                    "<i>Старый список полностью заменится на новый, лишние старые товары не попадут в Excel.</i>",
                                    reply_markup=TG_MAIN_KEYBOARD,
                                )
                            elif act == "clear_off":
                                _sync_query_lists(u_cfg)
                                sel_s = set(u_cfg.get("selected_queries", []))
                                queries = list(u_cfg.get("queries", []))
                                if not sel_s:
                                    await _ack_cb("⚠️ Сначала отметьте ✅ хотя бы 1 товар")
                                else:
                                    kept = [q for q in queries if q in sel_s]
                                    removed_cnt = len(queries) - len(kept)
                                    u_cfg["queries"] = kept
                                    u_cfg["selected_queries"] = list(kept)
                                    u_cfg["muted_queries"] = [q for q in u_cfg.get("muted_queries", []) if q in set(kept)]
                                    cfg["queries"] = kept
                                    cfg["selected_queries"] = list(kept)
                                    cfg["muted_queries"] = u_cfg["muted_queries"]
                                    _invalidate_user_cache(cid, clear_global=True)
                                    rt.user_store.save()
                                    save_config(cfg)
                                    await _ack_cb(f"🗑 Убрано снятых: {removed_cnt}")
                                    if mid:
                                        await tg_edit_message(client, token, cid, mid, format_products_menu_text(u_cfg), reply_markup=build_products_markup(u_cfg))
                            elif act == "all":
                                _sync_query_lists(u_cfg)
                                u_cfg["selected_queries"] = list(u_cfg.get("queries", []))
                                rt.user_store.save()
                                await _ack_cb("✅ Выбраны все товары")
                                if mid:
                                    await tg_edit_message(client, token, cid, mid, format_products_menu_text(u_cfg), reply_markup=build_products_markup(u_cfg))
                            elif act == "unmute_all":
                                u_cfg["muted_queries"] = []
                                cfg["muted_queries"] = []
                                rt.user_store.save()
                                save_config(cfg)
                                await _ack_cb("🔔 Все уведомления включены")
                                if mid:
                                    await tg_edit_message(client, token, cid, mid, format_products_menu_text(u_cfg), reply_markup=build_products_markup(u_cfg))
                            elif act == "compare":
                                await _ack_cb()
                                pq_txt = format_price_quality_comparison_tg(None, cid, max_rows=10, query_filter="ALL")
                                await tg_send_html_chunks(client, token, cid, pq_txt, reply_markup=build_pq_markup(u_cfg, "ALL"))
                            elif act == "refresh_now":
                                city_disp = str(u_cfg.get("region") or "Все регионы РФ")
                                await _ack_cb(f"🔄 Обновляю цены для: {city_disp}")
                                await tg_send_message(
                                    client,
                                    token,
                                    cid,
                                    f"🔄 <b>Запущено обновление цен и пошлин</b> для региона: <b>{_h(city_disp)}</b>...",
                                    reply_markup=TG_MAIN_KEYBOARD,
                                )
                                trigger_immediate_scan(cid)
                            else:
                                await _ack_cb()
                            continue

                        elif cb_data == "menu:shops_filter":
                            await _ack_cb()
                            txt = (
                                "🏬 <b>Выбор магазинов для поиска и сравнения</b>\n"
                                "<i>Нажимайте на площадки ниже, чтобы поставить или снять зелёную галочку ✅:</i>"
                            )
                            if mid:
                                await tg_edit_message(client, token, cid, mid, txt, reply_markup=build_shops_filter_markup(u_cfg))
                            else:
                                await tg_send_message(client, token, cid, txt, reply_markup=build_shops_filter_markup(u_cfg))
                            continue

                        elif cb_data.startswith("toggleshop:"):
                            sk = cb_data.split(":", 1)[1]
                            sp = dict(_get_shops_enabled_map(u_cfg))
                            if sk == "ALL":
                                for k, _, _ in SHOPS_META:
                                    sp[k] = True
                                await _ack_cb("✅ Включены все 6 магазинов")
                            elif sk in sp:
                                if sp[sk] and sum(1 for v in sp.values() if v) <= 1:
                                    await _ack_cb("⚠️ Должен быть выбран хотя бы 1 магазин")
                                    continue
                                sp[sk] = not sp[sk]
                                s_title = next((nm for k, nm, _ in SHOPS_META if k == sk), sk)
                                await _ack_cb(f"{'✅ Включён' if sp[sk] else '⬜ Выключен'}: {s_title}")
                            u_cfg["shops_priority"] = sp
                            cfg["shops_priority"] = sp
                            _invalidate_user_cache(cid, clear_global=True)
                            rt.user_store.save()
                            save_config(cfg)
                            if mid:
                                try:
                                    await client.post(
                                        f"https://api.telegram.org/bot{token}/editMessageReplyMarkup",
                                        json={
                                            "chat_id": cid,
                                            "message_id": mid,
                                            "reply_markup": build_shops_filter_markup(u_cfg),
                                        },
                                        timeout=6.0,
                                    )
                                except Exception:
                                    pass
                            continue

                        elif cb_data == "action:auto_price":
                            await _ack_cb("🤖 Включён авто-подбор цен")
                            u_cfg["price_is_custom"] = False
                            act_q = _active_queries(u_cfg) or ["товар"]
                            p_min, p_max = _auto_price_bounds(act_q)
                            u_cfg["min_price"], u_cfg["max_price"] = p_min, p_max
                            cfg["price_is_custom"] = False
                            cfg["min_price"], cfg["max_price"] = p_min, p_max
                            _invalidate_user_cache(cid, clear_global=True)
                            rt.user_store.save()
                            save_config(cfg)
                            rt.tg_dialog_mode.pop(cid, None)
                            await tg_send_message(
                                client,
                                token,
                                cid,
                                f"🤖 <b>Включён авто-подбор диапазона цены индивидуально для каждого товара!</b>\n⚡ Запускаю поиск...",
                                reply_markup=TG_MAIN_KEYBOARD,
                            )
                            trigger_immediate_scan(cid)
                            continue

                        elif cb_data.startswith("interval:"):
                            await _ack_cb()
                            new_int_raw = cb_data.split(":", 1)[1]
                            if new_int_raw in ("ladder", "auto"):
                                u_cfg["interval_minutes"] = "ladder"
                                u_cfg["use_ladder"] = True
                                u_cfg["ladder_step"] = 0
                                cfg["interval_minutes"] = "ladder"
                                cfg["use_ladder"] = True
                                ack_text = (
                                    "🔄 <b>Включён циклический авторежим уведомлений:</b>\n"
                                    "• 1-й шаг: через <b>3 часа</b>\n"
                                    "• 2-й шаг: через <b>1 день (24 часа)</b>\n"
                                    "• Далее цикл повторяется: <b>3 часа ➔ 1 день ➔ 3 часа ➔ 1 день</b>\n\n"
                                    "<i>(При новом поиске цикл стартует с 3 часов)</i>"
                                )
                            else:
                                new_int = int(new_int_raw)
                                u_cfg["interval_minutes"] = new_int
                                u_cfg["use_ladder"] = False
                                cfg["interval_minutes"] = new_int
                                cfg["use_ladder"] = False
                                ack_text = f"⏱ <b>Интервал автопроверки установлен: {format_interval_str(new_int)}!</b>"
                            rt.user_store.save()
                            save_config(cfg)
                            if mid:
                                await tg_edit_message(client, token, cid, mid, format_notify_settings_text(u_cfg), reply_markup=build_notify_markup(u_cfg))
                            await tg_send_message(client, token, cid, ack_text, reply_markup=TG_MAIN_KEYBOARD)
                            continue

                        elif cb_data.startswith("notify_mode:"):
                            await _ack_cb()
                            new_mode = cb_data.split(":", 1)[1]
                            u_cfg["alert_filter"] = new_mode
                            cfg["alert_filter"] = new_mode
                            rt.user_store.save()
                            save_config(cfg)
                            if mid:
                                await tg_edit_message(client, token, cid, mid, format_notify_settings_text(u_cfg), reply_markup=build_notify_markup(u_cfg))
                            continue

                        elif cb_data.startswith("toggle_q:"):
                            q_idx = int(cb_data.split(":", 1)[1])
                            queries = u_cfg.get("queries", [])
                            if 0 <= q_idx < len(queries):
                                target_q = queries[q_idx]
                                muted = set(u_cfg.get("muted_queries", []))
                                if target_q in muted:
                                    muted.remove(target_q)
                                    await _ack_cb(f"🔔 Вкл: {target_q[:30]}")
                                else:
                                    muted.add(target_q)
                                    await _ack_cb(f"🔕 Выкл: {target_q[:30]}")
                                u_cfg["muted_queries"] = sorted(list(muted))
                                cfg["muted_queries"] = u_cfg["muted_queries"]
                                rt.user_store.save()
                                save_config(cfg)
                                if mid:
                                    await tg_edit_message(client, token, cid, mid, format_notify_settings_text(u_cfg), reply_markup=build_notify_markup(u_cfg))
                            else:
                                await _ack_cb()
                            continue

                        elif cb_data == "action:scan":
                            act_q = _active_queries(u_cfg)
                            if not act_q:
                                await _ack_cb("⚠️ Сначала отметьте ✅ хотя бы 1 товар!")
                                await tg_send_message(client, token, cid, "⚠️ <b>Не отмечен ни один товар ✅.</b>\nВыберите товары кнопками ниже или напишите название в чат:", reply_markup=build_products_markup(u_cfg))
                                continue
                            await _ack_cb("⚡ Поиск запущен!")
                            u_cfg["ladder_step"] = 0
                            rt.user_store.save()
                            await tg_send_message(
                                client,
                                token,
                                cid,
                                f"🚀 <b>Поиск минимальных цен запущен</b> <i>({len(act_q)} поз., по {len(SHOPS_META)} магазинов на позицию)</i>…",
                                reply_markup=TG_MAIN_KEYBOARD,
                            )
                            trigger_immediate_scan(cid)
                            continue

                        elif cid and cb_data.startswith("shop:"):
                            shop_name = cb_data.split(":", 1)[1]
                            await _ack_cb(f"✅ {shop_name}")
                            av_view = rt.avito_view_filter.get(cid, "ALL") if shop_name == "Авито" else "ALL"
                            top_txt = format_top_items_tg(shop_name, cid, u_cfg)
                            top_markup = build_shops_top_markup(u_cfg, active_shop=shop_name, avito_view=av_view)
                            if mid:
                                await tg_edit_message(client, token, cid, mid, top_txt, reply_markup=top_markup)
                            else:
                                await tg_send_message(client, token, cid, top_txt, reply_markup=top_markup)
                            continue

                        elif cid and cb_data.startswith("avitoview:"):
                            av_view = cb_data.split(":", 1)[1].strip().upper()
                            if av_view not in ("ALL", "PICKUP", "DELIVERY"):
                                av_view = "ALL"
                            rt.avito_view_filter[cid] = av_view
                            lbl_toast = {"PICKUP": "📍 Только самовывоз 5–15 км", "DELIVERY": "🚚 Только Авито Доставка", "ALL": "🌐 Самовывоз + Доставка"}.get(av_view, "Авито")
                            await _ack_cb(lbl_toast)
                            top_txt = format_top_items_tg("Авито", cid, u_cfg)
                            top_markup = build_shops_top_markup(u_cfg, active_shop="Авито", avito_view=av_view)
                            if mid:
                                await tg_edit_message(client, token, cid, mid, top_txt, reply_markup=top_markup)
                            else:
                                await tg_send_message(client, token, cid, top_txt, reply_markup=top_markup)
                            continue

                        elif cid and cb_data.startswith("avitorad:"):
                            try:
                                new_rad = max(5, min(int(cb_data.split(":", 1)[1]), 15))
                            except Exception:
                                new_rad = 15
                            await _ack_cb(f"📍 Радиус Авито: {new_rad} км")
                            u_cfg["avito_radius_km"] = new_rad
                            cfg["avito_radius_km"] = new_rad
                            _invalidate_user_cache(cid, clear_global=True)
                            rt.user_store.save()
                            save_config(cfg)
                            if mid:
                                try:
                                    await client.post(
                                        f"https://api.telegram.org/bot{token}/editMessageReplyMarkup",
                                        json={
                                            "chat_id": cid,
                                            "message_id": mid,
                                            "reply_markup": build_cities_markup(
                                                find_city_page(u_cfg.get("region", "")),
                                                u_cfg.get("region", "Все регионы РФ"),
                                                str(u_cfg.get("avito_mode", "both")),
                                                new_rad,
                                            ),
                                        },
                                        timeout=6.0,
                                    )
                                except Exception:
                                    pass
                            if _active_queries(u_cfg):
                                await tg_send_message(
                                    client,
                                    token,
                                    cid,
                                    f"📍 <b>Радиус поиска на Авито установлен: ≤ {new_rad} км</b> (регион: <b>{_h(u_cfg.get('region', 'Все регионы РФ'))}</b>).\n⚡ <b>Обновляю предложения...</b>",
                                    reply_markup=TG_MAIN_KEYBOARD,
                                )
                                trigger_immediate_scan(cid)
                            continue

                        elif cid and cb_data.startswith("avitomode:"):
                            new_av_mode = cb_data.split(":", 1)[1].strip().lower()
                            if new_av_mode not in ("both", "pickup", "delivery"):
                                new_av_mode = "both"
                            m_lbl = {"pickup": "📍 Только самовывоз 5–15 км", "delivery": "🚚 Только Доставка", "both": "🌐 Самовывоз + Доставка"}[new_av_mode]
                            await _ack_cb(m_lbl)
                            u_cfg["avito_mode"] = new_av_mode
                            cfg["avito_mode"] = new_av_mode
                            _invalidate_user_cache(cid, clear_global=True)
                            rt.user_store.save()
                            save_config(cfg)
                            if mid:
                                try:
                                    await client.post(
                                        f"https://api.telegram.org/bot{token}/editMessageReplyMarkup",
                                        json={
                                            "chat_id": cid,
                                            "message_id": mid,
                                            "reply_markup": build_cities_markup(
                                                find_city_page(u_cfg.get("region", "")),
                                                u_cfg.get("region", "Все регионы РФ"),
                                                new_av_mode,
                                                int(u_cfg.get("avito_radius_km", 15) or 15),
                                            ),
                                        },
                                        timeout=6.0,
                                    )
                                except Exception:
                                    pass
                            if _active_queries(u_cfg):
                                await tg_send_message(
                                    client,
                                    token,
                                    cid,
                                    f"🧭 <b>Режим Авито обновлён:</b> {m_lbl} (регион: <b>{_h(u_cfg.get('region', 'Все регионы РФ'))}</b>).\n⚡ <b>Запускаю поиск...</b>",
                                    reply_markup=TG_MAIN_KEYBOARD,
                                )
                                trigger_immediate_scan(cid)
                            continue

                        elif cid and cb_data.startswith("city:"):
                            raw_city = cb_data.split(":", 1)[1]
                            city_info = resolve_city_settings(raw_city)
                            await _ack_cb(f"✅ Город: {city_info['city_name']}")
                            u_cfg["region"] = city_info["city_name"]
                            u_cfg["region_slug_avito"] = city_info["region_slug_avito"]
                            u_cfg["ymarket_lr"] = city_info["ymarket_lr"]
                            u_cfg["lat"] = None
                            u_cfg["lon"] = None
                            u_cfg["detected_place"] = ""
                            u_cfg["ladder_step"] = 0
                            cfg["region"] = city_info["city_name"]
                            cfg["region_slug_avito"] = city_info["region_slug_avito"]
                            cfg["ymarket_lr"] = city_info["ymarket_lr"]
                            _invalidate_user_cache(cid, clear_global=True)
                            rt.user_store.save()
                            save_config(cfg)
                            rt.tg_dialog_mode.pop(cid, None)
                            if mid:
                                try:
                                    await client.post(
                                        f"https://api.telegram.org/bot{token}/editMessageReplyMarkup",
                                        json={
                                            "chat_id": cid,
                                            "message_id": mid,
                                            "reply_markup": build_cities_markup(
                                                find_city_page(u_cfg["region"]),
                                                u_cfg["region"],
                                                str(u_cfg.get("avito_mode", "both")),
                                                int(u_cfg.get("avito_radius_km", 15) or 15),
                                            ),
                                        },
                                        timeout=6.0,
                                    )
                                except Exception:
                                    pass
                            act_q = _active_queries(u_cfg)
                            av_r = max(5, min(int(u_cfg.get("avito_radius_km", 15) or 15), 15))
                            await tg_send_message(
                                client,
                                token,
                                cid,
                                f"✅ <b>Город успешно изменён!</b>\n"
                                f"📍 <b>Выбрано:</b> ✅ {_h(u_cfg['region'])} <i>(Авито: самовывоз ≤{av_r} км + Доставка)</i>\n"
                                f"🎯 <b>Ищем ({len(act_q)}):</b> {_h(', '.join(act_q) or '—')}\n\n"
                                f"⚡ <b>Запускаю поиск для города {_h(u_cfg['region'])}...</b>",
                                reply_markup=TG_MAIN_KEYBOARD,
                            )
                            if act_q:
                                trigger_immediate_scan(cid)
                            continue

                        await _ack_cb()
                        continue

                    # 2. Обработка сообщений и геопозиции
                    msg = upd.get("message") or {}
                    cid = str(msg.get("chat", {}).get("id") or "")
                    if not cid:
                        continue

                    rt.active_chat_ids.add(cid)
                    save_subscribers(rt.active_chat_ids)
                    u_cfg = _get_user_effective_cfg(cfg, cid)

                    if not cfg.get("telegram", {}).get("chat_id"):
                        cfg.setdefault("telegram", {})["chat_id"] = cid
                        cfg["telegram"]["enabled"] = True
                        save_config(cfg)

                    raw_text_low = (msg.get("text") or "").strip().lower()
                    is_changelog_cmd = raw_text_low in ("/changelog", "/news", "/version", "🆕 обновления бота", "обновления", "версия")
                    if msg.get("location") or raw_text_low:
                        await _ensure_user_update_news(client, token, cid, u_cfg, skip_if_changelog=is_changelog_cmd)

                    # Приём геопозиции: мгновенное определение города/региона + самовывоз 5–15 км на Авито
                    loc = msg.get("location")
                    if isinstance(loc, dict) and loc.get("latitude") is not None and loc.get("longitude") is not None:
                        lat_v = float(loc["latitude"])
                        lon_v = float(loc["longitude"])
                        geo_info = await resolve_city_by_coords_async(lat_v, lon_v)
                        u_cfg["lat"] = lat_v
                        u_cfg["lon"] = lon_v
                        u_cfg["region"] = geo_info["city_name"]
                        u_cfg["region_slug_avito"] = geo_info["region_slug_avito"]
                        u_cfg["ymarket_lr"] = geo_info["ymarket_lr"]
                        u_cfg["detected_place"] = geo_info["detected_place"]
                        u_cfg["ladder_step"] = 0
                        cfg["region"] = geo_info["city_name"]
                        cfg["region_slug_avito"] = geo_info["region_slug_avito"]
                        cfg["ymarket_lr"] = geo_info["ymarket_lr"]
                        _invalidate_user_cache(cid, clear_global=True)
                        rt.user_store.save()
                        save_config(cfg)
                        rt.tg_dialog_mode.pop(cid, None)
                        act_q = _active_queries(u_cfg)
                        scan_note = (
                            f"⚡ <b>Запускаю обновление цен ({len(act_q)} поз.: {_h(', '.join(act_q))})...</b>"
                            if act_q
                            else "✍️ <b>Напишите название товара, чтобы сразу начать поиск в вашем регионе!</b>"
                        )
                        await tg_send_message(
                            client,
                            token,
                            cid,
                            f"📌 <b>Геолокация определена моментально!</b>\n"
                            f"📍 <b>Ваш регион:</b> ✅ <b>{_h(geo_info['detected_place'])}</b>\n"
                            f"🏬 <b>Настроены:</b> Авито (самовывоз <b>5–15 км</b> рядом с вами + Доставка по РФ), Яндекс Маркет, Ozon, WB, AliExpress, DNS.\n\n"
                            f"{scan_note}",
                            reply_markup=TG_MAIN_KEYBOARD,
                        )
                        if act_q:
                            trigger_immediate_scan(cid)
                        continue

                    text = (msg.get("text") or "").strip()
                    if not text:
                        continue

                    low = text.lower()
                    mode = rt.tg_dialog_mode.get(cid, "")

                    if low in ("/start", "/help", "привет", "здравствуйте"):
                        rt.tg_dialog_mode.pop(cid, None)
                        await tg_send_welcome_blocks(client, token, cid, u_cfg)
                    elif low in ("/status", "⚙️ текущие настройки"):
                        rt.tg_dialog_mode.pop(cid, None)
                        await tg_send_message(client, token, cid, format_settings_tg(u_cfg, cid), reply_markup=TG_MAIN_KEYBOARD)
                    elif low in ("/changelog", "/news", "/version", "🆕 обновления бота", "обновления", "версия"):
                        rt.tg_dialog_mode.pop(cid, None)
                        await tg_send_message(client, token, cid, format_update_news_tg(), reply_markup=TG_MAIN_KEYBOARD)
                    elif low in ("/items", "/products", "📦 мои товары (до 10)", "📦 мои товары", "мои товары", "список товаров"):
                        rt.tg_dialog_mode.pop(cid, None)
                        await tg_send_message(client, token, cid, format_products_menu_text(u_cfg), reply_markup=build_products_markup(u_cfg))
                    elif low in ("/compare", "⚖️ сравнение цена/качество", "цена/качество", "сравнение"):
                        rt.tg_dialog_mode.pop(cid, None)
                        pq_txt = format_price_quality_comparison_tg(None, cid, max_rows=10, query_filter="ALL")
                        await tg_send_html_chunks(client, token, cid, pq_txt, reply_markup=build_pq_markup(u_cfg, "ALL"))
                    elif low in ("/shops", "🏬 выбрать магазины", "магазины"):
                        rt.tg_dialog_mode.pop(cid, None)
                        await tg_send_message(
                            client,
                            token,
                            cid,
                            "🏬 <b>Выбор магазинов (отметьте галочками ✅ нужные площадки):</b>",
                            reply_markup=build_shops_filter_markup(u_cfg),
                        )
                    elif low in ("/notify", "/settings", "🔔 уведомления и интервал", "⚙️ настройки и интервал", "уведомления"):
                        rt.tg_dialog_mode.pop(cid, None)
                        await tg_send_message(client, token, cid, format_notify_settings_text(u_cfg), reply_markup=build_notify_markup(u_cfg))
                    elif low in ("/excel", "📥 скачать excel (.xlsx)"):
                        rt.tg_dialog_mode.pop(cid, None)
                        u_items, empty_reason = _get_scoped_items(cid, u_cfg)
                        if not u_items:
                            await tg_send_message(client, token, cid, empty_reason, reply_markup=TG_MAIN_KEYBOARD)
                        else:
                            u_reg_lbl = str(u_cfg.get("region") or "Все регионы РФ")
                            u_excel = rt.user_excel_paths.get(cid)
                            if not u_excel or not u_excel.exists():
                                u_excel = _CACHE_DIR / f"Мониторинг_Товаров_{cid}.xlsx"
                                save_to_excel(u_items, u_excel, city_name=u_reg_lbl)
                                rt.user_excel_paths[cid] = u_excel
                            await send_telegram_excel(
                                client,
                                token,
                                cid,
                                f"📥 Актуальная таблица Мониторинг_Товаров.xlsx — 📍 {u_reg_lbl} ({len(u_items)} позиций)",
                                excel_file=u_excel,
                            )
                    elif low in ("/top", "📋 топ цен по магазинам", "📋 лучшие цены по магазинам", "лучшие цены по магазинам"):
                        rt.tg_dialog_mode.pop(cid, None)
                        await tg_send_message(
                            client,
                            token,
                            cid,
                            format_top_items_tg("ALL", cid, u_cfg),
                            reply_markup=build_shops_top_markup(u_cfg, "ALL"),
                        )
                    elif low in (
                        "/scan",
                        "🔄 обновить сейчас",
                        "🔍 поиск новых цен",
                        "🔍 поиск новых цен по выбранному товару",
                        "поиск новых цен по выбранному товару",
                        "поиск новых цен",
                        "поиск",
                    ):
                        rt.tg_dialog_mode.pop(cid, None)
                        act_q = _active_queries(u_cfg)
                        if not act_q:
                            await tg_send_message(
                                client,
                                token,
                                cid,
                                "⚠️ <b>Список товаров пуст или ни один товар не отмечен ✅.</b>\nВыберите позиции или добавьте новый товар:",
                                reply_markup=build_products_markup(u_cfg),
                            )
                            continue
                        u_cfg["ladder_step"] = 0
                        rt.user_store.save()
                        await tg_send_message(
                            client,
                            token,
                            cid,
                            f"🚀 <b>Поиск минимальных цен запущен</b> <i>({len(act_q)} поз., по {len(SHOPS_META)} магазинов на позицию)</i>…",
                            reply_markup=TG_MAIN_KEYBOARD,
                        )
                        trigger_immediate_scan(cid)
                    elif low in ("/product", "/replace", "🎯 сменить товар", "🔄 заменить один товар", "заменить один товар", "заменить товар"):
                        _sync_query_lists(u_cfg)
                        cur_queries = list(u_cfg.get("queries", []))
                        if low in ("/replace", "🔄 заменить один товар", "заменить один товар", "заменить товар") and len(cur_queries) >= 2:
                            rt.tg_dialog_mode.pop(cid, None)
                            await tg_send_message(
                                client,
                                token,
                                cid,
                                f"🔄 <b>Какой именно товар заменить?</b> <i>(сейчас в списке {len(cur_queries)} шт.)</i>\n"
                                f"Нажмите на любую позицию ниже — заменится только она, а остальные товары сохранятся:",
                                reply_markup=build_replace_pick_markup(u_cfg),
                            )
                        else:
                            rt.tg_dialog_mode[cid] = "WAIT_PRODUCT"
                            await tg_send_message(
                                client,
                                token,
                                cid,
                                f"✍️ <b>Напишите новое название товара</b> (или несколько через запятую / с новой строки — список обновится):\n\n"
                                f"Примеры:\n"
                                f"• <code>бытовой утюг паровой</code>\n"
                                f"• <code>бумага а4, ручка шариковая</code>\n"
                                f"• <code>кресло компьютерное</code>\n"
                                f"<i>Если нужно добавить товар или заменить 1 позицию из списка — откройте «📦 Мои товары».</i>",
                                reply_markup=build_products_markup(u_cfg),
                            )
                    elif low in ("/price", "💰 сменить цену (от и до)"):
                        rt.tg_dialog_mode[cid] = "WAIT_PRICE"
                        await tg_send_message(
                            client,
                            token,
                            cid,
                            f"💰 Сейчас установлена цена: <b>от {u_cfg.get('min_price'):,} до {u_cfg.get('max_price'):,} ₽</b>.\n\n"
                            f"✍️ <b>Напишите минимальную и максимальную цену в рублях</b> через дефис или пробел:\n"
                            f"Например: <code>300 - 5000</code> или <code>15000 120000</code>",
                            reply_markup={"inline_keyboard": [[{"text": "🤖 Включить авто-подбор цены по каждому товару", "callback_data": "action:auto_price"}]]},
                        )
                    elif low in ("/city", "📍 сменить город"):
                        rt.tg_dialog_mode[cid] = "WAIT_CITY"
                        cur_city = u_cfg.get("detected_place") or u_cfg.get("region", "Все регионы РФ")
                        await tg_send_message(
                            client,
                            token,
                            cid,
                            f"📍 Сейчас выбран регион: ✅ <b>{_h(cur_city)}</b>\n"
                            f"<i>(Нажмите кнопку «📌 Отправить мою геолокацию» внизу, выберите город / радиус Авито 5–15 км / режим ниже или напишите город)</i>",
                            reply_markup=build_cities_markup(
                                find_city_page(u_cfg.get("region", "")),
                                u_cfg.get("region", "Все регионы РФ"),
                                str(u_cfg.get("avito_mode", "both")),
                                int(u_cfg.get("avito_radius_km", 15) or 15),
                            ),
                        )
                    elif mode == "WAIT_CITY":
                        city_info = resolve_city_settings(text)
                        u_cfg["region"] = city_info["city_name"]
                        u_cfg["region_slug_avito"] = city_info["region_slug_avito"]
                        u_cfg["ymarket_lr"] = city_info["ymarket_lr"]
                        u_cfg["lat"] = None
                        u_cfg["lon"] = None
                        u_cfg["detected_place"] = ""
                        u_cfg["ladder_step"] = 0
                        cfg["region"] = city_info["city_name"]
                        cfg["region_slug_avito"] = city_info["region_slug_avito"]
                        cfg["ymarket_lr"] = city_info["ymarket_lr"]
                        _invalidate_user_cache(cid, clear_global=True)
                        rt.user_store.save()
                        save_config(cfg)
                        rt.tg_dialog_mode.pop(cid, None)
                        act_q = _active_queries(u_cfg)
                        av_r = max(5, min(int(u_cfg.get("avito_radius_km", 15) or 15), 15))
                        await tg_send_message(
                            client,
                            token,
                            cid,
                            f"✅ <b>Город успешно изменён!</b>\n"
                            f"📍 <b>Выбрано:</b> ✅ {_h(u_cfg['region'])} <i>(Самовывоз ≤{av_r} км + Доставка по РФ)</i>\n"
                            f"🎯 <b>Ищем:</b> {_h(', '.join(act_q) or '—')}\n\n"
                            f"⚡ <b>Запускаю поиск...</b>",
                            reply_markup=TG_MAIN_KEYBOARD,
                        )
                        if act_q:
                            trigger_immediate_scan(cid)
                    elif mode in ("WAIT_PRICE", "WAIT_PRICE_AFTER_PRODUCT"):
                        nums = re.findall(r"\d+", text.replace(" ", ""))
                        if len(nums) >= 2:
                            p_min, p_max = sorted([int(nums[0]), int(nums[1])])
                            u_cfg["min_price"] = p_min
                            u_cfg["max_price"] = p_max
                            u_cfg["price_is_custom"] = True
                            cfg["min_price"] = p_min
                            cfg["max_price"] = p_max
                            cfg["price_is_custom"] = True
                        elif len(nums) == 1:
                            u_cfg["min_price"] = int(nums[0])
                            u_cfg["price_is_custom"] = True
                            cfg["min_price"] = int(nums[0])
                            cfg["price_is_custom"] = True
                        u_cfg["ladder_step"] = 0
                        _invalidate_user_cache(cid, clear_global=True)
                        rt.user_store.save()
                        save_config(cfg)
                        rt.tg_dialog_mode.pop(cid, None)
                        act_q = _active_queries(u_cfg)
                        await tg_send_message(
                            client,
                            token,
                            cid,
                            f"✅ <b>Диапазон цен сохранён!</b>\n"
                            f"🎯 <b>Ищем:</b> {_h(', '.join(act_q))}\n"
                            f"💰 <b>Цена:</b> от {u_cfg['min_price']:,} до {u_cfg['max_price']:,} ₽\n\n"
                            f"⚡ <b>Запускаю поиск минимальных цен...</b>",
                            reply_markup=TG_MAIN_KEYBOARD,
                        )
                        if act_q:
                            trigger_immediate_scan(cid)
                    else:
                        # Ввод товаров (WAIT_REPLACE_PRODUCT:<idx> / WAIT_ADD_PRODUCT / WAIT_PRODUCT / прямой ввод в чат)
                        was_add_mode = mode == "WAIT_ADD_PRODUCT"
                        rep_idx: Optional[int] = None
                        if mode.startswith("WAIT_REPLACE_PRODUCT:"):
                            try:
                                rep_idx = int(mode.split(":", 1)[1])
                            except Exception:
                                rep_idx = None
                        rt.tg_dialog_mode.pop(cid, None)

                        action_header = "✅ <b>Список товаров обновлён и выбран для поиска ✅ (уведомления 🔔 включены):</b>"
                        if "|" in text:
                            parts = [p.strip() for p in text.split("|")]
                            q_part = parts[0]
                            p_part = parts[1] if len(parts) > 1 else ""
                            c_part = parts[2] if len(parts) > 2 else ""
                            new_q = [x.strip() for x in re.split(r"[,;\n]+", q_part) if x.strip()][:MAX_PRODUCTS]
                            if new_q:
                                u_cfg["queries"] = list(dict.fromkeys(new_q))[:MAX_PRODUCTS]
                                u_cfg["selected_queries"] = list(u_cfg["queries"])
                                muted_s = set(u_cfg.get("muted_queries", [])) - set(u_cfg["queries"])
                                u_cfg["muted_queries"] = sorted(list(muted_s))
                                cfg["queries"] = u_cfg["queries"]
                                cfg["selected_queries"] = u_cfg["selected_queries"]
                            nums = re.findall(r"\d+", p_part.replace(" ", ""))
                            if len(nums) >= 2:
                                p_min, p_max = sorted([int(nums[0]), int(nums[1])])
                                u_cfg["min_price"], u_cfg["max_price"] = p_min, p_max
                                u_cfg["price_is_custom"] = True
                                cfg["min_price"], cfg["max_price"] = p_min, p_max
                                cfg["price_is_custom"] = True
                            if c_part:
                                city_info = resolve_city_settings(c_part)
                                u_cfg["region"] = city_info["city_name"]
                                u_cfg["region_slug_avito"] = city_info["region_slug_avito"]
                                u_cfg["ymarket_lr"] = city_info["ymarket_lr"]
                                u_cfg["lat"] = None
                                u_cfg["lon"] = None
                                u_cfg["detected_place"] = ""
                                cfg["region"] = city_info["city_name"]
                                cfg["region_slug_avito"] = city_info["region_slug_avito"]
                                cfg["ymarket_lr"] = city_info["ymarket_lr"]
                        else:
                            new_q = [x.strip() for x in re.split(r"[,;\n]+", text) if x.strip()][:MAX_PRODUCTS]
                            if new_q:
                                new_q = list(dict.fromkeys(new_q))
                                existing = [str(x).strip() for x in (u_cfg.get("queries") or []) if str(x).strip()]
                                existing_sel = list(u_cfg.get("selected_queries") or [])
                                if rep_idx is not None and 0 <= rep_idx < len(existing):
                                    old_q = existing[rep_idx]
                                    replaced_list = existing[:rep_idx] + [x for x in new_q if x not in existing or x == old_q] + existing[rep_idx + 1 :]
                                    replaced_list = list(dict.fromkeys(replaced_list))[:MAX_PRODUCTS]
                                    u_cfg["queries"] = replaced_list
                                    sel_new = [x for x in existing_sel if x != old_q] + list(new_q)
                                    u_cfg["selected_queries"] = [q for q in u_cfg["queries"] if q in set(sel_new)]
                                    muted_s = (set(u_cfg.get("muted_queries", [])) - {old_q} - set(new_q)) & set(u_cfg["queries"])
                                    u_cfg["muted_queries"] = sorted(list(muted_s))
                                    action_header = f"✏️ <b>Позиция #{rep_idx + 1} «{_h(old_q)}» заменена на «{_h(', '.join(new_q))}»!</b>"
                                elif was_add_mode:
                                    for item_q in new_q:
                                        if item_q not in existing:
                                            if len(existing) >= MAX_PRODUCTS:
                                                unselected_idx = next((i for i, old_q in enumerate(existing) if old_q not in set(existing_sel)), -1)
                                                if unselected_idx >= 0:
                                                    existing.pop(unselected_idx)
                                                else:
                                                    existing.pop()
                                            existing.append(item_q)
                                    u_cfg["queries"] = existing[:MAX_PRODUCTS]
                                    sel_union = list(dict.fromkeys(existing_sel + new_q))
                                    u_cfg["selected_queries"] = [q for q in u_cfg["queries"] if q in set(sel_union)]
                                    muted_s = set(u_cfg.get("muted_queries", [])) - set(new_q)
                                    u_cfg["muted_queries"] = sorted(list(muted_s))
                                    action_header = "✅ <b>Товар добавлен и выбран для поиска ✅ (уведомления 🔔 включены):</b>"
                                else:
                                    u_cfg["queries"] = list(new_q)[:MAX_PRODUCTS]
                                    u_cfg["selected_queries"] = list(u_cfg["queries"])
                                    muted_s = set(u_cfg.get("muted_queries", [])) & set(u_cfg["queries"]) - set(new_q)
                                    u_cfg["muted_queries"] = sorted(list(muted_s))

                                u_cfg["price_is_custom"] = False
                                p_min, p_max = _auto_price_bounds(u_cfg["selected_queries"] or new_q)
                                u_cfg["min_price"], u_cfg["max_price"] = p_min, p_max
                                cfg["queries"] = u_cfg["queries"]
                                cfg["selected_queries"] = u_cfg["selected_queries"]
                                cfg["muted_queries"] = u_cfg["muted_queries"]
                                cfg["min_price"], cfg["max_price"] = p_min, p_max
                                cfg["price_is_custom"] = False
                        u_cfg["ladder_step"] = 0
                        _invalidate_user_cache(cid, clear_global=True)
                        rt.user_store.save()
                        save_config(cfg)
                        act_q = _active_queries(u_cfg)
                        await tg_send_message(
                            client,
                            token,
                            cid,
                            f"{action_header}\n"
                            f"🎯 <b>Ищем ({len(act_q)}):</b> {_h(', '.join(act_q))}\n"
                            f"📍 <b>Регион:</b> ✅ {_h(u_cfg.get('detected_place') or u_cfg.get('region', 'Все регионы РФ'))}\n"
                            f"💰 <b>Диапазон цен:</b> от {u_cfg['min_price']:,} до {u_cfg['max_price']:,} ₽\n\n"
                            f"⚡ <b>Запускаю поиск минимальных цен и сверку карточек!</b>",
                            reply_markup=TG_MAIN_KEYBOARD,
                        )
                        if act_q:
                            trigger_immediate_scan(cid)
                if rt.last_update_id > prev_upd_id:
                    rt.user_store.meta["last_update_id"] = rt.last_update_id
                    rt.user_store.save()
            except Exception:
                await asyncio.sleep(1.0)


def prompt_queries_on_start(cfg: dict, timeout_sec: int = 8) -> dict:
    if not sys.stdin or not sys.stdin.isatty():
        return cfg

    queries = cfg.get("queries", [])
    cur_min = int(cfg.get("min_price", 20000))
    cur_max = int(cfg.get("max_price", 160000))
    cur_city = cfg.get("region", "Все регионы РФ")

    console.print("\n[bold cyan]======================================================================[/bold cyan]")
    console.print("[bold white]  🎯 ТЕКУЩИЕ НАСТРОЙКИ ПОИСКА В 6 МАРКЕТПЛЕЙСАХ:[/bold white]")
    console.print(f"     Регион поиска: [bold cyan]📍 {cur_city}[/bold cyan]")
    for idx, q in enumerate(queries, start=1):
        console.print(f"     [bold yellow]{idx}. {q}[/bold yellow]")
    console.print(f"     Диапазон цен: [bold green]от {cur_min:,} до {cur_max:,} ₽[/bold green]")
    console.print("[bold cyan]----------------------------------------------------------------------[/bold cyan]")
    console.print("  • Нажмите [bold green]ENTER[/bold green] (или подождите 8 сек) — чтобы искать с этими настройками")
    console.print("  • Или [bold yellow]введите новый товар[/bold yellow] (либо нажмите ПРОБЕЛ, чтобы поменять цену или регион):")
    console.print("  • Также управление кнопками доступно прямо в Telegram-боте 24/7!")
    console.print("[bold cyan]======================================================================[/bold cyan]")

    sys.stdout.write(f"Новые товары (автостарт через {timeout_sec} сек): ")
    sys.stdout.flush()

    buf: List[str] = []
    start_t = time.time()
    typing_started = False

    while True:
        if msvcrt.kbhit():
            ch = msvcrt.getwch()
            if ch in ("\r", "\n"):
                sys.stdout.write("\n")
                sys.stdout.flush()
                break
            elif ch == "\x08":
                if buf:
                    buf.pop()
                    sys.stdout.write("\b \b")
                    sys.stdout.flush()
            elif ch == "\x03":
                raise KeyboardInterrupt
            else:
                if not typing_started:
                    typing_started = True
                    sys.stdout.write("\rНовые товары (Enter = оставить текущие): " + " " * 15 + "\r")
                    sys.stdout.write("Новые товары (Enter = оставить текущие): ")
                buf.append(ch)
                sys.stdout.write(ch)
                sys.stdout.flush()
        else:
            if not typing_started and (time.time() - start_t) >= timeout_sec:
                sys.stdout.write("\n[Автостарт с текущими настройками...]\n")
                sys.stdout.flush()
                break
            time.sleep(0.05)

    if not typing_started:
        return cfg

    changed = False
    user_input = "".join(buf).strip()
    if user_input:
        new_queries = [x.strip() for x in user_input.replace(";", ",").split(",") if x.strip()]
        if new_queries:
            cfg["queries"] = new_queries
            changed = True

    try:
        min_str = input(f"Минимальная цена ОТ в рублях (Enter = оставить {cur_min}): ").strip()
        if min_str:
            digits = "".join(c for c in min_str if c.isdigit())
            if digits:
                cfg["min_price"] = int(digits)
                changed = True

        max_str = input(f"Максимальная цена ДО в рублях (Enter = оставить {cur_max}): ").strip()
        if max_str:
            digits = "".join(c for c in max_str if c.isdigit())
            if digits:
                cfg["max_price"] = int(digits)
                changed = True

        city_str = input(f"Город поиска (Enter = оставить {cur_city}): ").strip()
        if city_str:
            city_info = resolve_city_settings(city_str)
            cfg["region"] = city_info["city_name"]
            cfg["region_slug_avito"] = city_info["region_slug_avito"]
            cfg["ymarket_lr"] = city_info["ymarket_lr"]
            changed = True
    except EOFError:
        pass

    if changed:
        save_config(cfg)
        console.print(
            f"[bold green]✅ Настройки сохранены:[/bold green] "
            f"📍 {cfg.get('region', 'Все регионы РФ')} | "
            f"{', '.join(cfg.get('queries', []))} | "
            f"Цена: [bold yellow]{cfg.get('min_price'):,} – {cfg.get('max_price'):,} ₽[/bold yellow]"
        )
    return cfg


async def run_cycle(cfg: dict, engine: IsolatedChromeEngine, target_cid: str = "") -> List[ProductItem]:
    rt.is_scanning = True
    queries: List[str] = _active_queries(cfg) or ["кресло компьютерное"]
    min_price = int(cfg.get("min_price", 50))
    max_price = int(cfg.get("max_price", 250000))
    price_is_custom = bool(cfg.get("price_is_custom", False))
    limit = int(cfg.get("limit_per_shop", 100))
    shops = cfg.get("shops_priority", {})
    city_name = cfg.get("region", "Все регионы РФ")
    region_slug = cfg.get("region_slug_avito", "all")
    ymarket_lr = int(cfg.get("ymarket_lr", 213))
    avito_delivery = bool(cfg.get("avito_delivery_only", True))
    avito_all_ru = bool(cfg.get("avito_all_russia_with_delivery", True))
    avito_radius = max(5, min(int(cfg.get("avito_radius_km", 15) or 15), 15))
    avito_mode = str(cfg.get("avito_mode", "both") or "both").strip().lower()
    if avito_mode not in ("both", "pickup", "delivery"):
        avito_mode = "both"
    u_lat = cfg.get("lat")
    u_lon = cfg.get("lon")

    history = load_history()
    collected: List[ProductItem] = []

    base_cfg = load_config()
    tg = base_cfg.get("telegram", {})
    token = str(tg.get("bot_token") or "").strip()
    cid = str(target_cid or tg.get("chat_id") or "").strip()
    progress_msg_id: Optional[int] = None
    pos_states: List[dict] = [
        {
            "query": q,
            "states": {sk: ("off" if not shops.get(sk, True) else "queued") for sk, _, _ in SHOPS_META},
            "counts": {sk: 0 for sk, _, _ in SHOPS_META},
        }
        for q in queries
    ]

    async def _update_live_notification(step_label: str, done: bool = False) -> None:
        nonlocal progress_msg_id
        if not (token and cid):
            return
        text_body = format_scan_progress_tg(
            step_label=step_label,
            city_name=city_name,
            min_price=min_price,
            max_price=max_price,
            price_is_custom=price_is_custom,
            pos_states=pos_states,
            done=done,
        )
        try:
            async with httpx.AsyncClient(timeout=10.0) as client:
                if progress_msg_id is None:
                    progress_msg_id = await tg_send_message(client, token, cid, text_body)
                else:
                    await tg_edit_message(client, token, cid, progress_msg_id, text_body)
        except Exception:
            pass

    async def _run_shop_step(
        q_idx: int,
        q: str,
        sk_key: str,
        sh_name: str,
        color: str,
        step_suffix: str,
        parser_fn,
        *p_args,
        **p_kwargs,
    ) -> None:
        if not shops.get(sk_key, True):
            return
        pos_states[q_idx]["states"][sk_key] = "running"
        rt.current_step = f"Поз. {q_idx + 1} «{q}» → {step_suffix}"
        await _update_live_notification(rt.current_step)
        console.print(f"  [{step_suffix.split(' ', 1)[0]}] Сбор [{color}]{sh_name}[/{color}] ({city_name})...")
        try:
            items = await asyncio.wait_for(parser_fn(*p_args, **p_kwargs), timeout=55.0)
            console.print(f"        Найдено на {sh_name}: [bold green]{len(items)}[/bold green] поз.")
            collected.extend(items)
            pos_states[q_idx]["counts"][sk_key] = len(items)
        except asyncio.TimeoutError:
            console.print(f"        [yellow]⏱ Таймаут {sh_name} (>55с) — идём дальше[/yellow]")
            pos_states[q_idx]["counts"][sk_key] = 0
        except Exception as e:
            console.print(f"        [red]Ошибка {sh_name}:[/red] {e}")
            pos_states[q_idx]["counts"][sk_key] = 0
        pos_states[q_idx]["states"][sk_key] = "done"

    try:
        try:
            await engine.apply_region_context(
                city_name=city_name,
                ymarket_lr=ymarket_lr,
                lat=u_lat,
                lon=u_lon,
            )
        except Exception:
            pass
        wb_geo = None
        if shops.get("wb", True):
            try:
                g_lat, g_lon = get_city_coords(city_name, u_lat, u_lon)
                wb_geo = await fetch_wb_geo_payload(g_lat, g_lon)
            except Exception:
                wb_geo = None

        await _update_live_notification("Запуск сканирования (по 6 магазинов на позицию)...", done=False)

        for q_idx, q in enumerate(queries):
            if price_is_custom:
                q_min, q_max = min_price, max_price
            else:
                q_min, q_max = _auto_price_bounds([q])
            console.print(
                f"\n[bold cyan]🔎 Поисковый запрос {q_idx + 1}/{len(queries)}:[/bold cyan] [bold yellow]«{q}»[/bold yellow] "
                f"({q_min:,} – {q_max:,} ₽ | 📍 {city_name})"
            )

            await _run_shop_step(q_idx, q, "ozon", "Ozon", "bold blue", f"1/6 Ozon ({city_name})", parse_ozon, engine, q, q_min, q_max, limit, city_name=city_name)
            await _run_shop_step(q_idx, q, "aliexpress", "AliExpress", "bold red", f"2/6 AliExpress ({city_name})", parse_aliexpress, engine, q, q_min, q_max, limit, city_name=city_name)
            await _run_shop_step(q_idx, q, "ymarket", "Яндекс Маркет", "bold yellow", f"3/6 Яндекс Маркет ({city_name})", parse_ymarket, engine, q, q_min, q_max, limit, city_name=city_name, lr=ymarket_lr)
            await _run_shop_step(
                q_idx,
                q,
                "avito",
                "Авито",
                "bold green",
                f"4/6 Авито (самовывоз ≤{avito_radius} км + Доставка)",
                parse_avito,
                engine,
                q,
                q_min,
                q_max,
                region_slug,
                avito_delivery,
                avito_all_ru,
                limit,
                city_name=city_name,
                radius_km=avito_radius,
                lat=u_lat,
                lon=u_lon,
                avito_mode=avito_mode,
            )
            await _run_shop_step(q_idx, q, "dns", "DNS", "bold dark_orange", f"5/6 DNS ({city_name})", parse_dns, engine, q, q_min, q_max, limit, city_name=city_name)
            await _run_shop_step(
                q_idx,
                q,
                "wb",
                "Wildberries",
                "bold magenta",
                f"6/6 Wildberries ({city_name})",
                parse_wb,
                engine,
                q,
                q_min,
                q_max,
                limit,
                city_name=city_name,
                lat=u_lat,
                lon=u_lon,
                wb_geo=wb_geo,
            )
            await _update_live_notification(f"Позиция {q_idx + 1} «{q}» проверена по 6 магазинам", done=False)

        v_stats = {"checked": 0, "total": len(collected), "updated": 0, "blocked": 0, "duties_found": 0}
        if collected:
            filter_category_outliers(collected)
            rt.current_step = f"Сверка карточек и внутренних пошлин для Excel ({len(collected)} шт.)..."
            await _update_live_notification(rt.current_step, done=False)
            console.print(f"  [🔍] Параллельная сверка карточек и пошлин для всех {len(collected)} товаров Excel (лимит 90с)...")

            async def _card_progress(st: dict) -> None:
                rt.current_step = (
                    f"Сверка карточек и пошлин: {st.get('checked', 0)}/{st.get('total', len(collected))} "
                    f"(пошлин с сайта: {st.get('duties_found', 0)})"
                )
                await _update_live_notification(rt.current_step, done=False)

            v_stats = await verify_all_cards(
                engine,
                collected,
                max_total_sec=90.0,
                progress_cb=_card_progress,
                ymarket_lr=ymarket_lr,
            )
            filter_category_outliers(collected)
            console.print(
                f"        Проверено карточек: [bold cyan]{v_stats['checked']}/{v_stats.get('total', len(collected))}[/bold cyan], "
                f"внутренних пошлин найдено: [bold yellow]{v_stats.get('duties_found', 0)}[/bold yellow], "
                f"уточнено цен/статусов: [bold green]{v_stats['updated']}[/bold green]"
            )

        await _update_live_notification("Формирование таблицы Excel и расчёт минимальных цен...", done=True)

        price_drop_alerts = []
        for item in collected:
            url_base = item.url.split("?")[0]
            hist_key = f"{city_name}::{url_base}"
            old_price = history.get(hist_key)
            if old_price is None:
                old_price = history.get(url_base)
            if old_price is None:
                item.price_change = "🆕 Новый лот"
            elif item.price < old_price:
                diff = old_price - item.price
                item.price_change = f"📉 -{diff:,} ₽ (было {old_price:,})"
                if item.status_code == "GREEN" and str(item.in_stock).startswith("✅"):
                    price_drop_alerts.append((item, diff, old_price))
            elif item.price > old_price:
                diff = item.price - old_price
                item.price_change = f"📈 +{diff:,} ₽"
            else:
                item.price_change = "Без изменений"
            history[hist_key] = item.price

        save_history(history)

        # Защита от устаревания: если пока шёл обход, пользователь сменил город или список товаров, не сохраняем старые данные
        if target_cid and rt.user_store.exists(str(target_cid)):
            live_u = rt.user_store.get(str(target_cid))
            live_q = _active_queries(live_u)
            if _norm_reg(live_u.get("region")) != _norm_reg(city_name) or (live_q and set(x.lower() for x in live_q) != set(x.lower() for x in queries)):
                console.print(f"[yellow]⚠️ Регион или список товаров изменился во время обхода ({city_name}) — устаревшие результаты пропущены.[/yellow]")
                return collected

        now_str = datetime.now().strftime("%d.%m.%Y %H:%M")
        rt.last_updated = now_str
        scan_meta = _build_scan_meta(
            city_name=city_name,
            queries=queries,
            cid=str(target_cid or cid or ""),
            avito_mode=avito_mode,
            avito_radius_km=avito_radius,
        )
        if collected:
            rt.items = collected
            rt.global_scan_meta = scan_meta
            if target_cid:
                rt.user_items[str(target_cid)] = collected
                rt.user_scan_meta[str(target_cid)] = scan_meta
        save_ts = time.time()
        rt.user_store.meta["last_global_scan_ts"] = save_ts
        rt.user_store.meta["last_global_meta"] = scan_meta
        u_key = str(target_cid or cid or "")
        if u_key and rt.user_store.exists(u_key):
            u_entry = rt.user_store.get(u_key)
            u_entry["last_scan"] = save_ts
            u_entry["last_scan_region"] = city_name
            u_entry["last_scan_queries"] = list(queries)
        rt.user_store.save()
        if collected:
            _save_items_snapshot_to_store(collected, scan_meta)

        out_excel_path = EXCEL_PATH
        if target_cid:
            user_excel = _CACHE_DIR / f"Мониторинг_Товаров_{target_cid}.xlsx"
            save_to_excel(collected, user_excel, city_name=city_name)
            rt.user_excel_paths[str(target_cid)] = user_excel
            out_excel_path = user_excel
        saved_file = save_to_excel(collected, EXCEL_PATH, city_name=city_name)

        green_cnt = sum(1 for i in collected if i.status_code == "GREEN" and str(i.in_stock).startswith("✅"))
        yellow_cnt = sum(1 for i in collected if i.status_code == "YELLOW" and str(i.in_stock).startswith("✅"))
        red_cnt = sum(1 for i in collected if i.status_code == "RED" or str(i.in_stock).startswith("❌"))
        shop_counts = {name: sum(1 for i in collected if i.shop == name) for _, name, _ in SHOPS_META}

        console.print(
            f"\n[bold green]✅ Таблица Excel обновлена:[/bold green] [underline]{saved_file}[/underline]\n"
            f"   Город: [bold cyan]{city_name}[/bold cyan] | "
            f"Всего в таблице: [bold]{len(collected)}[/bold] | "
            f"🟢 Зелёные (✅): [bold green]{green_cnt}[/bold green] | "
            f"🟡 Жёлтые (✅): [bold yellow]{yellow_cnt}[/bold yellow] | "
            f"🔴 Красные (❌): [bold red]{red_cnt}[/bold red]"
        )

        # Отправка результатов: если запуск персональный (target_cid) — только этому пользователю; иначе — только с тем же регионом
        if target_cid:
            recipients = {str(target_cid)}
        else:
            candidates = set(rt.active_chat_ids)
            if cid:
                candidates.add(cid)
            recipients = set()
            for r_c in candidates:
                if rt.user_store.exists(r_c):
                    r_reg = _norm_reg(rt.user_store.get(r_c).get("region"))
                    if r_reg != _norm_reg(city_name):
                        continue
                recipients.add(r_c)

        alert_filter = cfg.get("alert_filter", "GREEN_ONLY")
        muted_queries = set(cfg.get("muted_queries", []))

        if token and recipients:
            try:
                async with httpx.AsyncClient(timeout=45.0, trust_env=False) as client:
                    if alert_filter != "OFF" and price_drop_alerts:
                        try:
                            for it, diff, old_p in price_drop_alerts[:5]:
                                if it.query in muted_queries:
                                    continue
                                if alert_filter == "GREEN_ONLY" and it.status_code != "GREEN":
                                    continue
                                link_lbl = (
                                    "Перейти на AliExpress (🔥 скидка!) ↗"
                                    if it.shop == "AliExpress"
                                    else "Перейти к карточке товара ↗"
                                )
                                drop_duty = f" ({it.base_price:,} ₽ + пошлина {it.customs_duty:,} ₽)" if it.customs_duty > 0 else ""
                                drop_price_lbl = f"Итого {it.price:,} ₽" if it.customs_duty > 0 else f"{it.price:,} ₽"
                                drop_text = (
                                    f"📉 <b>ВНИМАНИЕ! ЦЕНА СНИЗИЛАСЬ на -{diff:,} ₽!</b>\n\n"
                                    f"🛍 <b>[{_h(it.shop)}]</b> {_h(it.title[:85])}\n"
                                    f"💰 <b>Новая цена: {drop_price_lbl}</b><i>{drop_duty} (было {old_p:,} ₽)</i>\n"
                                    f"📍 {_h(it.location_delivery)}\n"
                                    f"👤 {_h(it.seller_name)} (⭐ {it.seller_rating})\n"
                                    f"🔗 <a href=\"{_h(clean_product_url(it.url))}\">{link_lbl}</a>"
                                )
                                for r_cid in recipients:
                                    await tg_send_message(client, token, r_cid, drop_text, reply_markup=TG_MAIN_KEYBOARD)
                                await asyncio.sleep(0.35)
                        except Exception as e:
                            console.print(f"[yellow]⚠️ [TG DropAlerts] {e}[/yellow]")

                    if collected and alert_filter != "OFF":
                        try:
                            all_greens = sorted(
                                [x for x in collected if x.status_code == "GREEN" and str(x.in_stock).startswith("✅")],
                                key=lambda x: x.price,
                            )
                            all_yellows = sorted(
                                [x for x in collected if x.status_code == "YELLOW" and str(x.in_stock).startswith("✅")],
                                key=lambda x: x.price,
                            )
                            abs_min_line = ""
                            if all_greens:
                                cheapest = all_greens[0]
                                abs_p_txt = (
                                    f"<b>Итого {cheapest.price:,} ₽</b> <i>({cheapest.base_price:,} ₽ + пошлина {cheapest.customs_duty:,} ₽)</i>"
                                    if cheapest.customs_duty > 0
                                    else f"<b>{cheapest.price:,} ₽</b> (без пошлины)"
                                )
                                abs_min_line = (
                                    f"💎 <b>Абсолютный минимум:</b> {abs_p_txt} — "
                                    f"<b>[{_h(cheapest.shop)}]</b> {_h(cheapest.title[:55])} | "
                                    f"<a href=\"{_h(clean_product_url(cheapest.url))}\">Открыть ↗</a>\n"
                                )
                                if all_yellows and all_yellows[0].price < cheapest.price * 0.95:
                                    cy = all_yellows[0]
                                    cy_p_txt = (
                                        f"<b>Итого {cy.price:,} ₽</b> <i>({cy.base_price:,} ₽ + пошлина {cy.customs_duty:,} ₽)</i>"
                                        if cy.customs_duty > 0
                                        else f"<b>{cy.price:,} ₽</b>"
                                    )
                                    abs_min_line += (
                                        f"🟡 <b>Дешевле (с нюансом):</b> {cy_p_txt} — "
                                        f"<b>[{_h(cy.shop)}]</b> {_h(cy.title[:45])} <i>({_h((cy.defects_summary or 'проверить')[:42])})</i> | "
                                        f"<a href=\"{_h(clean_product_url(cy.url))}\">Открыть ↗</a>\n"
                                    )
                                abs_min_line += "\n"
                            elif all_yellows:
                                cy = all_yellows[0]
                                cy_p_txt = (
                                    f"<b>Итого {cy.price:,} ₽</b> <i>({cy.base_price:,} ₽ + пошлина {cy.customs_duty:,} ₽)</i>"
                                    if cy.customs_duty > 0
                                    else f"<b>{cy.price:,} ₽</b>"
                                )
                                abs_min_line = (
                                    f"🟡 <b>Минимальная цена (требует проверки):</b> {cy_p_txt} — "
                                    f"<b>[{_h(cy.shop)}]</b> {_h(cy.title[:50])} | "
                                    f"<a href=\"{_h(clean_product_url(cy.url))}\">Открыть ↗</a>\n\n"
                                )
                            best_lines = [
                                f"🔔 <b>Результаты поиска минимальных цен ({_h(rt.last_updated)})</b>",
                                f"🎯 <b>Товар:</b> {_h(', '.join(queries))}",
                                f"📍 <b>Регион:</b> {_h(city_name)} | <b>Проверено:</b> {len(collected)} шт. (🟢 {green_cnt} / 🟡 {yellow_cnt} / 🔴 {red_cnt})\n",
                                abs_min_line + "🏆 <b>Минимальная цена по каждому из 6 магазинов:</b>",
                            ]
                            for sh_name in ("Ozon", "AliExpress", "Яндекс Маркет", "Авито", "DNS", "Wildberries"):
                                sh_items = sorted(
                                    [x for x in collected if x.shop == sh_name and x.status_code == "GREEN" and str(x.in_stock).startswith("✅")],
                                    key=lambda x: x.price,
                                )
                                y_items = sorted(
                                    [x for x in collected if x.shop == sh_name and x.status_code == "YELLOW" and str(x.in_stock).startswith("✅")],
                                    key=lambda x: x.price,
                                )
                                if sh_items:
                                    best = sh_items[0]
                                    best_p_txt = (
                                        f"Итого <b>{best.price:,} ₽</b> <i>({best.base_price:,} + пошлина {best.customs_duty:,} ₽)</i>"
                                        if best.customs_duty > 0
                                        else f"от <b>{best.price:,} ₽</b>"
                                    )
                                    best_lines.append(
                                        f"• 🟢 <b>{_h(sh_name)}</b>: {best_p_txt} — "
                                        f"{_h(best.title[:48])} | <a href=\"{_h(clean_product_url(best.url))}\">Открыть ↗</a>"
                                    )
                                    if y_items and y_items[0].price < best.price * 0.92:
                                        best_y = y_items[0]
                                        by_p_txt = (
                                            f"Итого <b>{best_y.price:,} ₽</b> <i>({best_y.base_price:,} + пошлина {best_y.customs_duty:,} ₽)</i>"
                                            if best_y.customs_duty > 0
                                            else f"<b>{best_y.price:,} ₽</b>"
                                        )
                                        best_lines.append(
                                            f"   ↳ 🟡 <i>дешевле (нюанс):</i> {by_p_txt} — "
                                            f"{_h(best_y.title[:40])} | <a href=\"{_h(clean_product_url(best_y.url))}\">Открыть ↗</a>"
                                        )
                                else:
                                    if y_items:
                                        best_y = y_items[0]
                                        by_p_txt = (
                                            f"Итого <b>{best_y.price:,} ₽</b> <i>({best_y.base_price:,} + пошлина {best_y.customs_duty:,} ₽)</i>"
                                            if best_y.customs_duty > 0
                                            else f"от <b>{best_y.price:,} ₽</b>"
                                        )
                                        best_lines.append(
                                            f"• 🟡 <b>{_h(sh_name)}</b> (проверенный лот): {by_p_txt} — "
                                            f"{_h(best_y.title[:48])} | <a href=\"{_h(clean_product_url(best_y.url))}\">Открыть ↗</a>"
                                        )
                                    else:
                                        sk_map = {nm: k for k, nm, _ in SHOPS_META}
                                        if not shops.get(sk_map.get(sh_name, ""), True):
                                            stat_note = "отключён ⬜"
                                        else:
                                            sc = shop_counts.get(sh_name, 0)
                                            stat_note = "нет в наличии в регионе" if sc == 0 else f"собрано: {sc} шт."
                                        best_lines.append(f"• ⚪ <b>{_h(sh_name)}:</b> {stat_note}")
                            best_txt = "\n".join(best_lines)
                            pq_msg = format_price_quality_comparison_tg(collected, str(target_cid or ""), max_rows=5, u_cfg=cfg)
                            for r_cid in recipients:
                                try:
                                    await tg_send_html_chunks(
                                        client,
                                        token,
                                        r_cid,
                                        best_txt,
                                        reply_markup=build_shops_top_markup(cfg, "ALL"),
                                    )
                                except Exception as e:
                                    console.print(f"[yellow]⚠️ [TG BestLines] chat_id={r_cid}: {e}[/yellow]")
                                try:
                                    await tg_send_html_chunks(
                                        client,
                                        token,
                                        r_cid,
                                        pq_msg,
                                        reply_markup=build_pq_markup(cfg, "ALL"),
                                    )
                                except Exception as e:
                                    console.print(f"[yellow]⚠️ [TG PriceQuality] chat_id={r_cid}: {e}[/yellow]")
                        except Exception as e:
                            console.print(f"[yellow]⚠️ [TG Summary] {e}[/yellow]")

                    caption = (
                        f"📊 Таблица Мониторинг_Товаров.xlsx ({rt.last_updated})\n"
                        f"📍 Регион: {city_name} (Авито: самовывоз ≤{avito_radius} км + Доставка)\n"
                        f"🎯 Товар: {', '.join(queries)}\n"
                        f"🔍 Проверено карточек: {v_stats.get('checked', 0)}/{len(collected)} (пошлин с сайта: {v_stats.get('duties_found', 0)})\n"
                        f"📊 Всего: {len(collected)} поз. (🟢 Зелёные ✅: {green_cnt} | 🟡 Жёлтые ✅: {yellow_cnt} | 🔴 Красные ❌: {red_cnt})\n"
                        f"1. Ozon: {shop_counts['Ozon']} | 2. AliExpress: {shop_counts['AliExpress']} | "
                        f"3. Я.Маркет: {shop_counts['Яндекс Маркет']} | 4. Авито: {shop_counts['Авито']} | "
                        f"5. DNS: {shop_counts['DNS']} | 6. WB: {shop_counts['Wildberries']}"
                    )
                    for r_cid in recipients:
                        try:
                            await send_telegram_excel(client, token, r_cid, caption, excel_file=out_excel_path)
                        except Exception as e:
                            console.print(f"[red]❌ [TG Excel loop] chat_id={r_cid}: {e}[/red]")
            except Exception as e:
                console.print(f"[red]❌ [TG Client] {e}[/red]")
    finally:
        rt.is_scanning = False
        rt.current_step = "Готов к работе"

    return collected


async def main_async(once: bool = False, skip_prompt: bool = False, cloud_minutes: int = 0) -> None:
    loop = asyncio.get_running_loop()

    def _quiet_exception_handler(loop_obj, context):
        exc = context.get("exception")
        if isinstance(exc, ConnectionResetError):
            return
        loop_obj.default_exception_handler(context)

    loop.set_exception_handler(_quiet_exception_handler)

    await rt.user_store.load()
    cfg_boot = load_config()
    startup_cid = str(cfg_boot.get("telegram", {}).get("chat_id") or "")
    if (not startup_cid or not rt.user_store.exists(startup_cid)) and len(rt.user_store.users) == 1:
        startup_cid = next(iter(rt.user_store.users.keys()))
    effective_cfg = {**cfg_boot, **rt.user_store.get(startup_cid)} if (startup_cid and rt.user_store.exists(startup_cid)) else cfg_boot

    rt.items = load_items_from_existing_excel()
    if rt.items:
        ok_xl, _ = _meta_matches_user(rt.global_scan_meta or None, effective_cfg, rt.items, require_today=True)
        if not ok_xl:
            rt.items = []
            rt.global_scan_meta = {}
        elif startup_cid:
            rt.user_items[startup_cid] = rt.items
            rt.user_scan_meta[startup_cid] = dict(rt.global_scan_meta)
    if not rt.items:
        _restore_items_snapshot_from_store(effective_cfg, startup_cid)

    tg_task = None if once else asyncio.create_task(telegram_bot_loop())

    cfg = cfg_boot if (once or skip_prompt or cloud_minutes > 0) else prompt_queries_on_start(cfg_boot)
    chrome_cfg = cfg.get("chrome", {})
    chrome_exec = chrome_cfg.get("executable_path")
    engine = IsolatedChromeEngine(
        chrome_path=chrome_exec if chrome_exec else None,
        debug_port=int(chrome_cfg.get("debug_port", 9234)),
        disable_images=bool(chrome_cfg.get("disable_images", False)),
        bypass_vpn_proxy=bool(chrome_cfg.get("bypass_vpn_proxy", False)),
        minimized=True,
    )

    first_run = True
    cloud_deadline = (time.time() + cloud_minutes * 60) if cloud_minutes > 0 else 0.0
    if cloud_minutes > 0 and rt.items:
        last_scan_ts = float(effective_cfg.get("last_scan") or rt.user_store.meta.get("last_global_scan_ts") or 0.0)
        interval_min = get_effective_interval_minutes(effective_cfg)
        elapsed_sec = (time.time() - last_scan_ts) if last_scan_ts > 0 else 1e9
        has_fresh_same_scope = bool(
            rt.global_scan_meta and _meta_matches_user(rt.global_scan_meta, effective_cfg, rt.items, require_today=True)[0]
        )
        if has_fresh_same_scope and last_scan_ts > 0 and elapsed_sec < interval_min * 60:
            await asyncio.sleep(1.5)
            if not rt.scan_event.is_set() and not rt.scan_queue:
                rem_wait = max(5, int(interval_min * 60 - elapsed_sec))
                if cloud_deadline:
                    rem_cloud = cloud_deadline - time.time()
                    if rem_cloud > 120:
                        wait_first = min(rem_wait, max(5, int(rem_cloud - 90)))
                        console.print(
                            f"[dim]♻️ Данные за сегодня ({effective_cfg.get('region')}) уже актуальны. "
                            f"Ждём до следующего обновления по графику ({format_interval_str(max(1, wait_first // 60))}).[/dim]"
                        )
                        try:
                            await asyncio.wait_for(rt.scan_event.wait(), timeout=wait_first)
                        except asyncio.TimeoutError:
                            pass
                else:
                    try:
                        await asyncio.wait_for(rt.scan_event.wait(), timeout=rem_wait)
                    except asyncio.TimeoutError:
                        pass
    try:
        while True:
            if cloud_deadline and time.time() >= (cloud_deadline - 120):
                break
            rt.scan_event.clear()
            base_cfg = load_config()
            active_cid = rt.scan_queue.pop(0) if rt.scan_queue else ""
            fallback_cid = str(base_cfg.get("telegram", {}).get("chat_id") or "")
            if active_cid and rt.user_store.exists(active_cid):
                u_cfg = rt.user_store.get(active_cid)
                cycle_cfg = {**base_cfg, **u_cfg}
            elif fallback_cid and rt.user_store.exists(fallback_cid):
                active_cid = fallback_cid
                cycle_cfg = {**base_cfg, **rt.user_store.get(fallback_cid)}
            elif len(rt.user_store.users) == 1:
                only_cid = next(iter(rt.user_store.users.keys()))
                active_cid = only_cid
                cycle_cfg = {**base_cfg, **rt.user_store.get(only_cid)}
            else:
                cycle_cfg = base_cfg

            rt.current_scan_task = asyncio.create_task(run_cycle(cycle_cfg, engine, target_cid=active_cid))
            try:
                await rt.current_scan_task
            except asyncio.CancelledError:
                if not rt.scan_event.is_set():
                    raise
                console.print("\n[yellow]⚡ Текущий обход прерван для мгновенного запуска нового запроса из Telegram![/yellow]")
                continue
            if first_run and EXCEL_PATH.exists() and not once and not skip_prompt and cloud_minutes == 0 and sys.platform == "win32":
                first_run = False
                try:
                    os.startfile(str(EXCEL_PATH))
                except Exception:
                    pass
            if once:
                break
            if cloud_deadline and time.time() >= (cloud_deadline - 120):
                break
            if rt.scan_event.is_set() or rt.scan_queue:
                continue
            interval = get_effective_interval_minutes(cycle_cfg)
            wait_sec = interval * 60
            if cloud_deadline:
                rem_cloud = cloud_deadline - time.time()
                if rem_cloud <= 120:
                    break
                wait_sec = min(wait_sec, max(5, int(rem_cloud - 90)))
            console.print(
                f"\n[dim]⏳ Следующее обновление через {format_interval_str(interval)}. "
                f"(Управление в Telegram-боте активно 24/7)[/dim]"
            )
            try:
                await asyncio.wait_for(rt.scan_event.wait(), timeout=wait_sec)
            except asyncio.TimeoutError:
                if cloud_deadline and time.time() >= (cloud_deadline - 150):
                    break
                if active_cid and rt.user_store.exists(active_cid):
                    advance_ladder_step(rt.user_store.get(active_cid))
                    rt.user_store.save()
                elif is_auto_interval(base_cfg):
                    advance_ladder_step(base_cfg)
                    save_config(base_cfg)
            if cloud_deadline and time.time() >= (cloud_deadline - 120):
                break
    finally:
        if tg_task:
            tg_task.cancel()
        engine.close()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Мониторинг 6 маркетплейсов в Excel + Telegram-бот")
    parser.add_argument("--once", action="store_true", help="Выполнить один цикл сбора и выйти")
    parser.add_argument("--bg", action="store_true", help="Запуск без стартового вопроса в консоли")
    parser.add_argument("--cloud-minutes", type=int, default=0, help="Время облачной сессии в минутах (для GitHub Actions)")
    args = parser.parse_args()

    try:
        asyncio.run(main_async(once=args.once, skip_prompt=args.bg, cloud_minutes=args.cloud_minutes))
    except KeyboardInterrupt:
        console.print("\n[yellow]Мониторинг остановлен пользователем.[/yellow]")
        sys.exit(0)
