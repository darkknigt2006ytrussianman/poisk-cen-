"""
Персональные настройки каждого пользователя Telegram-бота (по chat_id).

Хранение:
  • локально — JSON во временной папке (работает всегда);
  • в облаке — секретный GitHub Gist, если задан секрет STATE_GIST_TOKEN
    (токен GitHub только с правом «gist»). Так настройки переживают перезапуск
    облачной сессии и не попадают в публичный репозиторий.
"""
import asyncio
import hashlib
import hmac
import json
import os
import secrets
import tempfile
import time
import zlib
from pathlib import Path
from typing import Dict, Optional

import httpx

GIST_FILENAME = "poisk_bot_state.json"
GIST_DESCRIPTION = "poisk-bot private state"

DEFAULT_USER = {
    "queries": [],
    "selected_queries": [],
    "min_price": 100,
    "max_price": 250000,
    "price_is_custom": False,
    "region": "Все регионы РФ",
    "region_slug_avito": "all",
    "ymarket_lr": 213,
    "avito_mode": "both",       # both | pickup | delivery
    "avito_radius_km": 15,      # 5..15 км
    "interval_minutes": 180,
    "use_ladder": False,
    "ladder_step": 0,
    "last_announced_version": "",
    "welcomed": False,
    "alert_filter": "GREEN_ONLY",  # GREEN_ONLY | ALL | OFF
    "muted_queries": [],
    "shops_priority": {
        "ozon": True,
        "aliexpress": True,
        "ymarket": True,
        "avito": True,
        "dns": True,
        "wb": True,
    },
    "lat": None,
    "lon": None,
    "detected_place": "",
    "last_scan": 0.0,
    "last_scan_region": "",
    "last_scan_queries": [],
    "last_seen": 0.0,
}


def get_cache_dir() -> Path:
    custom = os.environ.get("BOT_CACHE_DIR", "").strip()
    if custom:
        p = Path(custom)
        try:
            p.mkdir(parents=True, exist_ok=True)
            return p
        except Exception:
            pass
    return Path(tempfile.gettempdir())


def _derive_keys(secret: str, salt_label: str = "poisk_state_v1") -> tuple[bytes, bytes]:
    enc_key = hashlib.sha256(f"{salt_label}:enc:{secret}".encode("utf-8")).digest()
    mac_key = hashlib.sha256(f"{salt_label}:mac:{secret}".encode("utf-8")).digest()
    return enc_key, mac_key


def _keystream(enc_key: bytes, nonce: bytes, length: int) -> bytes:
    out = bytearray()
    counter = 0
    while len(out) < length:
        out.extend(hashlib.sha256(enc_key + nonce + counter.to_bytes(4, "big")).digest())
        counter += 1
    return bytes(out[:length])


def encrypt_blob(raw_data: bytes, secret: str) -> bytes:
    """Шифрует сжатый JSON локальным секретным ключом (CTR-SHA256 + HMAC-SHA256)."""
    compressed = zlib.compress(raw_data, level=6)
    enc_key, mac_key = _derive_keys(secret)
    nonce = secrets.token_bytes(16)
    ks = _keystream(enc_key, nonce, len(compressed))
    cipher = bytes(a ^ b for a, b in zip(compressed, ks))
    tag = hmac.new(mac_key, b"v1" + nonce + cipher, hashlib.sha256).digest()
    return b"PS1" + nonce + tag + cipher


def decrypt_blob(blob: bytes, secret: str) -> Optional[bytes]:
    if not blob or len(blob) < 51 or not blob.startswith(b"PS1"):
        return None
    nonce = blob[3:19]
    tag = blob[19:51]
    cipher = blob[51:]
    enc_key, mac_key = _derive_keys(secret)
    expected = hmac.new(mac_key, b"v1" + nonce + cipher, hashlib.sha256).digest()
    if not hmac.compare_digest(tag, expected):
        return None
    ks = _keystream(enc_key, nonce, len(cipher))
    plain_comp = bytes(a ^ b for a, b in zip(cipher, ks))
    try:
        return zlib.decompress(plain_comp)
    except Exception:
        return None


class UserStore:
    def __init__(self, local_path: Optional[Path] = None):
        cache_dir = get_cache_dir()
        self.local_path = local_path or (cache_dir / "poisk_bot_users.json")
        self.enc_path = self.local_path.parent / "poisk_bot_state.enc"
        self.items_enc_path = self.local_path.parent / "poisk_items_cache.enc"
        self.users: Dict[str, dict] = {}
        self.meta: dict = {}
        self.gist_token = os.environ.get("STATE_GIST_TOKEN", "").strip()
        self.gist_id = os.environ.get("STATE_GIST_ID", "").strip()
        self._push_task: Optional[asyncio.Task] = None

    def _split_payload(self, data: Optional[dict]) -> None:
        if not isinstance(data, dict):
            return
        meta = data.pop("__meta__", {})
        if isinstance(meta, dict):
            self.meta = meta
        self.users = {str(k): v for k, v in data.items() if isinstance(v, dict)}

    def _build_payload(self) -> dict:
        out: dict = dict(self.users)
        if self.meta:
            out["__meta__"] = dict(self.meta)
        return out

    @staticmethod
    def _payload_ts(data: Optional[dict]) -> float:
        if not isinstance(data, dict):
            return 0.0
        m = data.get("__meta__") or {}
        try:
            return float(m.get("updated_at") or 0.0)
        except Exception:
            return 0.0

    # ---------- загрузка / сохранение ----------
    async def load(self) -> None:
        best_data: Optional[dict] = None
        best_ts: float = -1.0

        def _consider(raw: Optional[dict]) -> None:
            nonlocal best_data, best_ts
            if not isinstance(raw, dict) or not raw:
                return
            ts = self._payload_ts(raw)
            if best_data is None or ts >= best_ts:
                best_data = raw
                best_ts = ts

        if self.local_path.exists():
            try:
                _consider(json.loads(self.local_path.read_text(encoding="utf-8")) or {})
            except Exception:
                pass

        enc_secret = os.environ.get("TG_BOT_TOKEN", "").strip()
        if enc_secret and self.enc_path.exists():
            try:
                plain = decrypt_blob(self.enc_path.read_bytes(), enc_secret)
                if plain:
                    _consider(json.loads(plain.decode("utf-8")) or {})
            except Exception:
                pass

        if self.gist_token:
            remote = await self._pull_gist()
            _consider(remote)

        if best_data is not None:
            self._split_payload(json.loads(json.dumps(best_data)))
            self._write_local()

    def _write_local(self) -> None:
        payload = self._build_payload()
        raw_str = json.dumps(payload, ensure_ascii=False, indent=1)
        try:
            self.local_path.parent.mkdir(parents=True, exist_ok=True)
            self.local_path.write_text(raw_str, encoding="utf-8")
        except Exception:
            pass
        enc_secret = os.environ.get("TG_BOT_TOKEN", "").strip()
        if enc_secret:
            try:
                self.enc_path.write_bytes(encrypt_blob(raw_str.encode("utf-8"), enc_secret))
            except Exception:
                pass

    def save(self) -> None:
        self.meta["updated_at"] = time.time()
        self._write_local()
        if not self.gist_token:
            return
        try:
            loop = asyncio.get_running_loop()
        except RuntimeError:
            return
        if self._push_task and not self._push_task.done():
            return  # отправка уже запланирована, в неё попадут свежие данные
        self._push_task = loop.create_task(self._debounced_push())

    def save_items_snapshot(self, snapshot: dict) -> None:
        enc_secret = os.environ.get("TG_BOT_TOKEN", "").strip()
        if not enc_secret:
            return
        try:
            self.items_enc_path.parent.mkdir(parents=True, exist_ok=True)
            raw = json.dumps(snapshot, ensure_ascii=False).encode("utf-8")
            self.items_enc_path.write_bytes(encrypt_blob(raw, enc_secret))
        except Exception:
            pass

    def load_items_snapshot(self) -> Optional[dict]:
        enc_secret = os.environ.get("TG_BOT_TOKEN", "").strip()
        if not enc_secret or not self.items_enc_path.exists():
            return None
        try:
            plain = decrypt_blob(self.items_enc_path.read_bytes(), enc_secret)
            if plain:
                data = json.loads(plain.decode("utf-8"))
                if isinstance(data, dict):
                    return data
        except Exception:
            pass
        return None

    async def _debounced_push(self) -> None:
        await asyncio.sleep(5.0)
        await self._push_gist()

    # ---------- доступ к пользователям ----------
    def exists(self, cid: str) -> bool:
        return str(cid) in self.users

    def get(self, cid: str) -> dict:
        cid = str(cid)
        if cid not in self.users:
            self.users[cid] = json.loads(json.dumps(DEFAULT_USER))
        u = self.users[cid]
        for k, v in DEFAULT_USER.items():
            if k not in u or u[k] is None:
                u[k] = json.loads(json.dumps(v))
            elif isinstance(v, dict) and isinstance(u[k], dict):
                for dk, dv in v.items():
                    u[k].setdefault(dk, dv)
        if u.get("queries") and not u.get("selected_queries") and "selected_queries" not in u:
            u["selected_queries"] = list(u["queries"])
        return u

    def touch(self, cid: str) -> dict:
        u = self.get(cid)
        u["last_seen"] = time.time()
        return u

    # ---------- GitHub Gist ----------
    def _gh_headers(self) -> dict:
        return {
            "Authorization": f"token {self.gist_token}",
            "Accept": "application/vnd.github+json",
            "User-Agent": "poisk-bot",
        }

    async def _find_gist_id(self, client: httpx.AsyncClient) -> str:
        if self.gist_id:
            return self.gist_id
        for page in range(1, 4):
            r = await client.get("https://api.github.com/gists", params={"per_page": 100, "page": page}, headers=self._gh_headers())
            if r.status_code != 200:
                return ""
            gists = r.json()
            for g in gists:
                if GIST_FILENAME in (g.get("files") or {}):
                    self.gist_id = g["id"]
                    return self.gist_id
            if len(gists) < 100:
                break
        return ""

    async def _pull_gist(self) -> Optional[dict]:
        try:
            async with httpx.AsyncClient(timeout=15.0) as client:
                gid = await self._find_gist_id(client)
                if not gid:
                    return None
                r = await client.get(f"https://api.github.com/gists/{gid}", headers=self._gh_headers())
                if r.status_code != 200:
                    return None
                f = (r.json().get("files") or {}).get(GIST_FILENAME) or {}
                content = f.get("content")
                if f.get("truncated") and f.get("raw_url"):
                    content = (await client.get(f["raw_url"], headers=self._gh_headers())).text
                return json.loads(content or "{}")
        except Exception:
            return None

    async def _push_gist(self) -> None:
        payload = {"files": {GIST_FILENAME: {"content": json.dumps(self._build_payload(), ensure_ascii=False, indent=1)}}}
        try:
            async with httpx.AsyncClient(timeout=15.0) as client:
                gid = await self._find_gist_id(client)
                if gid:
                    await client.patch(f"https://api.github.com/gists/{gid}", json=payload, headers=self._gh_headers())
                else:
                    r = await client.post(
                        "https://api.github.com/gists",
                        json={**payload, "description": GIST_DESCRIPTION, "public": False},
                        headers=self._gh_headers(),
                    )
                    if r.status_code in (200, 201):
                        self.gist_id = r.json().get("id", "")
        except Exception:
            pass
