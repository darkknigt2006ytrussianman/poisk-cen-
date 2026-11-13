import asyncio
import json
import os
import shutil
import subprocess
import tempfile
import time
from pathlib import Path
from typing import Optional, Any, Set, Dict

import httpx
import websockets

from wifi_bridge import DirectWifiProxyBridge


class IsolatedChromeEngine:
    """
    Управляет отдельным изолированным экземпляром Chromium-Gost / Chrome.
    Маршрутизирует трафик через DirectWifiProxyBridge (NET_NODE РФ-прокси или прямой интерфейс).
    """

    def __init__(
        self,
        chrome_path: Optional[str] = None,
        profile_dir: Optional[Path] = None,
        debug_port: int = 9234,
        disable_images: bool = False,
        bypass_vpn_proxy: bool = False,
        minimized: bool = True,
        force_no_proxy: bool = False,
    ):
        self.chrome_path = self._resolve_chrome_path(chrome_path)
        self.profile_dir = profile_dir or (Path(tempfile.gettempdir()) / f"chrome_bot_profile_{debug_port}")
        self.debug_port = debug_port
        self.disable_images = disable_images
        self.bypass_vpn_proxy = False if force_no_proxy else bypass_vpn_proxy
        self.force_no_proxy = bool(force_no_proxy)
        self.minimized = minimized
        self.process: Optional[subprocess.Popen] = None
        has_env_node = bool(os.environ.get("NET_NODE", "").strip())
        self.bridge = None if force_no_proxy else (DirectWifiProxyBridge(listen_port=debug_port + 100) if (bypass_vpn_proxy or has_env_node) else None)
        self._warmed_domains: Set[str] = set()
        self._msg_id = 0
        self.last_page_diag: Dict[str, Any] = {}
        self._active_region_key: str = ""
        self._region_coords: Optional[tuple[float, float]] = None
        self._wb_region_applied: str = ""

    @staticmethod
    def _resolve_chrome_path(default_path: Optional[str] = None) -> str:
        which_gost = shutil.which("chromium-gost")
        which_chrome = shutil.which("google-chrome") or shutil.which("chromium") or shutil.which("chromium-browser")
        candidates = [
            default_path,
            os.environ.get("CHROME_PATH"),
            os.path.expandvars(r"%LOCALAPPDATA%\chromium-gost\chrome.exe"),
            os.path.expandvars(r"%LOCALAPPDATA%\chromium-gost\Application\chrome.exe"),
            r"C:\Program Files\chromium-gost\chrome.exe",
            r"C:\Program Files (x86)\chromium-gost\chrome.exe",
            r"C:\chromium-gost\chrome.exe",
            which_gost,
            "/usr/bin/chromium-gost",
            "/usr/bin/chromium-gost-stable",
            "/opt/chromium-gost/chrome",
            "/opt/chromium-gost/chromium-gost",
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
            which_chrome,
            "/usr/bin/google-chrome",
            "/usr/bin/google-chrome-stable",
            "/usr/bin/chromium",
            "/usr/bin/chromium-browser",
        ]
        for path in candidates:
            if path and os.path.exists(path):
                return path
        raise FileNotFoundError("Не найден исполняемый файл Chromium-Gost / Google Chrome / Chromium.")

    def is_running(self) -> bool:
        try:
            with httpx.Client(trust_env=False, timeout=0.4) as client:
                resp = client.get(f"http://127.0.0.1:{self.debug_port}/json/version")
                return resp.status_code == 200
        except Exception:
            return False

    async def _is_running_async(self) -> bool:
        try:
            async with httpx.AsyncClient(trust_env=False, timeout=0.6) as client:
                resp = await client.get(f"http://127.0.0.1:{self.debug_port}/json/version")
                return resp.status_code == 200
        except Exception:
            return False

    async def ensure_started(self) -> None:
        use_bridge = bool(
            self.bridge
            and (
                getattr(self.bridge, "gateway_node", None)
                or (self.bypass_vpn_proxy and os.name == "nt" and (getattr(self.bridge, "lan_ip", None) or getattr(self.bridge, "if_index", None)))
            )
        )
        if use_bridge and self.bridge:
            await self.bridge.start()

        if await self._is_running_async():
            return

        self.profile_dir.mkdir(parents=True, exist_ok=True)
        has_node = bool(self.bridge and getattr(self.bridge, "gateway_node", None))
        print(f"[Браузер] Запуск {Path(self.chrome_path).name} (NET_NODE={'да' if has_node else 'нет'}, мост={'вкл' if use_bridge else 'прямой'})", flush=True)

        args = [
            self.chrome_path,
            f"--remote-debugging-port={self.debug_port}",
            f"--user-data-dir={self.profile_dir.resolve()}",
            "--no-first-run",
            "--no-default-browser-check",
            "--disable-blink-features=AutomationControlled",
            "--mute-audio",
            "--remote-allow-origins=*",
            "--lang=ru-RU,ru",
        ]

        if self.disable_images:
            args.append("--blink-settings=imagesEnabled=false")

        if os.name != "nt":
            args.extend([
                "--no-sandbox",
                "--disable-dev-shm-usage",
                "--use-gl=angle",
                "--use-angle=swiftshader",
                "--enable-webgl",
                "--ignore-gpu-blocklist",
            ])
            if not os.environ.get("DISPLAY"):
                args.append("--headless=new")

        if use_bridge and self.bridge:
            args.append(f"--proxy-server={self.bridge.proxy_url}")
        elif os.name == "nt" or self.force_no_proxy:
            args.append("--no-proxy-server")

        args.append("--window-size=1366,900")
        args.append("about:blank")

        creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        self.process = subprocess.Popen(
            args,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=creation_flags,
        )

        for _ in range(30):
            if await self._is_running_async():
                return
            await asyncio.sleep(0.3)

        raise RuntimeError(f"Отдельный Chrome не ответил на порту {self.debug_port}")

    async def _get_or_create_tab_ws(self) -> str:
        await self.ensure_started()
        async with httpx.AsyncClient(trust_env=False, timeout=5.0) as client:
            resp = await client.get(f"http://127.0.0.1:{self.debug_port}/json/list")
            tabs = resp.json()
            for tab in tabs:
                if tab.get("type") == "page" and "webSocketDebuggerUrl" in tab:
                    return tab["webSocketDebuggerUrl"]

            new_resp = await client.put(f"http://127.0.0.1:{self.debug_port}/json/new?about:blank")
            return new_resp.json()["webSocketDebuggerUrl"]

    async def create_tab_ws(self) -> tuple[str, str]:
        """Создаёт выделенную вкладку Chrome для параллельного воркера и возвращает (tab_id, ws_url)."""
        await self.ensure_started()
        async with httpx.AsyncClient(trust_env=False, timeout=5.0) as client:
            new_resp = await client.put(f"http://127.0.0.1:{self.debug_port}/json/new?about:blank")
            data = new_resp.json()
            tab_id = str(data.get("id", ""))
            ws_url = str(data["webSocketDebuggerUrl"])
        if self._region_coords is not None:
            try:
                async with websockets.connect(ws_url, max_size=4 * 1024 * 1024, open_timeout=4.0, close_timeout=1.5) as ws:
                    await self._cdp_call(
                        ws,
                        "Emulation.setGeolocationOverride",
                        {"latitude": self._region_coords[0], "longitude": self._region_coords[1], "accuracy": 100},
                        timeout=2.5,
                    )
            except Exception:
                pass
        return tab_id, ws_url

    async def close_tab_ws(self, tab_id: str) -> None:
        if not tab_id:
            return
        try:
            async with httpx.AsyncClient(trust_env=False, timeout=3.0) as client:
                await client.get(f"http://127.0.0.1:{self.debug_port}/json/close/{tab_id}")
        except Exception:
            pass

    async def _cdp_call(self, ws, method: str, params: Optional[dict] = None, timeout: float = 12.0) -> dict:
        self._msg_id += 1
        call_id = self._msg_id
        payload = {"id": call_id, "method": method, "params": params or {}}
        await ws.send(json.dumps(payload))

        deadline = time.monotonic() + timeout
        while True:
            rem = max(0.2, deadline - time.monotonic())
            raw = await asyncio.wait_for(ws.recv(), timeout=rem)
            msg = json.loads(raw)
            if msg.get("id") == call_id:
                return msg

    async def _wait_page_ready(self, ws, deadline: float, base_wait: float, ready_selector: str) -> None:
        started = time.monotonic()
        min_init = 0.5 if base_wait <= 2.0 else 1.5
        initial = min(max(min_init, base_wait), 4.0, max(0.5, deadline - time.monotonic() - 4.0))
        await asyncio.sleep(initial)
        clicked_reload = False
        probe_js = (
            "(() => {"
            "const t = (document.title || '').toLowerCase();"
            "const txt = (document.body ? document.body.innerText : '') || '';"
            "const low = txt.slice(0, 400).toLowerCase();"
            "const btn = Array.from(document.querySelectorAll('button, a')).find(b => (b.innerText || '').trim().toLowerCase().includes('обновить'));"
            f"const cnt = {json.dumps(ready_selector)} ? document.querySelectorAll({json.dumps(ready_selector)}).length : 0;"
            "return { ready: document.readyState, t, low, hasReload: !!btn, cnt, len: txt.length };"
            "})()"
        )
        # Ждём прогрузки или прохождения Antibot Challenge в пределах отведённого окна fetch_page_js
        extra_until = deadline - 3.5
        while time.monotonic() < extra_until:
            try:
                res = await self._cdp_call(ws, "Runtime.evaluate", {"expression": probe_js, "returnByValue": True}, timeout=4.0)
                st = res.get("result", {}).get("result", {}).get("value") or {}
            except Exception:
                await asyncio.sleep(1.0)
                continue

            title_low = st.get("t", "")
            low = st.get("low", "")
            is_hard_block = (
                "проблема с ip" in low
                or "доступ ограничен" in title_low
                or "доступ ограничен" in low
                or "access denied" in title_low
            )
            if is_hard_block and st.get("ready") in ("interactive", "complete") and time.monotonic() - started > 2.2:
                return
            is_challenge = (
                "antibot" in title_low
                or "challenge" in title_low
                or "just a moment" in title_low
                or "проверка браузера" in low
                or "подтвердите" in low
                or "подождите" in title_low
            )
            is_net_err = "похоже, нет соединения" in title_low or "выключите vpn" in low
            if st.get("hasReload") and is_net_err:
                if not clicked_reload:
                    clicked_reload = True
                    try:
                        await self._cdp_call(
                            ws,
                            "Runtime.evaluate",
                            {"expression": "Array.from(document.querySelectorAll('button, a')).find(b => (b.innerText || '').trim().toLowerCase().includes('обновить'))?.click()", "returnByValue": True},
                            timeout=3.0,
                        )
                    except Exception:
                        pass
                    await asyncio.sleep(min(5.0, max(1.0, extra_until - time.monotonic())))
                    continue

            if is_challenge or is_net_err:
                await asyncio.sleep(1.5)
                continue

            if ready_selector:
                if st.get("cnt", 0) >= 3:
                    return
                if st.get("ready") == "complete" and time.monotonic() + 3.0 < extra_until and st.get("cnt", 0) == 0:
                    # React/SPA (например WB) может рендерить карточки через 2–4с после complete
                    await asyncio.sleep(1.5)
                    continue
                if st.get("ready") == "complete":
                    return
            else:
                if st.get("ready") == "complete" and st.get("len", 0) > 80:
                    return
            await asyncio.sleep(1.2)

    async def clear_domain_cookies(self, domain_substr: str) -> int:
        """Удаляет cookies только для указанного домена без сброса остальных магазинов."""
        key = (domain_substr or "").strip().lower()
        if not key:
            return 0
        removed = 0
        try:
            ws_url = await self._get_or_create_tab_ws()
            async with websockets.connect(ws_url, max_size=8 * 1024 * 1024, open_timeout=6.0, close_timeout=2.0) as ws:
                await self._cdp_call(ws, "Network.enable", {}, timeout=3.0)
                res = await self._cdp_call(ws, "Network.getAllCookies", {}, timeout=4.0)
                cookies = res.get("result", {}).get("cookies", []) or []
                for c in cookies:
                    dom = (c.get("domain") or "").lower()
                    if key in dom:
                        await self._cdp_call(
                            ws,
                            "Network.deleteCookies",
                            {"name": c.get("name", ""), "domain": c.get("domain", ""), "path": c.get("path", "/")},
                            timeout=2.5,
                        )
                        removed += 1
        except Exception:
            pass
        return removed

    async def apply_region_context(
        self,
        city_name: str = "Все регионы РФ",
        ymarket_lr: int = 213,
        lat: Optional[float] = None,
        lon: Optional[float] = None,
    ) -> None:
        """Настраивает геолокацию браузера и региональные cookies под выбранный город/регион."""
        key = f"{city_name}|{int(ymarket_lr or 213)}|{round(lat, 3) if lat is not None else ''}|{round(lon, 3) if lon is not None else ''}"
        if key == self._active_region_key:
            return
        self._active_region_key = key
        self._region_coords = (float(lat), float(lon)) if (lat is not None and lon is not None) else None
        self._wb_region_applied = ""
        try:
            ws_url = await self._get_or_create_tab_ws()
            async with websockets.connect(ws_url, max_size=8 * 1024 * 1024, open_timeout=6.0, close_timeout=2.0) as ws:
                if self._region_coords is not None:
                    try:
                        await self._cdp_call(ws, "Browser.grantPermissions", {"permissions": ["geolocation"]}, timeout=3.0)
                        await self._cdp_call(
                            ws,
                            "Emulation.setGeolocationOverride",
                            {"latitude": self._region_coords[0], "longitude": self._region_coords[1], "accuracy": 100},
                            timeout=3.0,
                        )
                    except Exception:
                        pass
                await self._cdp_call(ws, "Network.enable", {}, timeout=3.0)
                lr_val = str(int(ymarket_lr or 213))
                for dom in (".yandex.ru", ".market.yandex.ru"):
                    await self._cdp_call(
                        ws,
                        "Network.setCookie",
                        {"name": "yandex_gid", "value": lr_val, "domain": dom, "path": "/"},
                        timeout=2.5,
                    )
                await self._cdp_call(
                    ws,
                    "Network.setCookie",
                    {
                        "name": "aep_usuc_f",
                        "value": "b_locale=ru_RU&c_tp=RUB&region=RU&site=rus",
                        "domain": ".aliexpress.ru",
                        "path": "/",
                    },
                    timeout=2.5,
                )
        except Exception:
            pass

    async def warmup_url(
        self,
        url: str,
        wait_sec: float = 8.0,
        ready_selector: str = "a[href]",
    ) -> None:
        from urllib.parse import urlsplit
        host = (urlsplit(url).hostname or "").lower()
        if not host or host in self._warmed_domains:
            return
        self._warmed_domains.add(host)
        try:
            await self.fetch_page_js(
                url,
                "document.title || ''",
                wait_sec=wait_sec,
                scroll=False,
                ready_selector=ready_selector,
                timeout_sec=28.0,
            )
        except Exception:
            pass

    async def _fetch_once(
        self,
        url: str,
        js_extractor: str,
        wait_sec: float,
        scroll: bool,
        ready_selector: str,
        timeout_sec: float,
        tab_ws: Optional[str] = None,
    ) -> Any:
        deadline = time.monotonic() + timeout_sec
        ws_url = tab_ws or await self._get_or_create_tab_ws()
        async with websockets.connect(ws_url, max_size=32 * 1024 * 1024, open_timeout=8.0, close_timeout=2.0) as ws:
            try:
                await self._cdp_call(ws, "Page.enable", {}, timeout=4.0)
            except Exception:
                pass
            nav_timeout = min(16.0, max(5.0, deadline - time.monotonic() - 4.0))
            await self._cdp_call(ws, "Page.navigate", {"url": url}, timeout=nav_timeout)
            await self._wait_page_ready(ws, deadline, wait_sec, ready_selector)

            if scroll and time.monotonic() < deadline - 4.0:
                for _ in range(4):
                    if time.monotonic() >= deadline - 3.0:
                        break
                    try:
                        await self._cdp_call(
                            ws,
                            "Runtime.evaluate",
                            {"expression": "window.scrollBy(0, 950);", "returnByValue": True},
                            timeout=3.0,
                        )
                    except Exception:
                        pass
                    await asyncio.sleep(0.35)
                await asyncio.sleep(0.5)

            eval_timeout = min(14.0, max(4.0, deadline - time.monotonic()))
            eval_res = await self._cdp_call(
                ws,
                "Runtime.evaluate",
                {
                    "expression": js_extractor,
                    "awaitPromise": True,
                    "returnByValue": True,
                },
                timeout=eval_timeout,
            )
            result_obj = eval_res.get("result", {}).get("result", {})
            val = result_obj.get("value")
            if not val and time.monotonic() + 1.5 < deadline:
                try:
                    d_res = await self._cdp_call(
                        ws,
                        "Runtime.evaluate",
                        {
                            "expression": (
                                "(() => ({ title: document.title || '', url: location.href || '', "
                                "ready: document.readyState || '', "
                                "links: document.querySelectorAll('a[href]').length, "
                                "head: ((document.body ? document.body.innerText : '') || '').replace(/\\s+/g, ' ').trim().slice(0, 140) }))()"
                            ),
                            "returnByValue": True,
                        },
                        timeout=2.5,
                    )
                    self.last_page_diag = d_res.get("result", {}).get("result", {}).get("value") or {}
                except Exception:
                    self.last_page_diag = {"url": url, "title": "diag_error"}
            elif val:
                self.last_page_diag = {}
            return val

    async def fetch_page_js(
        self,
        url: str,
        js_extractor: str,
        wait_sec: float = 4.5,
        scroll: bool = True,
        ready_selector: str = "",
        timeout_sec: float = 30.0,
        retries: int = 2,
        tab_ws: Optional[str] = None,
        allow_restart: bool = True,
    ) -> Any:
        attempts = max(1, int(retries or 1))
        for attempt in range(attempts):
            step_timeout = timeout_sec if attempt == 0 else min(24.0, timeout_sec)
            try:
                return await asyncio.wait_for(
                    self._fetch_once(url, js_extractor, wait_sec, scroll, ready_selector, step_timeout, tab_ws=tab_ws),
                    timeout=step_timeout + 1.0,
                )
            except Exception as exc:
                self.last_page_diag = {
                    "url": url,
                    "title": type(exc).__name__,
                    "head": str(exc)[:140],
                }
                if attempt + 1 < attempts:
                    if allow_restart and not tab_ws:
                        self._terminate_browser_only()
                    await asyncio.sleep(0.8)
                else:
                    return []

    def _terminate_browser_only(self) -> None:
        if self.process and self.process.poll() is None:
            try:
                self.process.terminate()
                try:
                    self.process.wait(timeout=2.0)
                except Exception:
                    self.process.kill()
            except Exception:
                pass
        self.process = None

    def close(self) -> None:
        self._terminate_browser_only()
        if self.bridge and getattr(self.bridge, "_server", None):
            try:
                self.bridge._server.close()
                self.bridge._server = None
            except Exception:
                pass


# Псевдоним для российского оптимизированного браузера на базе Chromium-Gost
ChromiumGhostEngine = IsolatedChromeEngine

