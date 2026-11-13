import asyncio
import base64
import ctypes
import os
import re
import socket
import struct
from typing import Optional, Dict, Tuple, List
from urllib.parse import urlsplit, unquote

IP_UNICAST_IF = 31


def _parse_single_node(raw_node: str) -> Optional[dict]:
    s = (raw_node or "").strip()
    if not s:
        return None
    scheme = "auto"
    if "://" in s:
        p = urlsplit(s)
        scheme = (p.scheme or "auto").lower()
        if scheme.startswith("socks"):
            scheme = "socks5"
        elif scheme in ("http", "https"):
            scheme = "http"
        if p.hostname and p.port:
            return {
                "scheme": scheme,
                "host": p.hostname,
                "port": int(p.port),
                "user": unquote(p.username or ""),
                "password": unquote(p.password or ""),
            }
        s = s.split("://", 1)[1]

    if "@" in s:
        left, right = s.split("@", 1)
        # Поддержка user:pass@host:port И host:port@user:pass
        if re.match(r"^[^:]+:\d+$", right.strip()):
            host, port_s = right.strip().rsplit(":", 1)
            user, _, pw = left.partition(":")
            return {"scheme": scheme, "host": host, "port": int(port_s), "user": user, "password": pw}
        if re.match(r"^[^:]+:\d+$", left.strip()):
            host, port_s = left.strip().rsplit(":", 1)
            user, _, pw = right.partition(":")
            return {"scheme": scheme, "host": host, "port": int(port_s), "user": user, "password": pw}

    parts = s.split(":")
    if len(parts) >= 2 and parts[1].strip().isdigit():
        return {
            "scheme": scheme,
            "host": parts[0].strip(),
            "port": int(parts[1].strip()),
            "user": parts[2].strip() if len(parts) >= 3 else "",
            "password": ":".join(parts[3:]).strip() if len(parts) >= 4 else "",
        }
    return None


def parse_gateway_nodes(raw: str) -> List[dict]:
    nodes: List[dict] = []
    for chunk in re.split(r"[\r\n,;]+", raw or ""):
        parsed = _parse_single_node(chunk)
        if parsed and parsed.get("host") and parsed.get("port"):
            nodes.append(parsed)
    return nodes


def detect_physical_lan_interface() -> Tuple[Optional[str], Optional[int]]:
    """
    Определяет IPv4-адрес и аппаратный InterfaceIndex сетевого адаптера.
    """
    lan_ip: Optional[str] = None
    candidates = []
    try:
        _, _, ips = socket.gethostbyname_ex(socket.gethostname())
        for ip in ips:
            if (
                ip.startswith("127.")
                or ip.startswith("169.254.")
                or ip.startswith("198.18.")
                or ip.startswith("198.19.")
            ):
                continue
            if ip.startswith("192.168."):
                lan_ip = ip
                break
            candidates.append(ip)
        if not lan_ip and candidates:
            lan_ip = candidates[0]
    except Exception:
        pass

    if not lan_ip:
        return None, None

    # Определяем InterfaceIndex через GetBestInterfaceEx / iphlpapi или сопоставление
    if_index: Optional[int] = None
    try:
        iphlpapi = ctypes.windll.iphlpapi
        # Структура sockaddr_in для шлюза локальной подсети
        gw_parts = lan_ip.split(".")
        gw_ip = f"{gw_parts[0]}.{gw_parts[1]}.{gw_parts[2]}.1"
        dest_addr = struct.unpack("<I", socket.inet_aton(gw_ip))[0]
        best_idx = ctypes.c_ulong(0)
        res = iphlpapi.GetBestInterface(dest_addr, ctypes.byref(best_idx))
        if res == 0 and best_idx.value > 0:
            if_index = int(best_idx.value)
    except Exception:
        pass

    return lan_ip, if_index


_AD_TRACKER_SUFFIXES = (
    "doubleclick.net",
    "googlesyndication.com",
    "googleadservices.com",
    "adriver.ru",
    "hybrid.ai",
    "otm-r.com",
    "upravel.com",
    "buzzoola.com",
    "sape.ru",
    "acint.net",
    "adhigh.net",
    "soloway.ru",
    "targetads.io",
    "bidster.net",
    "sspnet.tech",
    "onetarget.ru",
    "bidderstack.com",
    "nitrortb.ru",
    "dynotech.io",
    "utraff.com",
    "getads.ru",
    "adspector.io",
    "adspend.space",
    "opendsp.ru",
    "bidvol.com",
    "lotus-dsp.ru",
    "abdm.ru",
    "melvad.com",
    "kombinat.digital",
    "programmatica.com",
    "giftclick.ru",
    "weborama.fr",
    "weborama.com",
    "silvermob.com",
    "adnxs.com",
    "al-adtech.com",
    "adtec.ru",
)

_PROTECTED_RU_SHOP_SUFFIXES = (
    "ozon.ru",
    "ozone.ru",
    "ozonusercontent.com",
    "wildberries.ru",
    "wbbasket.ru",
    "wbstatic.net",
    "geobasket.ru",
    "dns-shop.ru",
    "market.yandex.ru",
    "yandex.ru",
    "aliexpress.ru",
)


def _host_matches_suffix(host: str, suffixes: Tuple[str, ...]) -> bool:
    """Проверяет, совпадает ли хост с доменным суффиксом без ложных подстрок."""
    h = (host or "").strip().lower().rstrip(".")
    return bool(h) and any(h == s or h.endswith("." + s) for s in suffixes)


class DirectWifiProxyBridge:
    """
    Локальный сетевой мост маршрутизации для изолированного браузера Chrome (Все регионы РФ).
    """

    def __init__(self, listen_host: str = "127.0.0.1", listen_port: int = 9334, gateway_node: Optional[str] = None):
        self.listen_host = listen_host
        self.listen_port = listen_port
        is_gh = os.environ.get("GITHUB_ACTIONS") == "true"
        self.lan_ip, self.if_index = detect_physical_lan_interface() if (os.name == "nt" and not is_gh) else (None, None)
        self.gateway_node = (gateway_node or os.environ.get("NET_NODE") or "").strip()
        self.gateway_nodes = parse_gateway_nodes(self.gateway_node)
        self._dns_cache: Dict[str, str] = {}
        self._server: Optional[asyncio.AbstractServer] = None
        self._logged_ok = False
        self._logged_fallback_hosts: set[str] = set()
        self._direct_only_hosts: set[str] = set()
        self._direct_only_suffixes: set[str] = set()
        self._connect_sem: Optional[asyncio.Semaphore] = None

    def mark_direct_suffixes(self, *suffixes: str) -> None:
        """Переводит семейство доменов (например avito.ru, avito.st) на единый прямой канал без прокси."""
        for s in suffixes:
            clean = (s or "").strip().lower().lstrip(".")
            if clean:
                self._direct_only_suffixes.add(clean)

    @property
    def proxy_url(self) -> str:
        return f"http://{self.listen_host}:{self.listen_port}"

    def _bind_socket_to_hardware(self, sock: socket.socket) -> None:
        if self.if_index:
            # В Windows IP_UNICAST_IF принимает network byte order (big-endian) для IPv4
            sock.setsockopt(socket.IPPROTO_IP, IP_UNICAST_IF, struct.pack("!I", self.if_index))
        if self.lan_ip:
            sock.bind((self.lan_ip, 0))

    def resolve_domain_direct(self, host: str) -> str:
        try:
            socket.inet_aton(host)
            return host
        except OSError:
            pass

        if host in self._dns_cache:
            return self._dns_cache[host]

        for dns_srv in ("77.88.8.8", "77.88.8.1", "8.8.8.8", "1.1.1.1"):
            sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            try:
                self._bind_socket_to_hardware(sock)
                sock.settimeout(2.5)
                tid = b"\xca\xfe"
                flags = b"\x01\x00"
                counts = struct.pack("!HHHH", 1, 0, 0, 0)
                qname = b"".join(bytes([len(p)]) + p.encode("idna") for p in host.strip(".").split(".")) + b"\x00"
                qtype_class = struct.pack("!HH", 1, 1)
                packet = tid + flags + counts + qname + qtype_class
                sock.sendto(packet, (dns_srv, 53))
                data, _ = sock.recvfrom(1024)
                idx = len(packet)
                ancount = struct.unpack("!H", data[6:8])[0]
                for _ in range(ancount):
                    while idx < len(data):
                        if data[idx] & 0xC0 == 0xC0:
                            idx += 2
                            break
                        elif data[idx] == 0:
                            idx += 1
                            break
                        else:
                            idx += 1 + data[idx]
                    rtype, rclass, ttl, rdlen = struct.unpack("!HHIH", data[idx : idx + 10])
                    idx += 10
                    if rtype == 1 and rdlen == 4:
                        ip_str = socket.inet_ntoa(data[idx : idx + 4])
                        self._dns_cache[host] = ip_str
                        return ip_str
                    idx += rdlen
            except Exception:
                continue
            finally:
                sock.close()

        ip = socket.gethostbyname(host)
        self._dns_cache[host] = ip
        return ip

    async def _connect_proxy_tcp(self, s_host: str, s_port: int) -> Tuple[asyncio.StreamReader, asyncio.StreamWriter]:
        loop = asyncio.get_running_loop()
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            if os.name == "nt" and os.environ.get("GITHUB_ACTIONS") != "true":
                self._bind_socket_to_hardware(sock)
            sock.setblocking(False)
            await asyncio.wait_for(loop.sock_connect(sock, (s_host, s_port)), timeout=5.5)
            return await asyncio.open_connection(sock=sock)
        except Exception:
            sock.close()
            raise

    async def _open_socks5_channel(
        self,
        node: dict,
        target_host: str,
        target_port: int,
        use_domain: bool = True,
    ) -> Tuple[asyncio.StreamReader, asyncio.StreamWriter]:
        s_host, s_port = node["host"], node["port"]
        s_user, s_pass = node.get("user", ""), node.get("password", "")
        reader, writer = await self._connect_proxy_tcp(s_host, s_port)
        try:
            writer.write(b"\x05\x02\x00\x02" if s_user else b"\x05\x01\x00")
            await writer.drain()
            greeting = await asyncio.wait_for(reader.readexactly(2), timeout=6.0)
            if greeting[0] != 5:
                raise ValueError("Not a SOCKS5 protocol")
            if greeting[1] == 2:
                u_b = s_user.encode("utf-8")
                p_b = s_pass.encode("utf-8")
                writer.write(b"\x01" + bytes([len(u_b)]) + u_b + bytes([len(p_b)]) + p_b)
                await writer.drain()
                auth_resp = await asyncio.wait_for(reader.readexactly(2), timeout=6.0)
                if auth_resp[1] != 0:
                    raise PermissionError("Node handshake auth failed")
            elif greeting[1] != 0:
                raise RuntimeError(f"Unsupported SOCKS5 auth method: {greeting[1]}")

            is_ipv4 = False
            try:
                socket.inet_aton(target_host)
                is_ipv4 = True
            except OSError:
                pass

            if use_domain and not is_ipv4 and len(target_host.encode("idna")) <= 255:
                host_b = target_host.encode("idna")
                req = b"\x05\x01\x00\x03" + bytes([len(host_b)]) + host_b + struct.pack("!H", target_port)
            else:
                loop = asyncio.get_running_loop()
                real_ip = target_host if is_ipv4 else await loop.run_in_executor(None, self.resolve_domain_direct, target_host)
                req = b"\x05\x01\x00\x01" + socket.inet_aton(real_ip) + struct.pack("!H", target_port)

            writer.write(req)
            await writer.drain()

            hdr = await asyncio.wait_for(reader.readexactly(4), timeout=6.5)
            if hdr[1] != 0:
                raise RuntimeError(f"Channel error code: {hdr[1]}")
            atyp = hdr[3]
            if atyp == 1:
                await asyncio.wait_for(reader.readexactly(6), timeout=4.0)
            elif atyp == 3:
                dlen_b = await asyncio.wait_for(reader.readexactly(1), timeout=4.0)
                await asyncio.wait_for(reader.readexactly(dlen_b[0] + 2), timeout=4.0)
            elif atyp == 4:
                await asyncio.wait_for(reader.readexactly(18), timeout=4.0)
            return reader, writer
        except Exception:
            try:
                writer.close()
            except Exception:
                pass
            raise

    async def _open_http_proxy_channel(
        self,
        node: dict,
        target_host: str,
        target_port: int,
    ) -> Tuple[asyncio.StreamReader, asyncio.StreamWriter]:
        s_host, s_port = node["host"], node["port"]
        s_user, s_pass = node.get("user", ""), node.get("password", "")
        reader, writer = await self._connect_proxy_tcp(s_host, s_port)
        try:
            headers = [
                f"CONNECT {target_host}:{target_port} HTTP/1.1",
                f"Host: {target_host}:{target_port}",
                "Proxy-Connection: keep-alive",
            ]
            if s_user or s_pass:
                cred = base64.b64encode(f"{s_user}:{s_pass}".encode("utf-8")).decode("ascii")
                headers.append(f"Proxy-Authorization: Basic {cred}")
            req = ("\r\n".join(headers) + "\r\n\r\n").encode("latin-1")
            writer.write(req)
            await writer.drain()
            raw_hdr = await asyncio.wait_for(reader.readuntil(b"\r\n\r\n"), timeout=8.0)
            status_line = raw_hdr.split(b"\r\n", 1)[0].decode("latin-1", errors="ignore")
            if " 200" not in status_line:
                raise RuntimeError(f"HTTP proxy rejected CONNECT: {status_line[:60]}")
            return reader, writer
        except Exception:
            try:
                writer.close()
            except Exception:
                pass
            raise

    async def _open_node_channel(self, target_host: str, target_port: int, node: Optional[dict] = None) -> Tuple[asyncio.StreamReader, asyncio.StreamWriter]:
        if node is None:
            if not self.gateway_nodes:
                raise ValueError("Invalid node format")
            node = self.gateway_nodes[0]
        scheme = node.get("scheme", "auto")
        if scheme == "http":
            return await self._open_http_proxy_channel(node, target_host, target_port)
        try:
            return await self._open_socks5_channel(node, target_host, target_port, use_domain=True)
        except ValueError:
            if scheme == "auto":
                return await self._open_http_proxy_channel(node, target_host, target_port)
            raise
        except RuntimeError as e:
            msg = str(e)
            if scheme in ("auto", "socks5") and any(c in msg for c in ("code: 1", "code: 2", "code: 4", "code: 8")):
                return await self._open_socks5_channel(node, target_host, target_port, use_domain=False)
            raise

    async def _open_target_connection(self, host: str, port: int) -> Tuple[asyncio.StreamReader, asyncio.StreamWriter]:
        if self._connect_sem is None:
            self._connect_sem = asyncio.Semaphore(32)
        last_err = ""
        acl_refused = False
        is_cloud = os.name != "nt" or os.environ.get("GITHUB_ACTIONS") == "true"
        is_avito_host = _host_matches_suffix(host, ("avito.ru", "avito.st"))
        is_direct_only = (
            host in self._direct_only_hosts
            or (bool(self._direct_only_suffixes) and _host_matches_suffix(host, tuple(self._direct_only_suffixes)))
        )
        if not is_direct_only and self.gateway_nodes:
            max_rounds = 2 if (is_cloud and _host_matches_suffix(host, _PROTECTED_RU_SHOP_SUFFIXES)) else 1
            for round_idx in range(max_rounds):
                async with self._connect_sem:
                    for idx, node in enumerate(self.gateway_nodes):
                        try:
                            conn = await self._open_node_channel(host, port, node=node)
                            if not self._logged_ok:
                                self._logged_ok = True
                                print(f"[Шлюз РФ] Прокси-узел #{idx + 1} подключён ({node.get('scheme', 'auto').upper()}).", flush=True)
                            return conn
                        except Exception as e:
                            last_err = str(e) or type(e).__name__
                if "code: 2" in last_err:
                    acl_refused = True
                    break
                if round_idx + 1 < max_rounds:
                    await asyncio.sleep(0.25)
            if acl_refused:
                if len(self._direct_only_hosts) < 500:
                    self._direct_only_hosts.add(host)
                if is_avito_host:
                    self.mark_direct_suffixes("avito.ru", "avito.st")
            elif is_avito_host:
                # Если узел не пропускает один из хостов Авито, переводим весь кластер Авито на единый прямой маршрут
                # без рассинхрона IP между HTML (www.avito.ru) и скриптами/API (avito.st, m.avito.ru).
                self.mark_direct_suffixes("avito.ru", "avito.st")
            if is_cloud and not acl_refused and host not in self._direct_only_hosts and _host_matches_suffix(host, _PROTECTED_RU_SHOP_SUFFIXES):
                if host not in self._logged_fallback_hosts:
                    self._logged_fallback_hosts.add(host)
                    print(f"[Шлюз РФ] Сбой прокси-узла для {host}: {last_err[:80]}", flush=True)
                raise RuntimeError(f"RU gateway unreachable for {host}: {last_err}")
        if self.gateway_nodes and host not in self._logged_fallback_hosts:
            self._logged_fallback_hosts.add(host)
            print(f"[Шлюз РФ] Прямой резервный канал для {host} ({last_err[:80] or 'direct'}).", flush=True)
        loop = asyncio.get_running_loop()
        if os.name == "nt" and (self.lan_ip or self.if_index):
            real_ip = await loop.run_in_executor(None, self.resolve_domain_direct, host)
            return await self._open_hardware_connection(real_ip, port)
        return await asyncio.wait_for(asyncio.open_connection(host, port), timeout=5.0)

    async def _open_hardware_connection(self, real_ip: str, port: int) -> Tuple[asyncio.StreamReader, asyncio.StreamWriter]:
        loop = asyncio.get_running_loop()
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            self._bind_socket_to_hardware(sock)
            sock.setblocking(False)
            await asyncio.wait_for(loop.sock_connect(sock, (real_ip, port)), timeout=5.0)
            return await asyncio.open_connection(sock=sock)
        except Exception:
            sock.close()
            raise

    async def _pipe(self, reader: asyncio.StreamReader, writer: asyncio.StreamWriter) -> None:
        try:
            while not reader.at_eof():
                chunk = await asyncio.wait_for(reader.read(65536), timeout=60.0)
                if not chunk:
                    break
                writer.write(chunk)
                await asyncio.wait_for(writer.drain(), timeout=15.0)
        except Exception:
            pass
        finally:
            try:
                if not writer.is_closing():
                    writer.close()
            except Exception:
                pass

    async def _handle_client(self, client_reader: asyncio.StreamReader, client_writer: asyncio.StreamWriter) -> None:
        try:
            header_data = await asyncio.wait_for(client_reader.readuntil(b"\r\n\r\n"), timeout=10.0)
            lines = header_data.decode("latin-1", errors="ignore").split("\r\n")
            if not lines or not lines[0]:
                client_writer.close()
                return

            parts = lines[0].split(" ")
            if len(parts) < 2:
                client_writer.close()
                return

            method, target = parts[0].upper(), parts[1]
            loop = asyncio.get_running_loop()

            if method == "CONNECT":
                if ":" in target:
                    host, port_s = target.rsplit(":", 1)
                    port = int(port_s)
                else:
                    host, port = target, 443

                if _host_matches_suffix(host, _AD_TRACKER_SUFFIXES):
                    client_writer.write(b"HTTP/1.1 403 Forbidden\r\nConnection: close\r\n\r\n")
                    await client_writer.drain()
                    return

                remote_reader, remote_writer = await self._open_target_connection(host, port)
                client_writer.write(b"HTTP/1.1 200 Connection Established\r\n\r\n")
                await client_writer.drain()

                await asyncio.gather(
                    self._pipe(client_reader, remote_writer),
                    self._pipe(remote_reader, client_writer),
                )
            else:
                from urllib.parse import urlsplit

                parsed = urlsplit(target)
                host = parsed.hostname or ""
                port = parsed.port or 80
                path = (parsed.path or "/") + (f"?{parsed.query}" if parsed.query else "")

                if _host_matches_suffix(host, _AD_TRACKER_SUFFIXES):
                    client_writer.write(b"HTTP/1.1 403 Forbidden\r\nConnection: close\r\n\r\n")
                    await client_writer.drain()
                    return

                remote_reader, remote_writer = await self._open_target_connection(host, port)

                rebuilt = f"{method} {path} HTTP/1.1\r\n" + "\r\n".join(lines[1:])
                remote_writer.write(rebuilt.encode("latin-1"))
                await remote_writer.drain()

                await asyncio.gather(
                    self._pipe(client_reader, remote_writer),
                    self._pipe(remote_reader, client_writer),
                )
        except Exception:
            pass
        finally:
            try:
                client_writer.close()
            except Exception:
                pass

    async def start(self) -> None:
        if self._server is not None:
            return
        self._server = await asyncio.start_server(self._handle_client, self.listen_host, self.listen_port)

    async def stop(self) -> None:
        if self._server is not None:
            self._server.close()
            try:
                await asyncio.wait_for(self._server.wait_closed(), timeout=1.0)
            except Exception:
                pass
            self._server = None
