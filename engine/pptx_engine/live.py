# -*- coding: utf-8 -*-
"""实时预览 HTTP 侧车 —— 「让我能实时看到 AI 在写 PPT」的服务端。

为什么需要它：MCP 是请求/响应模型，工具调用之间模型不会主动汇报进度，
DSH 客户端也拿不到中间态。所以在引擎进程里再起一个**只读** HTTP 面，
Node 宿主把它代理成同源路由 `/ppt/*`，客户端面板每 1.2s 轮询一次，
就能形成「PowerPoint 窗口 + DSH 内缩略图 + 操作字幕」三路实时反馈。

线程安全约束（重要）：
    COM 对象只能在 `CoInitialize` 过的那个线程里用。本 HTTP 服务跑在
    独立 daemon 线程，**绝不触碰任何 COM 对象** —— 它只读两样东西：
      1. `<workdir>/.live/state.json`（Deck 每次变更后原子写入）
      2. `<workdir>/.live/slide.png` （Deck 每次变更后导出的当前页快照）
    这样即使 HTTP 线程和 COM 线程并发，也不会有跨线程 COM 崩溃。
"""
from __future__ import annotations

import base64
import json
import os
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Callable, Optional

_STATE_CACHE: dict = {"mtime": 0.0, "path": "", "data": {}}


def read_state(state_path) -> dict:
    """读 state.json（带 mtime 缓存，避免每次轮询都解析）。"""
    p = str(state_path)
    try:
        mt = os.path.getmtime(p)
    except OSError:
        return {}
    if _STATE_CACHE["path"] == p and _STATE_CACHE["mtime"] == mt:
        return _STATE_CACHE["data"]
    try:
        with open(p, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:  # noqa: BLE001
        return {}
    _STATE_CACHE.update({"path": p, "mtime": mt, "data": data})
    return data


def make_handler(get_live_dir: Callable[[], Optional[Path]], get_info: Callable[[], dict]):
    """构造 RequestHandler 类。

    Args:
        get_live_dir: 返回当前 deck 的 .live 目录（无 deck 时返回 None）
        get_info:     返回引擎级信息（后端/版本/主题列表等），不依赖 deck
    """

    class Handler(BaseHTTPRequestHandler):
        protocol_version = "HTTP/1.1"
        server_version = "dsh-ppt-live/1.0"

        # ---- 基础设施 ----
        def log_message(self, fmt, *args):  # noqa: A003
            pass  # 静默 —— 轮询频率高，刷日志会污染 stderr

        def _send(self, code: int, body: bytes, ctype: str, extra: Optional[dict] = None):
            self.send_response(code)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Access-Control-Allow-Origin", "*")
            self.send_header("Cache-Control", "no-store, max-age=0")
            for k, v in (extra or {}).items():
                self.send_header(k, v)
            self.end_headers()
            try:
                self.wfile.write(body)
            except Exception:  # noqa: BLE001
                pass

        def _json(self, obj, code: int = 200):
            self._send(code, json.dumps(obj, ensure_ascii=False).encode("utf-8"),
                       "application/json; charset=utf-8")

        def _live_dir(self) -> Optional[Path]:
            try:
                d = get_live_dir()
                return Path(d) if d else None
            except Exception:  # noqa: BLE001
                return None

        # ---- 路由 ----
        def do_GET(self):  # noqa: N802
            path = self.path.split("?", 1)[0].rstrip("/") or "/"
            qs = self.path.split("?", 1)[1] if "?" in self.path else ""

            if path in ("/", "/health"):
                info = {}
                try:
                    info = get_info() or {}
                except Exception:  # noqa: BLE001
                    pass
                return self._json({"ok": True, "service": "dsh-ppt-live", **info})

            if path == "/state":
                d = self._live_dir()
                if not d:
                    return self._json({"deck": None, "ops": [], "snapshot": None})
                return self._json(read_state(d / "state.json") or {"deck": None})

            if path == "/ops":
                d = self._live_dir()
                since = 0
                for kv in qs.split("&"):
                    if kv.startswith("since="):
                        try:
                            since = int(kv[6:])
                        except ValueError:
                            since = 0
                st = read_state(d / "state.json") if d else {}
                ops = [o for o in (st.get("ops") or []) if int(o.get("seq", 0)) > since]
                return self._json({"since": since, "ops": ops,
                                   "deck": st.get("deck"),
                                   "maxSeq": max([int(o.get("seq", 0)) for o in (st.get("ops") or [])] or [0])})

            if path in ("/snapshot.png", "/snapshot"):
                d = self._live_dir()
                if not d:
                    return self._send(404, b"no deck", "text/plain; charset=utf-8")
                png = d / "slide.png"
                if not png.exists():
                    return self._send(404, b"no snapshot", "text/plain; charset=utf-8")
                try:
                    data = png.read_bytes()
                except Exception:  # noqa: BLE001
                    return self._send(500, b"read error", "text/plain; charset=utf-8")
                return self._send(200, data, "image/png",
                                  {"Last-Modified": str(int(png.stat().st_mtime))})

            if path == "/live":
                # 一次拿全：状态 + 操作流水 + base64 快照（省一次请求，面板首屏用）
                d = self._live_dir()
                st = read_state(d / "state.json") if d else {}
                img = None
                if d:
                    png = d / "slide.png"
                    if png.exists():
                        try:
                            img = "data:image/png;base64," + base64.b64encode(png.read_bytes()).decode("ascii")
                        except Exception:  # noqa: BLE001
                            img = None
                return self._json({"state": st, "snapshot": img})

            return self._send(404, b"not found", "text/plain; charset=utf-8")

        def do_POST(self):  # noqa: N802
            # 侧车是只读的；POST 一律 405，避免被误当成控制面
            self._send(405, b"read-only", "text/plain; charset=utf-8")

    return Handler


class _NoReuseServer(ThreadingHTTPServer):
    """禁用 SO_REUSEADDR。

    Windows 上 allow_reuse_address=1 可能让第二个进程「绑定成功」到已被占用的端口，
    形成双监听者：请求被随机路由到旧进程、读到旧目录 → /snapshot.png 404
    （2026-10-28 实测，遗留 python 进程占着 8790 所致）。禁用后绑定会干净地
    OSError，走端口顺延逻辑。
    """
    allow_reuse_address = False
    daemon_threads = True


class LiveServer:
    """实时预览服务。start() 起 daemon 线程，stop() 关闭。"""

    def __init__(self, get_live_dir: Callable[[], Optional[Path]],
                 get_info: Callable[[], dict], port: int = 8790,
                 host: str = "127.0.0.1"):
        self.get_live_dir = get_live_dir
        self.get_info = get_info
        self.port = int(port)
        self.host = host
        self._httpd = None
        self._thread: Optional[threading.Thread] = None

    @property
    def running(self) -> bool:
        return self._thread is not None and self._thread.is_alive()

    @property
    def url(self) -> str:
        return f"http://{self.host}:{self.port}"

    def start(self) -> dict:
        """启动；端口被占用时自动向后找 20 个（返回实际端口）。"""
        if self.running:
            return {"ok": True, "port": self.port, "url": self.url, "reused": True}
        last_err = None
        for candidate in range(self.port, self.port + 20):
            try:
                handler = make_handler(self.get_live_dir, self.get_info)
                self._httpd = _NoReuseServer((self.host, candidate), handler)
                self.port = candidate
                self._thread = threading.Thread(
                    target=self._httpd.serve_forever,
                    name="dsh-ppt-live", daemon=True)
                self._thread.start()
                return {"ok": True, "port": candidate, "url": self.url, "reused": False}
            except OSError as e:
                last_err = e
                continue
        return {"ok": False, "port": None, "url": "",
                "error": f"端口 {self.port}~{self.port + 19} 均不可用：{last_err}"}

    def stop(self) -> None:
        try:
            if self._httpd is not None:
                self._httpd.shutdown()
                self._httpd.server_close()
        except Exception:  # noqa: BLE001
            pass
        self._httpd = None
        self._thread = None
