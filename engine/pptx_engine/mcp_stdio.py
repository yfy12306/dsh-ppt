# -*- coding: utf-8 -*-
"""零依赖 MCP stdio 服务端 —— 手写 JSON-RPC 2.0，不引入 `mcp` 包。

为什么不用官方 `mcp` SDK：本机唯一带 `mcp` 的解释器是 hermes-agent 的 venv
（`include-system-site-packages = false`，且被 hermes 自己管着，随时可能重建）。
把协议层手写掉，引擎就只依赖 pywin32 + Pillow（+ 可选 PyMuPDF/matplotlib），
任何 Python 3.9+ 都能跑，插件寿命不受第三方 venv 摆布。

协议要点（MCP stdio transport）：
  * 换行分隔的 JSON-RPC 2.0，一行一条消息，UTF-8，消息内不得含裸换行
  * 服务端必须实现：initialize / notifications/initialized / ping /
    tools/list / tools/call
  * **stdout 被协议独占** —— 本模块启动时把 `sys.stdout` 重定向到 stderr，
    只保留一个私有句柄写协议帧。这样任何库里的意外 `print()` 都不会
    撕坏协议流（这是手写 MCP 服务端最常踩的坑）。
"""
from __future__ import annotations

import json
import os
import sys
import traceback
from typing import Any, Callable, Optional

PROTOCOL_VERSIONS = ("2024-11-05", "2025-03-26", "2025-06-18", "2025-11-25")
DEFAULT_PROTOCOL_VERSION = "2025-06-18"

SERVER_NAME = "dsh-ppt"
SERVER_VERSION = "1.0.0"


def _err(msg: str) -> None:
    """日志一律 stderr。"""
    try:
        sys.stderr.write(f"[mcp] {msg}\n")
        sys.stderr.flush()
    except Exception:  # noqa: BLE001
        pass


class Tool:
    """一个 MCP 工具的元数据 + 处理函数。"""

    def __init__(self, name: str, description: str, schema: dict,
                 handler: Callable[[dict], Any], title: str = ""):
        self.name = name
        self.description = description
        self.schema = schema
        self.handler = handler
        self.title = title or name

    def descriptor(self) -> dict:
        d = {"name": self.name, "description": self.description,
             "inputSchema": self.schema}
        # annotations 是 2025-03-26 起的可选字段，老客户端会忽略
        d["annotations"] = {"title": self.title, "readOnlyHint": False,
                            "destructiveHint": False, "openWorldHint": False}
        return d


class McpStdioServer:
    """极简 MCP 服务端。

    用法::

        srv = McpStdioServer()

        @srv.tool("echo", "回显", {"type":"object","properties":{"text":{"type":"string"}}})
        def _echo(args):
            return {"text": args.get("text", "")}

        srv.serve()
    """

    def __init__(self, name: str = SERVER_NAME, version: str = SERVER_VERSION,
                 instructions: str = ""):
        self.name = name
        self.version = version
        self.instructions = instructions
        self.tools: dict[str, Tool] = {}
        self._out = None
        self._protocol_version = DEFAULT_PROTOCOL_VERSION
        self._initialized = False

    # ------------------------------------------------------------ 注册
    def tool(self, name: str, description: str, schema: Optional[dict] = None,
             title: str = ""):
        """装饰器：注册一个工具。handler 收 dict，返回可 JSON 序列化的对象或字符串。"""
        schema = schema or {"type": "object", "properties": {}, "additionalProperties": True}

        def deco(fn):
            self.tools[name] = Tool(name, description, schema, fn, title=title)
            return fn
        return deco

    def register(self, name: str, description: str, schema: dict,
                 handler: Callable[[dict], Any], title: str = "") -> None:
        self.tools[name] = Tool(name, description, schema, handler, title=title)

    # ------------------------------------------------------------ I/O
    def _write(self, obj: dict) -> None:
        if self._out is None:
            return
        line = json.dumps(obj, ensure_ascii=False, separators=(",", ":"))
        try:
            self._out.write(line + "\n")
            self._out.flush()
        except Exception as e:  # noqa: BLE001
            _err(f"写协议帧失败：{e}")

    def _result(self, rid, result: Any) -> None:
        self._write({"jsonrpc": "2.0", "id": rid, "result": result})

    def _error(self, rid, code: int, message: str, data: Any = None) -> None:
        e = {"code": code, "message": message}
        if data is not None:
            e["data"] = data
        self._write({"jsonrpc": "2.0", "id": rid, "error": e})

    # ------------------------------------------------------------ 方法分发
    def _handle_initialize(self, rid, params: dict) -> None:
        want = str(params.get("protocolVersion") or "")
        self._protocol_version = want if want in PROTOCOL_VERSIONS else DEFAULT_PROTOCOL_VERSION
        caps = {"tools": {"listChanged": False}}
        res = {
            "protocolVersion": self._protocol_version,
            "capabilities": caps,
            "serverInfo": {"name": self.name, "version": self.version},
        }
        if self.instructions:
            res["instructions"] = self.instructions
        self._result(rid, res)
        _err(f"initialize 完成，协议版本 {self._protocol_version}，工具 {len(self.tools)} 个")

    def _handle_tools_list(self, rid, _params: dict) -> None:
        self._result(rid, {"tools": [t.descriptor() for t in self.tools.values()]})

    def _handle_tools_call(self, rid, params: dict) -> None:
        name = str(params.get("name") or "")
        args = params.get("arguments")
        if not isinstance(args, dict):
            args = {}
        meta = params.get("_meta") or {}

        t = self.tools.get(name)
        if t is None:
            # 工具不存在属于协议级错误（-32602 invalid params）
            self._error(rid, -32602, f"Unknown tool: {name}",
                        {"available": sorted(self.tools.keys())})
            return
        try:
            out = t.handler(args)
        except Exception as e:  # noqa: BLE001
            tb = traceback.format_exc(limit=6)
            _err(f"工具 {name} 异常：{e}\n{tb}")
            payload = {"ok": False, "tool": name,
                       "error": f"{type(e).__name__}: {e}", "trace": tb.splitlines()[-4:]}
            self._result(rid, {"content": [{"type": "text",
                                            "text": json.dumps(payload, ensure_ascii=False, indent=2)}],
                               "isError": True})
            return

        text = out if isinstance(out, str) else json.dumps(out, ensure_ascii=False, indent=2, default=str)
        res = {"content": [{"type": "text", "text": text}], "isError": False}
        if meta:
            res["_meta"] = {"echo": meta}
        self._result(rid, res)

    def _dispatch(self, msg: dict) -> None:
        method = str(msg.get("method") or "")
        rid = msg.get("id")
        params = msg.get("params")
        if not isinstance(params, dict):
            params = {}
        is_request = rid is not None

        if method == "initialize":
            return self._handle_initialize(rid, params)
        if method == "notifications/initialized" or method == "initialized":
            self._initialized = True
            return  # 通知，不回帧
        if method == "ping":
            if is_request:
                self._result(rid, {})
            return
        if method == "tools/list":
            return self._handle_tools_list(rid, params)
        if method == "tools/call":
            return self._handle_tools_call(rid, params)
        if method in ("notifications/cancelled", "notifications/progress",
                      "notifications/roots/list_changed", "logging/setLevel",
                      "completion/complete", "resources/list", "prompts/list"):
            if is_request:
                # 明确不支持的能力：返回空集合而不是报错，避免客户端重试风暴
                if method == "resources/list":
                    self._result(rid, {"resources": []})
                elif method == "prompts/list":
                    self._result(rid, {"prompts": []})
                else:
                    self._result(rid, {})
            return

        if is_request:
            self._error(rid, -32601, f"Method not found: {method}")
        # 未知通知：静默忽略

    # ------------------------------------------------------------ 主循环
    def serve(self) -> None:
        """阻塞式主循环：读 stdin 一行一条 JSON-RPC，写 stdout。"""
        # 关键：把 sys.stdout 换成 stderr，只留私有句柄写协议。
        # 这样 matplotlib / PIL / 任何三方库的意外 print 都不会污染协议流。
        try:
            self._out = sys.stdout
            sys.stdout = sys.stderr
            # Windows 控制台默认 GBK，强制 UTF-8 避免中文日志炸 UnicodeEncodeError
            for stream in (sys.stdout, sys.stderr):
                try:
                    stream.reconfigure(encoding="utf-8", errors="replace")
                except Exception:  # noqa: BLE001
                    pass
            try:
                self._out.reconfigure(encoding="utf-8", errors="replace")
            except Exception:  # noqa: BLE001
                pass
        except Exception:  # noqa: BLE001
            pass

        stdin = sys.stdin
        _err(f"{self.name} v{self.version} 启动，pid={os.getpid()}，工具 {len(self.tools)} 个")
        while True:
            try:
                line = stdin.readline()
            except Exception as e:  # noqa: BLE001
                _err(f"读 stdin 失败：{e}")
                break
            if not line:
                _err("stdin 关闭（宿主断开），退出")
                break
            line = line.strip()
            if not line:
                continue
            try:
                msg = json.loads(line)
            except json.JSONDecodeError as e:
                _err(f"JSON 解析失败（已跳过该行）：{e}")
                # 解析错误无法关联 id，按规范回 id=null 的错误帧
                self._error(None, -32700, f"Parse error: {e}")
                continue
            if isinstance(msg, list):
                for m in msg:
                    if isinstance(m, dict):
                        self._dispatch(m)
            elif isinstance(msg, dict):
                try:
                    self._dispatch(msg)
                except Exception as e:  # noqa: BLE001
                    _err(f"分发异常：{e}\n{traceback.format_exc(limit=4)}")
                    if msg.get("id") is not None:
                        self._error(msg["id"], -32603, f"Internal error: {e}")
        try:
            if self._out is not None:
                self._out.flush()
        except Exception:  # noqa: BLE001
            pass
