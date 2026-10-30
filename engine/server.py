# -*- coding: utf-8 -*-
"""dsh-ppt 引擎入口 —— MCP 工具定义与全局会话状态。

工具面分三组：
  ppt_*     演示文稿制作（核心）
  zotero_*  文献检索与参考文献页（可选，Zotero 未装时优雅降级）
  实时预览   由 live.py 的 HTTP 侧车承担，不占 MCP 工具位

会话模型：进程内单例 `ENGINE` 持有当前 Deck。COM 对象不能跨进程，
所以「新建 → 逐页绘制 → 保存」必须在同一个 Python 进程里完成，
这也正是把引擎做成常驻 MCP server（而非一次性 CLI）的原因。
"""
from __future__ import annotations

import atexit
import json
import os
import sys
import time
from pathlib import Path
from typing import Optional

# 自举：本文件位于 engine/ 下、包 pptx_engine/ 之外，既可能被
# `python engine/server.py` 当脚本跑，也可能被 `python -m pptx_engine.server` 当模块跑。
# 脚本方式下没有包上下文，相对导入会炸，所以先把 engine/ 塞进 sys.path 再用绝对导入。
if __package__ in (None, ""):
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from pptx_engine import assets as A          # noqa: E402
from pptx_engine import charts as C          # noqa: E402
from pptx_engine import elements as el       # noqa: E402
from pptx_engine import office as O          # noqa: E402
from pptx_engine import quality as Q         # noqa: E402
from pptx_engine import renderer as R        # noqa: E402
from pptx_engine import themes as T          # noqa: E402
from pptx_engine import verify as V          # noqa: E402
from pptx_engine import zotero as Z          # noqa: E402
from pptx_engine.deck import Deck            # noqa: E402
from pptx_engine.live import LiveServer      # noqa: E402
from pptx_engine.mcp_stdio import McpStdioServer  # noqa: E402

# ---------------------------------------------------------------- 全局会话
DEFAULT_WORKDIR = Path(os.environ.get("DSH_PPT_WORKDIR") or
                       (Path.home() / "Desktop" / "dsh-ppt-out"))


class Engine:
    """进程内单例：当前 deck + 实时预览服务 + 引擎级配置。"""

    def __init__(self):
        self.deck: Optional[Deck] = None
        self.workdir: Path = DEFAULT_WORKDIR
        self.assets_dir: Path = self.workdir / ".ppt-assets"
        self.theme: dict = T.get_theme("default")
        self.visible: bool = os.environ.get("DSH_PPT_VISIBLE", "1") not in ("0", "false", "False")
        self.prefer_backend: str = os.environ.get("DSH_PPT_BACKEND", "auto")
        self.live_port: int = int(os.environ.get("DSH_PPT_LIVE_PORT", "8790"))
        self.snapshot_scale: int = int(os.environ.get("DSH_PPT_SNAPSHOT_SCALE", "960"))
        self.auto_pdf: bool = os.environ.get("DSH_PPT_AUTO_PDF", "1") not in ("0", "false", "False")
        self.quit_on_exit: bool = os.environ.get("DSH_PPT_QUIT_ON_EXIT", "0") not in ("0", "false", "False")
        self.live: Optional[LiveServer] = None
        self._started = time.time()
        self._errors: list[dict] = []

    # ------------------------------------------------------------ 生命周期
    def ensure_workdir(self) -> Path:
        self.workdir.mkdir(parents=True, exist_ok=True)
        self.assets_dir.mkdir(parents=True, exist_ok=True)
        el.set_assets_dir(self.assets_dir)
        return self.workdir

    def start_live(self) -> dict:
        if self.live and self.live.running:
            return {"ok": True, "port": self.live.port, "url": self.live.url, "reused": True}
        self.live = LiveServer(self._live_dir, self.live_info, port=self.live_port)
        res = self.live.start()
        if res.get("ok"):
            self.live_port = res["port"]
        return res

    def _live_dir(self) -> Optional[Path]:
        return self.deck.live_dir if self.deck else None

    def live_info(self) -> dict:
        return {
            "engine": "dsh-ppt",
            "version": "1.0.0",
            "hasDeck": self.deck is not None,
            "workdir": str(self.workdir),
            "theme": self.theme.get("id"),
            "visible": self.visible,
            "backends": [b["id"] for b in O.BACKENDS],
        }

    def require_deck(self) -> Deck:
        if self.deck is None:
            raise RuntimeError("当前没有打开的演示文稿。请先调用 ppt_new 新建，或 ppt_open 打开已有文件。")
        return self.deck

    def new_deck(self, theme: str = "default", width: int = 960, height: int = 540,
                 bg: Optional[str] = None, visible: Optional[bool] = None,
                 workdir: Optional[str] = None, font: str = "微软雅黑",
                 ensure_assets: bool = True, footer: Optional[bool] = None,
                 footer_text: Optional[str] = None, transition: Optional[str] = None,
                 entrance: Optional[str] = None) -> Deck:
        """新建 deck；若已有未保存的 deck，先自动落一份恢复副本再关。"""
        if self.deck is not None:
            self._auto_recover()
        self.workdir = Path(workdir).expanduser().resolve() if workdir else self.workdir
        self.ensure_workdir()
        self.theme = T.get_theme(theme)
        if transition:
            self.theme["transition"] = str(transition)
        if entrance:
            self.theme["entrance"] = str(entrance)
        if footer is not None:
            self.theme["footer"] = bool(footer)
        if footer_text:
            self.theme["footerText"] = str(footer_text)
        if visible is not None:
            self.visible = bool(visible)

        if ensure_assets:
            try:
                A.ensure_default_assets(self.assets_dir)
                A.ensure_theme_assets(self.theme, self.assets_dir)
            except Exception as e:  # noqa: BLE001
                self._note_error("assets", e)
        el.set_assets_dir(self.assets_dir)

        # 未显式给底图时，用主题自带的纹理背景（纯白底是调研点名的反模式）
        if bg is None:
            cand = self.assets_dir / f"bg_{self.theme.get('id', 'business')}.png"
            if cand.exists():
                bg = str(cand)

        th = dict(self.theme)
        th["assetsDir"] = str(self.assets_dir)
        self.deck = Deck(width=width, height=height, workdir=str(self.workdir),
                         visible=self.visible, bg=bg, font=font,
                         prefer_backend=self.prefer_backend, live=True,
                         snapshot_scale=self.snapshot_scale, theme=th)
        self.start_live()
        return self.deck

    def _auto_recover(self) -> Optional[str]:
        """关闭旧 deck 前，把未保存内容落到 recovery/ 下，避免用户东西丢失。"""
        if self.deck is None:
            return None
        try:
            if self.deck.slide_count > 0:
                rec_dir = self.workdir / "recovery"
                rec_dir.mkdir(parents=True, exist_ok=True)
                p = rec_dir / f"autosave-{time.strftime('%Y%m%d-%H%M%S')}.pptx"
                self.deck.save(str(p), pdf=False)
                self.deck.close(quit_app=False)
                return str(p)
        except Exception as e:  # noqa: BLE001
            self._note_error("auto_recover", e)
        try:
            self.deck.close(quit_app=False)
        except Exception:  # noqa: BLE001
            pass
        self.deck = None
        return None

    def close_deck(self, quit_app: Optional[bool] = None) -> None:
        if self.deck is None:
            return
        self.deck.close(quit_app=self.quit_on_exit if quit_app is None else quit_app)
        self.deck = None

    def _note_error(self, where: str, e: Exception) -> None:
        self._errors.append({"t": time.time(), "where": where,
                             "error": f"{type(e).__name__}: {e}"})
        if len(self._errors) > 50:
            del self._errors[: len(self._errors) - 50]
        print(f"[engine] {where} 异常：{e}", file=sys.stderr, flush=True)

    def shutdown(self) -> None:
        try:
            self.close_deck()
        except Exception:  # noqa: BLE001
            pass
        try:
            if self.live:
                self.live.stop()
        except Exception:  # noqa: BLE001
            pass


ENGINE = Engine()
atexit.register(ENGINE.shutdown)


# ---------------------------------------------------------------- JSON Schema 片段
def _obj(props: dict, required: Optional[list] = None, extra: bool = True) -> dict:
    s = {"type": "object", "properties": props, "additionalProperties": extra}
    if required:
        s["required"] = required
    return s


_COLOR = {"type": ["string", "array", "integer", "null"],
          "description": "颜色：\"#RRGGBB\" / [r,g,b] / 命名色（blue/red/ink/gray/white/...）"}
_BOX = {"type": "array", "items": {"type": "number"}, "minItems": 4, "maxItems": 4,
        "description": "[x, y, w, h]，单位 pt。画布默认 960x540，安全底线 y+h≤518"}

_ELEMENT_SCHEMA = {
    "type": "object",
    "properties": {
        "type": {"type": "string", "description": "元素类型，见 ppt_elements 返回的可用清单"},
        "box": _BOX,
        "x": {"type": "number"}, "y": {"type": "number"},
        "w": {"type": "number"}, "h": {"type": "number"},
        "text": {"type": "string"}, "title": {"type": "string"},
        "sec": {"type": "string"}, "label": {"type": "string"},
        "fs": {"type": "number", "description": "字号 pt"},
        "color": _COLOR, "fill": _COLOR, "line": _COLOR, "accent": _COLOR,
        "bold": {"type": "boolean"}, "italic": {"type": "boolean"},
        "align": {"type": "string", "enum": ["left", "center", "right", "justify"]},
        "anchor": {"type": "string", "enum": ["top", "middle", "bottom"]},
        "font": {"type": "string"},
        "parts": {"type": "array", "description": "分段着色：[[文字,颜色,是否加粗], ...]"},
        "items": {"type": "array", "description": "卡片/列表/时间轴条目"},
        "rows": {"type": "array", "description": "表格行或 gray_box 行"},
        "header": {"type": "array", "description": "表格表头"},
        "cols": {"type": "array", "description": "表格列宽或卡片列数"},
        "src": {"type": "string", "description": "图片路径或素材名"},
        "img": {"type": "string"},
        "chart": {"type": "string", "description": "图表类型：bar/hbar/line/area/pie/donut/scatter"},
        "data": {"type": "object", "description": "图表数据 {labels:[], series:[{name,values}]}"},
        "shape": {"type": "string"}, "direction": {"type": "string"},
        "radius": {"type": "number"}, "lw": {"type": "number"},
        "transparency": {"type": "number"}, "spacing": {"type": "number"},
        "numbered": {"type": "boolean"}, "fit": {"type": "boolean"},
        "mode": {"type": "string", "enum": ["auto", "native", "rects"]},
        "logo": {"type": "string"}, "subtitle": {"type": "string"},
        "meta": {"type": "string"}, "value": {"type": "string"},
        "note": {"type": "string"}, "author": {"type": "string"},
        "bullet": {"type": "string"}, "gap": {"type": "number"},
    },
    "required": ["type"],
    "additionalProperties": True,
}

_SLIDE_SCHEMA = {
    "type": "object",
    "properties": {
        "bg": {"type": ["string", "null"], "description": "本页底图（路径或素材名）"},
        "elements": {"type": "array", "items": _ELEMENT_SCHEMA},
    },
    "additionalProperties": True,
}


# ---------------------------------------------------------------- 工具实现
srv = McpStdioServer(
    instructions=(
        "dsh-ppt：用真实 Microsoft PowerPoint / WPS 演示做 PPT。"
        "标准流程：ppt_new（选主题）→ ppt_build（JSON 一次建整份）或 "
        "ppt_slide + ppt_elements（逐页精修）→ ppt_quality（pre-flight 审查）→ "
        "ppt_save（出 pptx+pdf）→ ppt_verify（PDF 实测溢出/重叠）。"
        "用户能在真实 PowerPoint 窗口和 DSH 预览面板里实时看到每一步。"
    )
)


@srv.tool("ppt_status", "环境与会话状态：Office 后端探测、当前 deck、实时预览地址、主题、Zotero 可用性。做任何 PPT 任务前先调这个。",
          _obj({}, extra=False), title="PPT 状态")
def _t_status(args):
    backends = O.probe_backends()
    live_res = ENGINE.start_live()
    return {
        "ok": True,
        "engine": {"version": "1.0.0", "uptimeSec": round(time.time() - ENGINE._started, 1),
                   "python": sys.version.split()[0], "pid": os.getpid()},
        "backends": backends,
        "activeBackend": (ENGINE.deck.app.backend["label"] if ENGINE.deck else None),
        "deck": (ENGINE.deck.state() if ENGINE.deck else None),
        "workdir": str(ENGINE.workdir),
        "assetsDir": str(ENGINE.assets_dir),
        "theme": ENGINE.theme.get("id"),
        "themes": T.theme_brief(),
        "visible": ENGINE.visible,
        "live": live_res,
        "liveRoutes": {"state": "/ppt/state", "snapshot": "/ppt/snapshot.png",
                       "ops": "/ppt/ops", "combined": "/ppt/live"},
        "elementTypes": R.route_names(),
        "chartTypes": C.chart_kinds(),
        "zotero": Z.probe(),
        "recentErrors": ENGINE._errors[-5:],
    }


@srv.tool("ppt_themes", "列出全部设计主题（配色/字号/间距/规则）。选主题时用。",
          _obj({"id": {"type": "string", "description": "指定主题 id 则返回完整定义"}}, extra=False),
          title="主题清单")
def _t_themes(args):
    tid = args.get("id")
    if tid:
        return {"ok": True, "theme": T.get_theme(tid)}
    return {"ok": True, "themes": T.theme_brief(),
            "layoutRules": T.BASE_LAYOUT,
            "dimensions": Q.REVIEW_DIMENSIONS}


@srv.tool("ppt_new", "新建演示文稿。会启动真实 PowerPoint 窗口（visible=true 时用户可实时观看）。已有未保存 deck 会先自动存到 recovery/ 再关。",
          _obj({
              "theme": {"type": "string", "description": "风格 id：default（干净商务，默认）。可用 ppt_themes 看全清单"},
              "width": {"type": "number", "description": "画布宽 pt，默认 960（16:9）"},
              "height": {"type": "number", "description": "画布高 pt，默认 540"},
              "bg": {"type": ["string", "null"], "description": "全局底图（路径或素材名，如 bg_page.png）"},
              "visible": {"type": ["boolean", "null"], "description": "是否弹出可见 PowerPoint 窗口（实时观看），默认 true"},
              "workdir": {"type": ["string", "null"], "description": "工作目录，默认 ~/Desktop/dsh-ppt-out"},
              "font": {"type": "string", "description": "中文字体，默认 微软雅黑（防方块）"},
              "backend": {"type": ["string", "null"], "description": "auto|ms-office|wps"},
              "footer": {"type": ["boolean", "null"], "description": "是否自动加页脚（页码+deck名），默认 true；封面/分隔页可单页关"},
              "footerText": {"type": ["string", "null"], "description": "页脚左侧文字，默认主题名"},
              "transition": {"type": ["string", "null"], "description": "整页切换：none/fade/push/wipe/split/cover/blinds，默认随主题（多为 fade 0.6s）"},
              "entrance": {"type": ["string", "null"], "description": "元素入场：none/fade/fly/appear/dissolve，默认 fade 错落；不需要动画传 none"},
          }, extra=False), title="新建演示")
def _t_new(args):
    if args.get("backend"):
        ENGINE.prefer_backend = str(args["backend"])
    d = ENGINE.new_deck(
        theme=args.get("theme") or "default",
        width=int(args.get("width") or 960),
        height=int(args.get("height") or 540),
        bg=args.get("bg"),
        visible=args.get("visible"),
        workdir=args.get("workdir"),
        font=args.get("font") or "微软雅黑",
        footer=args.get("footer"),
        footer_text=args.get("footerText"),
        transition=args.get("transition"),
        entrance=args.get("entrance"),
    )
    typo = ENGINE.theme.get("typo") or {}
    return {"ok": True, "deck": d.state(), "theme": ENGINE.theme,
            "typography": {"fonts": typo.get("resolvedFonts"), "scale": typo.get("scale"),
                           "chartPalette": typo.get("chartPalette")},
            "motion": {"transition": d.transition, "transitionDuration": d.transition_duration,
                       "entrance": d.entrance, "stagger": d.entrance_stagger},
            "live": (ENGINE.live.url if ENGINE.live else None),
            "hint": "下一步：ppt_build 一次建整份，或 ppt_slide + ppt_elements 逐页画。"
                    "封面/章节页建议 footer=false。"}


@srv.tool("ppt_open", "打开已有 pptx 继续编辑（可见模式下用户能实时看到修改）。",
          _obj({"path": {"type": "string", "description": "pptx 路径"},
                "visible": {"type": ["boolean", "null"]}},
               required=["path"], extra=False), title="打开演示")
def _t_open(args):
    if ENGINE.deck is not None:
        ENGINE._auto_recover()
    ENGINE.ensure_workdir()
    if args.get("visible") is not None:
        ENGINE.visible = bool(args["visible"])
    path = os.path.abspath(str(args["path"]))
    if not os.path.exists(path):
        return {"ok": False, "error": f"文件不存在：{path}"}
    th = dict(ENGINE.theme)
    th["assetsDir"] = str(ENGINE.assets_dir)
    el.set_assets_dir(ENGINE.assets_dir)
    ENGINE.deck = Deck(open_path=path, workdir=str(ENGINE.workdir),
                       visible=ENGINE.visible, font="微软雅黑",
                       prefer_backend=ENGINE.prefer_backend, live=True,
                       snapshot_scale=ENGINE.snapshot_scale, theme=th)
    ENGINE.start_live()
    return {"ok": True, "path": path, "deck": ENGINE.deck.state()}


@srv.tool("ppt_slide", "新增一页（可指定底图）。返回新页码。",
          _obj({"bg": {"type": ["string", "null"], "description": "本页底图，缺省用全局底图"},
                "elements": {"type": "array", "items": _ELEMENT_SCHEMA,
                             "description": "可选：建页同时把元素画上（省一次往返）"},
                "goto": {"type": "boolean", "description": "是否把可见窗口跳到本页，默认 true"}},
               extra=False), title="新增页")
def _t_slide(args):
    d = ENGINE.require_deck()
    s = d.add_slide(bg=args.get("bg"), goto=bool(args.get("goto", True)))
    idx = d.slide_count
    res = {"ok": True, "index": idx, "slideCount": idx}
    if args.get("elements"):
        res["elements"] = d.render_many(args["elements"], slide_index=idx)
        res["state"] = d.state()
    return res


@srv.tool("ppt_element", "在当前页（或指定页）画一个元素。元素类型见 ppt_status 的 elementTypes。",
          _obj({"element": _ELEMENT_SCHEMA,
                "slide": {"type": ["integer", "null"], "description": "页码，缺省为当前页"}},
               required=["element"], extra=False), title="画元素")
def _t_element(args):
    d = ENGINE.require_deck()
    out = d.render(args["element"], slide_index=args.get("slide"))
    return {"ok": True, "type": args["element"].get("type"),
            "slide": int(args.get("slide") or d.current_slide),
            "state": d.state()}


@srv.tool("ppt_elements", "在当前页（或指定页）批量画多个元素 —— 比逐个调 ppt_element 快得多，且只在最后刷新一次预览。",
          _obj({"elements": {"type": "array", "items": _ELEMENT_SCHEMA},
                "slide": {"type": ["integer", "null"]}},
               required=["elements"], extra=False), title="批量画元素")
def _t_elements(args):
    d = ENGINE.require_deck()
    res = d.render_many(args["elements"], slide_index=args.get("slide"))
    ok = sum(1 for r in res if r.get("ok"))
    return {"ok": True, "total": len(res), "succeeded": ok, "failed": len(res) - ok,
            "results": res, "state": d.state()}


@srv.tool("ppt_build", "JSON 驱动一次建整份 PPT —— 生产主力工具。传 slides[] 数组，每项含 bg + elements[]。会逐页建并实时刷新预览。",
          _obj({
              "slides": {"type": "array", "items": _SLIDE_SCHEMA, "description": "页面数组，按顺序生成"},
              "theme": {"type": ["string", "null"], "description": "可选：本次构建换主题"},
              "append": {"type": "boolean", "description": "true=在现有 deck 后追加；false（默认）=没有 deck 时新建"},
              "preflight": {"type": "boolean", "description": "构建前先做质量静态审查，默认 true；有 error 级问题时仍会构建但把问题一并返回"},
          }, required=["slides"], extra=False), title="JSON 建整份")
def _t_build(args):
    slides = args.get("slides") or []
    if not slides:
        return {"ok": False, "error": "slides 为空"}
    theme_id = args.get("theme") or (ENGINE.theme.get("id") if ENGINE.deck else "default")
    if ENGINE.deck is None or not args.get("append"):
        if ENGINE.deck is None:
            ENGINE.new_deck(theme=theme_id)
    d = ENGINE.require_deck()

    pre = None
    if args.get("preflight", True):
        try:
            pre = Q.check_deck({"slides": slides, "theme": theme_id,
                                "canvas": {"w": d.width, "h": d.height}}, theme=ENGINE.theme)
        except Exception as e:  # noqa: BLE001
            ENGINE._note_error("preflight", e)

    report = d.build({"slides": slides})
    out = {"ok": True, "theme": theme_id, "slideCount": d.slide_count,
           "built": report, "state": d.state()}
    if pre:
        out["preflight"] = {"overallScore": pre["overallScore"], "pass": pre["pass"],
                            "errors": pre["errors"], "warnings": pre["warnings"],
                            "summary": pre["summary"],
                            "issues": pre["issues"][:30]}
    return out


@srv.tool("ppt_chart", "用 matplotlib 生成图表 PNG 并插入当前页。支持 bar/hbar/line/area/pie/donut/scatter/bubble。中文字体已配好，不会出方块。",
          _obj({
              "chart": {"type": "string", "description": "图表类型"},
              "data": {"type": "object", "description": "{labels:[...], series:[{name,values:[...]}], xlabel, ylabel, colors:[...], stacked:bool, area:bool, donut:bool}"},
              "box": _BOX,
              "slide": {"type": ["integer", "null"]},
              "title": {"type": "string"},
              "dpi": {"type": "integer", "description": "默认 160"},
              "outName": {"type": ["string", "null"], "description": "指定输出文件名（默认按内容哈希命名，可复用缓存）"},
              "insert": {"type": "boolean", "description": "false=只生成 PNG 不插入，返回路径；默认 true"},
          }, required=["chart", "data"], extra=False), title="生成图表")
def _t_chart(args):
    ENGINE.ensure_workdir()
    png = C.make_chart(args["chart"], args.get("data") or {},
                       out_dir=(ENGINE.deck.workdir / ".charts" if ENGINE.deck else ENGINE.workdir / ".charts"),
                       title=args.get("title", ""), dpi=int(args.get("dpi") or 160),
                       theme=ENGINE.theme, out_name=args.get("outName"))
    out = {"ok": True, "png": png, "size": os.path.getsize(png),
           "chart": args["chart"], "kinds": C.chart_kinds()}
    if args.get("insert", True):
        d = ENGINE.require_deck()
        box = args.get("box") or [30, 100, 420, 300]
        d.render({"type": "pic", "box": box, "src": png, "fit": True},
                 slide_index=args.get("slide"))
        out["inserted"] = {"slide": int(args.get("slide") or d.current_slide), "box": box}
        out["state"] = d.state()
    return out


@srv.tool("ppt_table", "在当前页插入表格。mode=native 用 PowerPoint 原生表格（用户可继续编辑），mode=rects 用矩形拼接（上游为规避 WPS 边框 bug 的方案）。",
          _obj({
              "header": {"type": "array", "items": {"type": "string"}},
              "rows": {"type": "array", "description": "[[cell,...],...]，cell 可为字符串或 {text,color,bold,align}"},
              "box": _BOX,
              "cols": {"type": ["array", "null"], "description": "列宽权重，如 [2,3,1]"},
              "mode": {"type": "string", "enum": ["auto", "native", "rects"]},
              "slide": {"type": ["integer", "null"]},
              "hdrFs": {"type": "number"}, "cellFs": {"type": "number"},
              "headerFill": _COLOR, "altFill": _COLOR,
          }, required=["rows"], extra=False), title="插入表格")
def _t_table(args):
    d = ENGINE.require_deck()
    spec = {"type": "table", "box": args.get("box") or [16, 90, 928, 300],
            "header": args.get("header") or [], "rows": args.get("rows") or [],
            "cols": args.get("cols"), "mode": args.get("mode", "auto"),
            "hdrFs": args.get("hdrFs", 13), "cellFs": args.get("cellFs", 11),
            "headerFill": args.get("headerFill"), "altFill": args.get("altFill")}
    d.render(spec, slide_index=args.get("slide"))
    return {"ok": True, "rows": len(spec["rows"]), "mode": spec["mode"], "state": d.state()}


@srv.tool("ppt_image", "插入图片（本地绝对路径、相对工作目录路径、或素材名）。fit=true 时等比缩放居中防变形。",
          _obj({"src": {"type": "string", "description": "图片路径或素材名"},
                "box": _BOX, "fit": {"type": "boolean"},
                "slide": {"type": ["integer", "null"]}},
               required=["src"], extra=False), title="插入图片")
def _t_image(args):
    d = ENGINE.require_deck()
    p = el.resolve_image(args["src"])
    if p is None and ENGINE.assets_dir is not None:
        cand = ENGINE.assets_dir / str(args["src"])
        p = str(cand) if cand.exists() else None
    if p is None:
        return {"ok": False, "error": f"图片不存在：{args['src']}",
                "assetsDir": str(ENGINE.assets_dir),
                "availableAssets": [a["name"] for a in A.list_assets(ENGINE.assets_dir)]}
    d.render({"type": "pic", "box": args.get("box") or [30, 100, 400, 300],
              "src": p, "fit": args.get("fit", True)}, slide_index=args.get("slide"))
    return {"ok": True, "path": p, "state": d.state()}


@srv.tool("ppt_assets", "生成/列出 PIL 程序化素材（渐变横幅、竖条、圆环、箭头、页面背景、封面底图）。不依赖外部图床。",
          _obj({"action": {"type": "string", "enum": ["list", "generate", "theme"],
                           "description": "list=列出已有；generate=生成默认全套；theme=按当前主题生成"},
                "force": {"type": "boolean", "description": "true=已存在也重画"}},
               extra=False), title="素材管理")
def _t_assets(args):
    ENGINE.ensure_workdir()
    action = args.get("action") or "list"
    force = bool(args.get("force"))
    if action == "generate":
        A.ensure_default_assets(ENGINE.assets_dir, force=force)
    elif action == "theme":
        A.ensure_default_assets(ENGINE.assets_dir, force=force)
        A.ensure_theme_assets(ENGINE.theme, ENGINE.assets_dir, force=force)
    items = A.list_assets(ENGINE.assets_dir)
    return {"ok": True, "action": action, "assetsDir": str(ENGINE.assets_dir),
            "count": len(items), "assets": items}


@srv.tool("ppt_goto", "把可见 PowerPoint 窗口跳到指定页（不改内容）。用户实时观看时用。",
          _obj({"slide": {"type": "integer", "description": "页码，从 1 开始"}},
               required=["slide"], extra=False), title="跳页")
def _t_goto(args):
    d = ENGINE.require_deck()
    d.goto(int(args["slide"]))
    return {"ok": True, "currentSlide": d.current_slide, "slideCount": d.slide_count}


@srv.tool("ppt_remove_slide", "删除指定页。",
          _obj({"slide": {"type": "integer"}}, required=["slide"], extra=False), title="删除页")
def _t_remove(args):
    d = ENGINE.require_deck()
    d.remove_slide(int(args["slide"]))
    return {"ok": True, "slideCount": d.slide_count, "state": d.state()}


@srv.tool("ppt_snapshot", "导出某页为 PNG（默认当前页）。实时预览面板用的就是这个；也可单独取图给用户看。",
          _obj({"slide": {"type": ["integer", "null"], "description": "页码，缺省当前页"},
                "outPath": {"type": ["string", "null"], "description": "输出路径，缺省写实时预览缓存"},
                "scale": {"type": ["integer", "null"], "description": "导出宽度 px，默认 960"}},
               extra=False), title="导出页图")
def _t_snapshot(args):
    d = ENGINE.require_deck()
    if args.get("scale"):
        d.snapshot_scale = int(args["scale"])
    p = d.snapshot(slide_index=args.get("slide"), out_path=args.get("outPath"))
    if not p:
        return {"ok": False, "error": "导出失败（页码越界或 COM 异常）"}
    return {"ok": True, "path": p, "size": os.path.getsize(p),
            "slide": int(args.get("slide") or d.current_slide)}


@srv.tool("ppt_quality", "质量审查。pre-flight 模式对 slides JSON 做静态检查（画布溢出/标题间距/字号/中文引号/密度/视觉占比），不需要先建 deck；deck 模式审查已生成的内容。",
          _obj({"deckSpec": {"type": ["object", "null"], "description": "{slides:[{elements:[...]}]}，传了就走 pre-flight"},
                "theme": {"type": ["string", "null"]}},
               extra=False), title="质量审查")
def _t_quality(args):
    th = T.get_theme(args.get("theme") or ENGINE.theme.get("id"))
    spec = args.get("deckSpec")
    if spec:
        return {"ok": True, "mode": "preflight", "report": Q.check_deck(spec, theme=th)}
    d = ENGINE.require_deck()
    # deck 模式：从 COM 回读每页形状，做几何检查
    issues = []
    for i in range(1, d.slide_count + 1):
        try:
            s = d.slide(i)
            for sh in s.Shapes:
                try:
                    l, t, w, h = float(sh.Left), float(sh.Top), float(sh.Width), float(sh.Height)
                except Exception:  # noqa: BLE001
                    continue
                if t + h > float(th["layout"]["safeBottom"]) + 0.5:
                    issues.append({"slide": i, "level": "error", "rule": "canvas-overflow",
                                   "msg": f"形状 {sh.Name} 底部 {t + h:.0f} > {th['layout']['safeBottom']}"})
                if l + w > float(d.width) + 0.5:
                    issues.append({"slide": i, "level": "error", "rule": "canvas-overflow",
                                   "msg": f"形状 {sh.Name} 右侧 {l + w:.0f} > {d.width}"})
        except Exception as e:  # noqa: BLE001
            issues.append({"slide": i, "level": "warn", "rule": "read-error", "msg": str(e)})
    score = max(0, 100 - len([x for x in issues if x["level"] == "error"]) * 8
                - len([x for x in issues if x["level"] == "warn"]) * 2)
    return {"ok": True, "mode": "deck", "slideCount": d.slide_count,
            "report": {"overallScore": score, "pass": score >= 70,
                       "issues": issues[:60], "dimensions": Q.REVIEW_DIMENSIONS}}


@srv.tool("ppt_save", "保存 pptx（默认同时导 PDF），并做落盘存在性+大小校验（上游踩过 SaveAs 静默失败的坑）。",
          _obj({"path": {"type": ["string", "null"], "description": "输出路径，缺省 <workdir>/<时间戳>.pptx"},
                "pdf": {"type": ["boolean", "null"], "description": "是否同时导 PDF，默认跟随配置"},
                "verify": {"type": "boolean", "description": "保存后立即做 PDF 溢出校验，默认 true"},
                "close": {"type": "boolean", "description": "保存后是否关闭文稿，默认 false（保留窗口给用户看）"}},
               extra=False), title="保存")
def _t_save(args):
    d = ENGINE.require_deck()
    ENGINE.ensure_workdir()
    path = args.get("path") or str(ENGINE.workdir / f"deck-{time.strftime('%Y%m%d-%H%M%S')}.pptx")
    path = os.path.abspath(str(path))
    if not path.lower().endswith(".pptx"):
        path += ".pptx"
    do_pdf = ENGINE.auto_pdf if args.get("pdf") is None else bool(args["pdf"])
    res = d.save(path, pdf=do_pdf)
    out = {"ok": True, **res, "slideCount": d.slide_count}

    if args.get("verify", True) and res.get("pdf"):
        pv = V.verify_pdf(res["pdf"], width=float(d.width), height=float(d.height), verbose=False)
        xv = V.verify_pptx(res["pptx"]) if res.get("pptx") else None
        out["verify"] = pv
        out["pptxVerify"] = xv
        out["review"] = Q.merge_postflight(
            {"overallScore": 100.0, "issues": []}, pdf_verify=pv, pptx_verify=xv)
    if args.get("close"):
        ENGINE.close_deck()
        out["closed"] = True
    return out


@srv.tool("ppt_verify", "对已保存的 PDF/PPTX 做落盘校验：文本溢出画布、文本块重叠、空页、PPTX 与 PDF 页数一致性。",
          _obj({"pdf": {"type": ["string", "null"], "description": "PDF 路径，缺省用上次保存的"},
                "pptx": {"type": ["string", "null"]},
                "renderPages": {"type": "boolean", "description": "是否逐页渲染 PNG 供目检，默认 false"},
                "dpi": {"type": "integer"}},
               extra=False), title="落盘校验")
def _t_verify(args):
    d = ENGINE.deck
    pdf = args.get("pdf") or (d._saved_to and os.path.splitext(d._saved_to)[0] + ".pdf")
    pptx = args.get("pptx") or (d._saved_to if d else None)
    if not pdf and not pptx:
        return {"ok": False, "error": "没有可校验的文件，请先 ppt_save 或显式传 pdf/pptx 路径"}
    out = {"ok": True}
    pv = None
    if pdf and os.path.exists(str(pdf)):
        pv = V.verify_pdf(pdf, width=float(d.width if d else 960),
                          height=float(d.height if d else 540), verbose=False)
        out["pdf"] = pv
        if args.get("renderPages"):
            out["renderedPages"] = V.render_pages(pdf, str(ENGINE.workdir / "pages"),
                                                  dpi=int(args.get("dpi") or 110))
    if pptx and os.path.exists(str(pptx)):
        out["pptx"] = V.verify_pptx(pptx)
    out["review"] = Q.merge_postflight({"overallScore": 100.0, "issues": []},
                                       pdf_verify=pv, pptx_verify=out.get("pptx"))
    return out


@srv.tool("ppt_close", "关闭当前演示文稿。默认不退出 PowerPoint（保留窗口给用户查看/继续编辑）。",
          _obj({"quitApp": {"type": ["boolean", "null"], "description": "是否连 PowerPoint 一起退出"},
                "savePath": {"type": ["string", "null"], "description": "关闭前另存到该路径"}},
               extra=False), title="关闭文稿")
def _t_close(args):
    if ENGINE.deck is None:
        return {"ok": True, "note": "本来就没有打开的文稿"}
    n = ENGINE.deck.slide_count
    if args.get("savePath"):
        ENGINE.deck.save(str(args["savePath"]), pdf=False)
    ENGINE.close_deck(quit_app=args.get("quitApp"))
    return {"ok": True, "closedSlides": n, "quitApp": bool(args.get("quitApp"))}


@srv.tool("ppt_live", "实时预览控制：查看/开启 DSH 内实时预览服务地址与当前状态。用户「想看着 AI 写 PPT」时，把这个地址对应的面板打开即可。",
          _obj({"action": {"type": "string", "enum": ["status", "start", "state"],
                           "description": "status=服务信息；start=确保启动；state=当前 deck 状态+最近操作"}},
               extra=False), title="实时预览")
def _t_live(args):
    action = args.get("action") or "status"
    if action in ("start", "status"):
        res = ENGINE.start_live()
        info = {"ok": res.get("ok", False), "live": res,
                "routes": {"state": "/ppt/state", "snapshot": "/ppt/snapshot.png",
                           "ops": "/ppt/ops", "combined": "/ppt/live"},
                "note": "DSH 客户端面板会自动轮询；也可在浏览器直接开 " +
                        (ENGINE.live.url if ENGINE.live else "(未启动)")}
        if action == "state" or ENGINE.deck:
            info["deck"] = ENGINE.deck.state() if ENGINE.deck else None
        return info
    if ENGINE.deck is None:
        return {"ok": True, "deck": None, "note": "当前没有打开的文稿"}
    st = ENGINE.deck.state()
    png = ENGINE.deck.snapshot_path
    return {"ok": True, "deck": st,
            "snapshot": {"path": str(png), "exists": png.exists(),
                         "mtime": (png.stat().st_mtime if png.exists() else 0)}}


# ---------------------------------------------------------------- Zotero 组
@srv.tool("zotero_status", "探测 Zotero 环境：数据目录、zotero.sqlite、Local API、条目数。未安装时返回 available=false，不影响 PPT 功能。",
          _obj({}, extra=False), title="Zotero 状态")
def _t_z_status(args):
    return {"ok": True, **Z.probe()}


@srv.tool("zotero_search", "检索 Zotero 文献（只读直连 sqlite，Zotero 没开也能查）。支持按关键词/年份/类型/标签/分类过滤。",
          _obj({"query": {"type": "string", "description": "标题/作者/标签关键词，空则列最近添加"},
                "limit": {"type": "integer", "description": "默认 20"},
                "year": {"type": ["string", "null"]},
                "itemType": {"type": ["string", "null"], "description": "如 journalArticle/conferencePaper/book/thesis"},
                "tag": {"type": ["string", "null"]},
                "collection": {"type": ["string", "null"]}},
               extra=False), title="检索文献")
def _t_z_search(args):
    return {"ok": True, **Z.search(query=args.get("query", ""),
                                   limit=int(args.get("limit") or 20),
                                   year=args.get("year"), item_type=args.get("itemType"),
                                   tag=args.get("tag"), collection=args.get("collection"))}


@srv.tool("zotero_cite", "渲染引用串。本地实现 GB/T 7714 / APA / IEEE，不依赖 Zotero 运行时。",
          _obj({"refs": {"type": "array", "description": "文献 key 或 itemID 列表；也可传 search 返回的 item 对象"},
                "style": {"type": "string", "enum": ["gbt7714", "apa", "ieee", "plain"],
                          "description": "默认 gbt7714（中文汇报/论文标准）"},
                "numbered": {"type": "boolean", "description": "是否加 [1] [2] 编号，默认 true"}},
               required=["refs"], extra=False), title="渲染引用")
def _t_z_cite(args):
    style = args.get("style") or "gbt7714"
    items = []
    for r in args.get("refs") or []:
        if isinstance(r, dict) and r.get("title"):
            items.append(r)
        else:
            got = Z.get_item(str(r))
            if got.get("item"):
                items.append(got["item"])
    if not items:
        return {"ok": False, "error": "没有解析到任何文献条目", "style": style}
    cites = Z.render_bibliography(items, style, numbered=bool(args.get("numbered", True)))
    return {"ok": True, "style": style, "count": len(cites), "citations": cites,
            "items": [{"key": i.get("key"), "title": i.get("title"), "year": i.get("year")}
                      for i in items]}


@srv.tool("zotero_refs_slide", "把文献列表直接排成「参考文献页」的 slide spec，可原样塞进 ppt_build 的 slides[]。按条数自动分页，避免文字溢出。",
          _obj({"refs": {"type": "array", "description": "文献 key/itemID 列表，或 search 返回的 item 对象"},
                "style": {"type": "string", "enum": ["gbt7714", "apa", "ieee", "plain"]},
                "title": {"type": "string", "description": "页标题，默认「参考文献」"},
                "sec": {"type": "string", "description": "章节号，如 \"5.2\""},
                "perSlide": {"type": "integer", "description": "每页几条，默认 8"},
                "insert": {"type": "boolean", "description": "true=直接加到当前 deck 末尾；false（默认）=只返回 spec"}},
               required=["refs"], extra=False), title="参考文献页")
def _t_z_refs(args):
    style = args.get("style") or "gbt7714"
    items = []
    for r in args.get("refs") or []:
        if isinstance(r, dict) and r.get("title"):
            items.append(r)
        else:
            got = Z.get_item(str(r))
            if got.get("item"):
                items.append(got["item"])
    if not items:
        return {"ok": False, "error": "没有解析到任何文献条目"}
    slides = Z.refs_slide_spec(items, style=style, title=args.get("title") or "参考文献",
                               theme=ENGINE.theme, per_slide=int(args.get("perSlide") or 8),
                               sec=args.get("sec", ""))
    out = {"ok": True, "count": len(items), "pages": len(slides), "slides": slides}
    if args.get("insert"):
        d = ENGINE.require_deck()
        rep = d.build({"slides": slides})
        out["inserted"] = rep
        out["state"] = d.state()
    return out


@srv.tool("zotero_collections", "列出 Zotero 分类（文件夹），用于按分类取文献。",
          _obj({"limit": {"type": "integer"}}, extra=False), title="Zotero 分类")
def _t_z_coll(args):
    return {"ok": True, **Z.collections(limit=int(args.get("limit") or 100))}


# ---------------------------------------------------------------- 入口
def main() -> int:
    ENGINE.ensure_workdir()
    # 实时预览服务先起来 —— 即使还没建 deck，客户端面板也能拿到 health
    res = ENGINE.start_live()
    if res.get("ok"):
        print(f"[engine] 实时预览服务已启动：{res['url']}", file=sys.stderr, flush=True)
    else:
        print(f"[engine] 实时预览服务启动失败：{res.get('error')}", file=sys.stderr, flush=True)
    srv.serve()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
