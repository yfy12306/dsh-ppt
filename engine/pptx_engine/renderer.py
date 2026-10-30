# -*- coding: utf-8 -*-
"""元素类型路由 —— 把声明式 JSON spec 翻译成 elements.py 的绘图调用。

这是「JSON 数据驱动 PPT」的核心：模型只需产出 spec，不必懂 COM。
上游 harness-anything 的路由散落在各 build_*.py 里（13 种类型、每校一份副本），
本实现收敛成一张注册表，新增类型只要加一个 `@route("名字")` 函数。

spec 通用字段::

    type   必填，元素类型
    box    [x, y, w, h]（pt），部分元素用 x/y/w/h/d 单独给
    text   文本内容
    fs     字号
    color  颜色，接受 "#RRGGBB" / [r,g,b] / "blue"
    align  left|center|right|justify
    anchor top|middle|bottom
"""
from __future__ import annotations

from typing import Callable

from . import elements as el

_ROUTES: dict[str, Callable] = {}


def route(name: str):
    def deco(fn):
        _ROUTES[name] = fn
        return fn
    return deco


def route_names() -> list[str]:
    return sorted(_ROUTES.keys())


def render_element(slide, spec: dict, deck=None):
    """按 spec["type"] 路由到对应绘制函数。"""
    if not isinstance(spec, dict):
        raise ValueError(f"元素 spec 必须是对象，收到 {type(spec).__name__}")
    t = str(spec.get("type", "")).strip().lower()
    fn = _ROUTES.get(t)
    if fn is None:
        raise ValueError(f"未知元素类型：{t!r}。可用：{', '.join(route_names())}")
    return fn(slide, spec, deck)


# ---------------------------------------------------------------- 基础
def _box(spec, default=(0, 0, 100, 40)):
    return el._box(spec.get("box"), default)


@route("text")
@route("txt")
def _r_text(s, spec, deck):
    x, y, w, h = _box(spec, (16, 76, 928, 40))
    return el.txt(s, x, y, w, h, spec.get("text", ""),
                  fs=spec.get("fs", 14), color=spec.get("color", el.INK),
                  bold=spec.get("bold", False), align=spec.get("align", "left"),
                  anchor=spec.get("anchor", "middle"), font=spec.get("font"),
                  spacing=spec.get("spacing"), italic=spec.get("italic", False),
                  wrap=spec.get("wrap", True), role=spec.get("role"))


@route("rich")
def _r_rich(s, spec, deck):
    x, y, w, h = _box(spec, (16, 76, 928, 40))
    return el.rich(s, x, y, w, h, spec.get("parts") or [],
                   fs=spec.get("fs", 14), align=spec.get("align", "left"),
                   anchor=spec.get("anchor", "middle"), font=spec.get("font"),
                   spacing=spec.get("spacing"), role=spec.get("role"))


@route("rect")
def _r_rect(s, spec, deck):
    x, y, w, h = _box(spec)
    return el.rect(s, x, y, w, h, fill=spec.get("fill"), line=spec.get("line"),
                   lw=spec.get("lw", 1.0), transparency=spec.get("transparency"))


@route("rrect")
@route("rounded")
def _r_rrect(s, spec, deck):
    x, y, w, h = _box(spec)
    return el.rrect(s, x, y, w, h, fill=spec.get("fill"), line=spec.get("line"),
                    lw=spec.get("lw", 1.0), radius=spec.get("radius", 0.12),
                    transparency=spec.get("transparency"))


@route("oval")
@route("circle")
def _r_oval(s, spec, deck):
    x, y, w, h = _box(spec)
    return el.oval(s, x, y, w, h, fill=spec.get("fill"), line=spec.get("line"),
                   lw=spec.get("lw", 1.0))


@route("shape")
def _r_shape(s, spec, deck):
    x, y, w, h = _box(spec)
    return el.shape(s, spec.get("shape", "rect"), x, y, w, h,
                    fill=spec.get("fill"), line=spec.get("line"),
                    lw=spec.get("lw", 1.0), radius=spec.get("radius"),
                    transparency=spec.get("transparency"))


@route("arrow")
def _r_arrow(s, spec, deck):
    x, y, w, h = _box(spec, (16, 260, 200, 30))
    return el.arrow(s, x, y, w, h, fill=spec.get("fill", el.BLUE_M),
                    direction=spec.get("direction", "right"))


@route("hline")
def _r_hline(s, spec, deck):
    x = float(spec.get("x", 16))
    y = float(spec.get("y", 48))
    w = float(spec.get("w", 928))
    return el.hline(s, x, y, w, color=spec.get("color", el.GRAY_LN), wt=spec.get("wt", 1.0))


@route("vline")
def _r_vline(s, spec, deck):
    return el.vline(s, float(spec.get("x", 16)), float(spec.get("y", 12)),
                    float(spec.get("h", 30)), color=spec.get("color", el.GRAY_LN),
                    wt=spec.get("wt", 1.0))


@route("pic")
@route("image")
def _r_pic(s, spec, deck):
    x, y, w, h = _box(spec, (30, 100, 400, 300))
    src = spec.get("src") or spec.get("img") or spec.get("name")
    return el.pic(s, src, x, y, w, h, fit=spec.get("fit", True),
                  assets_dir=spec.get("assetsDir"))


# ---------------------------------------------------------------- 组合
@route("header")
def _r_header(s, spec, deck):
    width = float(spec.get("width", (deck.width if deck else 960)))
    brand = None
    if deck is not None and spec.get("brand", True):
        b = (deck.theme or {}).get("brand")
        if b and b.get("onContent"):
            brand = b
    return el.header(s, spec.get("sec", ""), spec.get("title", ""),
                     logo=spec.get("logo", ""), width=width,
                     fs=spec.get("fs"),
                     color=el.parse_color(spec.get("color")) or el.BLUE,
                     line_color=el.parse_color(spec.get("lineColor")) or (96, 148, 210),
                     accent_img=spec.get("accentImg", "accent_bar.png"),
                     assets_dir=spec.get("assetsDir"),
                     brand=brand)


@route("footer")
def _r_footer(s, spec, deck):
    w = float(spec.get("width", (deck.width if deck else 960)))
    hgt = float(spec.get("height", (deck.height if deck else 540)))
    return el.footer(s, width=w, height=hgt, left=spec.get("left", ""),
                     right=spec.get("right", ""), page=spec.get("page"),
                     total=spec.get("total"),
                     color=el.parse_color(spec.get("color")) or el.GRAY_TX,
                     rule=el.parse_color(spec.get("rule")) or el.GRAY_LN,
                     fs=spec.get("fs") or el.scale_size("caption", 11),
                     y=spec.get("y"))


@route("grad_bar")
def _r_grad_bar(s, spec, deck):
    x, y, w, h = _box(spec, (16, 186, 300, 32))
    return el.grad_bar(s, spec.get("img", "bar_blue.png"), x, y, w, h,
                       spec.get("text", ""), fs=spec.get("fs", 15),
                       color=spec.get("color", el.WHITE), assets_dir=spec.get("assetsDir"))


@route("gray_box")
def _r_gray_box(s, spec, deck):
    x, y, w, h = _box(spec, (16, 96, 928, 82))
    return el.gray_box(s, x, y, w, h, spec.get("rows") or [], fs=spec.get("fs", 14),
                       numbered=spec.get("numbered", True),
                       fill=el.parse_color(spec.get("fill")) or el.GRAY_BG,
                       line=el.parse_color(spec.get("line")) or el.GRAY_LN)


@route("pill")
def _r_pill(s, spec, deck):
    x, y, w, h = _box(spec, (286, 236, 54, 26))
    return el.pill(s, x, y, w, h, spec.get("text", ""), fs=spec.get("fs", 12.5),
                   fill=el.parse_color(spec.get("fill")) or el.RED,
                   color=el.parse_color(spec.get("color")) or el.WHITE,
                   rounded=spec.get("rounded", True), radius=spec.get("radius", 0.5))


@route("chip")
def _r_chip(s, spec, deck):
    x, y, w, h = _box(spec, (100, 100, 160, 30))
    return el.chip(s, x, y, w, h, spec.get("parts") or [[spec.get("text", ""), el.INK, False]],
                   fs=spec.get("fs", 12.5),
                   fill=el.parse_color(spec.get("fill")) or (249, 250, 252),
                   line=el.parse_color(spec.get("line")) or (206, 214, 224))


@route("num_circle")
def _r_num_circle(s, spec, deck):
    x = float(spec.get("x", 100))
    y = float(spec.get("y", 100))
    d = float(spec.get("d", 22))
    return el.num_circle(s, x, y, d, str(spec.get("label", "")),
                         fill=el.parse_color(spec.get("fill")) or el.BLUE_M,
                         fs=spec.get("fs", 12),
                         color=el.parse_color(spec.get("color")) or el.WHITE)


@route("cards")
@route("cards_2x3")
@route("cards_1x4")
def _r_cards(s, spec, deck):
    x, y, w, h = _box(spec, (16, 90, 928, 380))
    t = str(spec.get("type", "")).lower()
    cols = spec.get("cols")
    if cols is None:
        cols = 3 if t == "cards_2x3" else (4 if t == "cards_1x4" else 3)
    return el.cards(s, x, y, w, h, spec.get("items") or [], cols=int(cols),
                    fill=el.parse_color(spec.get("fill")) or el.WHITE,
                    line=el.parse_color(spec.get("line")) or el.GRAY_LN,
                    title_fs=spec.get("titleFs") or el.scale_size("h2", 20),
                    body_fs=spec.get("bodyFs") or el.scale_size("small", 14),
                    gap=spec.get("gap", 12),
                    title_color=el.parse_color(spec.get("titleColor")) or el.BLUE,
                    body_color=el.parse_color(spec.get("bodyColor")) or el.GRAY_TX,
                    accent=spec.get("accent"))


@route("cards_1x4_info")
@route("stat_cards")
def _r_stat_cards(s, spec, deck):
    x, y, w, h = _box(spec, (16, 90, 928, 110))
    return el.stat_cards(s, x, y, w, h, spec.get("items") or [],
                         fill=el.parse_color(spec.get("fill")) or el.WHITE,
                         line=el.parse_color(spec.get("line")) or el.GRAY_LN,
                         gap=spec.get("gap", 12),
                         num_fs=spec.get("numFs") or el.scale_size("numeric", 44),
                         label_fs=spec.get("labelFs") or el.scale_size("caption", 11),
                         color=el.parse_color(spec.get("color")) or el.BLUE)


@route("card_list_wide")
def _r_card_list_wide(s, spec, deck):
    """目录编号列表（上游 13 种元素之一）。"""
    x, y, w, h = _box(spec, (16, 90, 928, 380))
    items = spec.get("items") or []
    n = max(1, len(items))
    rh = h / n
    out = []
    for i, it in enumerate(items):
        if isinstance(it, dict):
            t, b = str(it.get("title", "")), str(it.get("body", ""))
        else:
            t, b = str(it[0]), (str(it[1]) if len(it) > 1 else "")
        ry = y + i * rh
        el.rrect(s, x, ry + 4, w, rh - 8, fill=el.WHITE, line=el.GRAY_LN, lw=0.9, radius=0.1)
        el.num_circle(s, x + 14, ry + rh / 2 - 13, 26, str(i + 1),
                      fill=el.parse_color(spec.get("accent")) or el.BLUE_M, fs=13)
        el.txt(s, x + 52, ry + 8, w * 0.3, rh - 16, t, fs=spec.get("titleFs", 15),
               color=el.parse_color(spec.get("titleColor")) or el.BLUE, bold=True, align="left")
        if b:
            el.txt(s, x + 52 + w * 0.32, ry + 8, w * 0.62, rh - 16, b,
                   fs=spec.get("bodyFs", 12), color=el.GRAY_TX, align="left")
        out.append(ry)
    return out


@route("timeline_horiz")
@route("timeline")
def _r_timeline(s, spec, deck):
    x, y, w, h = _box(spec, (40, 200, 880, 140))
    return el.timeline(s, x, y, w, h, spec.get("items") or [],
                       color=el.parse_color(spec.get("color")) or el.BLUE_M,
                       dot_d=spec.get("dotD", 18), fs=spec.get("fs", 12),
                       label_fs=spec.get("labelFs", 11))


@route("tagline_bar")
def _r_tagline(s, spec, deck):
    x, y, w, h = _box(spec, (16, 486, 928, 32))
    theme = (deck.theme if deck else None) or {}
    # 底色/文字色跟主题走（品牌主题用浅绿底+绿字），不再硬编码蓝灰
    fill = el.parse_color(spec.get("fill")) or el.parse_color(theme.get("light")) or (238, 244, 252)
    color = el.parse_color(spec.get("color")) or el.parse_color(theme.get("primary")) or el.BLUE
    return el.tagline_bar(s, x, y, w, h, spec.get("text", ""), fill=fill,
                          color=color, fs=spec.get("fs", 14))


@route("table")
def _r_table(s, spec, deck):
    x, y, w, h = _box(spec, (48, 90, 864, 300))
    small = el.scale_size("small", 14)
    return el.table(s, x, y, w, h=h if h else None,
                    header_row=spec.get("header") or spec.get("headerRow") or [],
                    rows=spec.get("rows") or [], cols=spec.get("cols"),
                    mode=spec.get("mode", "auto"),
                    hdr_fs=spec.get("hdrFs") or max(12, small),
                    cell_fs=spec.get("cellFs") or max(11, small - 1),
                    header_fill=el.parse_color(spec.get("headerFill")) or (32, 84, 160),
                    alt_fill=el.parse_color(spec.get("altFill")) or (238, 244, 252),
                    align=spec.get("align", "center"))


@route("chart")
def _r_chart(s, spec, deck):
    """matplotlib 生成图表 PNG 后插入（上游流程里「搜数据→生成图表→写 JSON」的自动化版）。"""
    from . import charts
    x, y, w, h = _box(spec, (30, 100, 420, 300))
    png = charts.make_chart(spec.get("chart") or spec.get("kind") or "bar",
                            spec.get("data") or {},
                            out_dir=(deck.workdir / ".charts" if deck else None),
                            title=spec.get("title", ""), dpi=spec.get("dpi", 160),
                            theme=spec.get("theme"), figsize=spec.get("figsize"))
    return el.pic(s, png, x, y, w, h, fit=spec.get("fit", True))


@route("bullets")
def _r_bullets(s, spec, deck):
    """项目符号列表。字号默认走主题 body（调研：正文 ≥18pt，3-4 条为宜）。"""
    x, y, w, h = _box(spec, (48, 90, 864, 380))
    items = spec.get("items") or []
    fs = spec.get("fs") or el.scale_size("body", 20)
    color = el.parse_color(spec.get("color")) or el.INK
    bullet = spec.get("bullet", "•")
    bcolor = el.parse_color(spec.get("bulletColor")) or el.BLUE_M
    gap = spec.get("gap")
    n = max(1, len(items))
    rh = float(gap) if gap else h / n
    out = []
    for i, it in enumerate(items):
        ry = y + i * rh
        el.txt(s, x, ry, 24, rh, bullet, fs=fs, color=bcolor, bold=True,
               align="left", anchor="top", role="body")
        if isinstance(it, dict):
            parts = [[str(it.get("text", "")), color, bool(it.get("bold", False))]]
            if it.get("sub"):
                parts.append([f"  {it['sub']}", el.GRAY_TX, False])
            el.rich(s, x + 30, ry, w - 30, rh, parts, fs=fs, align="left",
                    anchor="top", role="body")
        else:
            el.txt(s, x + 30, ry, w - 30, rh, str(it), fs=fs, color=color,
                   align="left", anchor="top", role="body")
        out.append(ry)
    return out


@route("kpi")
def _r_kpi(s, spec, deck):
    """单个大数字 KPI（标题 + 数值 + 说明）。"""
    x, y, w, h = _box(spec, (16, 90, 300, 140))
    el.rrect(s, x, y, w, h, fill=el.parse_color(spec.get("fill")) or el.WHITE,
             line=el.parse_color(spec.get("line")) or el.GRAY_LN, lw=0.9, radius=0.08)
    el.txt(s, x, y + 10, w, 22, spec.get("label", ""), fs=spec.get("labelFs", 12),
           color=el.GRAY_TX, align="center")
    el.txt(s, x, y + 34, w, h * 0.4, spec.get("value", ""), fs=spec.get("valueFs", 34),
           color=el.parse_color(spec.get("color")) or el.BLUE, bold=True, align="center")
    if spec.get("note"):
        el.txt(s, x + 10, y + h - 30, w - 20, 24, spec["note"], fs=spec.get("noteFs", 10.5),
               color=el.GRAY_TX, align="center")
    return (x, y, w, h)


@route("quote")
def _r_quote(s, spec, deck):
    x, y, w, h = _box(spec, (60, 180, 840, 140))
    el.rect(s, x, y, 5, h, fill=el.parse_color(spec.get("accent")) or el.BLUE_M)
    el.txt(s, x + 24, y + 8, w - 40, h - 46, spec.get("text", ""),
           fs=spec.get("fs", 18), color=el.parse_color(spec.get("color")) or el.INK,
           italic=True, align="left", anchor="top")
    if spec.get("author"):
        el.txt(s, x + 24, y + h - 32, w - 40, 24, f"— {spec['author']}",
               fs=spec.get("authorFs", 12), color=el.GRAY_TX, align="left")
    return (x, y, w, h)


@route("cover")
def _r_cover(s, spec, deck):
    """封面页。两种版式：

    品牌版（品牌主题）：深绿满底 + 左上浅色六边形标识 +
    右上白色 品牌字标 + 左对齐白色大标题 + 右下主绿色块 + 底部细线。
    默认版：留白 + 居中大标题 + 主题色细条（干净克制）。
    """
    w = float(spec.get("width", (deck.width if deck else 960)))
    h = float(spec.get("height", (deck.height if deck else 540)))
    theme = (deck.theme if deck else None) or {}
    brand = theme.get("brand")
    accent = el.parse_color(spec.get("accent")) or el.parse_color(theme.get("primary")) or el.BLUE

    if brand and spec.get("style", "brand") == "brand":
        deep = el.parse_color(theme.get("secondary")) or (16, 41, 26)
        el.rect(s, 0, 0, w, h, fill=deep)
        # 右下主绿色块（参考图的对角色块构成），不压文字区
        el.rect(s, w * 0.66, h * 0.66, w * 0.34, h * 0.34, fill=accent)
        try:
            el.pic(s, (brand.get("mark") or "").replace(".png", "_light.png"),
                   48, 40, 40, 40, fit=True)
        except Exception:  # noqa: BLE001
            pass
        el.txt(s, w - 48 - 260, 44, 260, 26, str(brand.get("wordmark", "")),
               fs=15, color=el.WHITE, bold=True, align="right", role="mono")
        el.txt(s, 48, h * 0.34, w - 200, 90, spec.get("title", ""),
               fs=spec.get("titleFs") or el.scale_size("display", 46),
               color=el.WHITE, bold=True, align="left", role="display")
        el.rect(s, 52, h * 0.34 + 100, 72, 5, fill=accent)
        if spec.get("subtitle"):
            el.txt(s, 48, h * 0.34 + 122, w - 240, 40, spec["subtitle"],
                   fs=spec.get("subtitleFs") or el.scale_size("body", 20),
                   color=(214, 226, 210), align="left", role="body")
        if spec.get("meta"):
            el.txt(s, 48, h - 64, w - 96, 26, spec["meta"],
                   fs=spec.get("metaFs") or el.scale_size("caption", 12),
                   color=(168, 184, 164), align="left", role="caption")
        return (0, 0, w, h)

    el.txt(s, 0, spec.get("titleY", h * 0.32), w, spec.get("titleH", 80),
           spec.get("title", ""), fs=spec.get("titleFs") or el.scale_size("display", 46),
           color=accent, bold=True, align="center", role="display")
    if spec.get("subtitle"):
        el.txt(s, 0, spec.get("subtitleY", h * 0.32 + 92), w, 36, spec["subtitle"],
               fs=spec.get("subtitleFs") or el.scale_size("body", 20),
               color=el.GRAY_TX, align="center", role="body")
    if spec.get("meta"):
        el.txt(s, 0, h - 78, w, 26, spec["meta"],
               fs=spec.get("metaFs") or el.scale_size("caption", 12),
               color=el.GRAY_TX, align="center", role="caption")
    el.rect(s, w * 0.36, h - 42, w * 0.28, 4, fill=accent)
    return (0, 0, w, h)


@route("divider")
@route("section")
def _r_divider(s, spec, deck):
    """章节分隔页（调研 Layout 8：title-only，大字号，≤5 词，给观众心理重置）。

    品牌版（品牌主题）：深绿满底 + 通栏主绿横带 + 白色大编号/大标题。
    默认版：留白 + 大编号 + 大标题 + 一句副题。
    """
    w = float(spec.get("width", (deck.width if deck else 960)))
    h = float(spec.get("height", (deck.height if deck else 540)))
    theme = (deck.theme if deck else None) or {}
    brand = theme.get("brand")
    primary = el.parse_color(spec.get("color")) or el.parse_color(theme.get("primary")) or el.BLUE

    if brand and spec.get("style", "brand") == "brand":
        deep = el.parse_color(theme.get("secondary")) or (16, 41, 26)
        el.rect(s, 0, 0, w, h, fill=deep)
        el.rect(s, 0, h * 0.44, w, 56, fill=primary)          # 通栏绿带（参考图）
        if spec.get("sec"):
            el.txt(s, 48, h * 0.16, 300, 90, str(spec["sec"]), fs=64,
                   color=primary, bold=True, align="left", role="numeric")
        el.txt(s, 48, h * 0.44 + 76, w - 160, 70, str(spec.get("title", "")),
               fs=spec.get("fs") or el.scale_size("display", 44),
               color=el.WHITE, bold=True, align="left", role="display")
        if spec.get("subtitle"):
            el.txt(s, 48, h * 0.44 + 156, w - 240, 34, spec["subtitle"],
                   fs=el.scale_size("body", 19), color=(190, 205, 186),
                   align="left", role="caption")
        return (0, 0, w, h)

    return el.divider(s, width=w, height=h, sec=spec.get("sec", ""),
                      title=spec.get("title", ""), subtitle=spec.get("subtitle", ""),
                      primary=primary, bg=spec.get("bg"),
                      text_color=spec.get("textColor"), accent=spec.get("accent"))


@route("toc")
def _r_toc(s, spec, deck):
    """目录页：两列编号列表。"""
    x, y, w, h = _box(spec, (60, 100, 840, 360))
    items = spec.get("items") or []
    cols = int(spec.get("cols", 2))
    n = len(items)
    per = (n + cols - 1) // cols if cols > 0 else n
    accent = el.parse_color(spec.get("accent")) or el.BLUE_M
    out = []
    for i, it in enumerate(items):
        c, r = divmod(i, per) if per else (0, i)
        cw = w / max(1, cols)
        cx = x + c * cw
        cy = y + r * (h / max(1, per))
        label = str(it.get("title", it) if isinstance(it, dict) else it)
        el.num_circle(s, cx, cy + 6, 26, str(i + 1), fill=accent, fs=13)
        el.txt(s, cx + 38, cy + 6, cw - 50, 26, label, fs=spec.get("fs", 16),
               color=el.INK, bold=True, align="left")
        out.append((cx, cy))
    return out
