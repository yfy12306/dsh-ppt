# -*- coding: utf-8 -*-
"""绘图原语与组合组件 —— 从 harness-anything/pptengine 抽取并针对 MS Office 加固。

坐标单位一律为 pt，画布默认 960x540（16:9）。颜色统一接受三种写法：
`"#1A3C8B"` / `[26,60,139]` / `"blue"`，内部转成 Office 需要的 BGR 整数。

## 上游踩坑清单（本模块已内置规避，勿删这些注释）

1. `sp.Adjustments.Item(1) = x` 在 Python 里是**语法错误**，必须写 `sp.Adjustments[1] = x`。
2. `Font.Bold = False` 在部分 Office/WPS 版本会抛异常 —— 只在需要加粗时赋 `True`，
   需要取消加粗时赋 `0`（VBA 的 msoFalse）而不是 Python 的 `False`。
3. `WordWrap = False` 与居中/右对齐同用会让文字跑出画布 —— 恒设 `True`。
4. 多行文本 + 垂直居中在 Office 里会算错位置 —— 改顶对齐 + 手动下移。
5. `AddTable` 的单元格边框 API 在 WPS 上不稳定 —— 保留「矩形拼接」模式兜底，
   MS Office 上默认走原生 `AddTable`（实测稳定，且可被 PowerPoint 正常编辑）。
6. 中文字体固定 `微软雅黑`，避免渲染成方块。
7. 上游每次生成前 `taskkill /F /IM wps.exe` —— **本实现刻意不做**，
   那会毁掉用户未保存的文档；改为只关闭自己创建的文稿（见 office.py）。
"""
from __future__ import annotations

import os
from pathlib import Path
from typing import Optional, Sequence, Union

from . import office as O

# ---------------------------------------------------------------- 全局可调
FONT = "微软雅黑"
TITLE_FONT = "微软雅黑"

# 色板（与上游 pptengine 保持一致，便于直接复用其 JSON 数据）
INK = (32, 32, 38)
WHITE = (255, 255, 255)
BLUE = (16, 62, 140)
BLUE_M = (36, 92, 170)
BLUE_L = (214, 231, 248)
TEAL = (18, 104, 124)
RED = (192, 24, 24)
ORANGE = (232, 126, 34)
GREEN = (24, 130, 84)
GRAY_BG = (243, 245, 248)
GRAY_TX = (88, 94, 104)
GRAY_LN = (210, 217, 226)

_NAMED = {
    "ink": INK, "white": WHITE, "blue": BLUE, "blue_m": BLUE_M, "blue_l": BLUE_L,
    "teal": TEAL, "red": RED, "orange": ORANGE, "green": GREEN,
    "gray": GRAY_TX, "gray_bg": GRAY_BG, "gray_ln": GRAY_LN,
    "black": (0, 0, 0), "yellow": (240, 196, 28), "purple": (102, 8, 116),
    "transparent": None, "none": None,
}

_ASSETS_DIR: Optional[Path] = None

# 主题排版（由 Deck 注入）：角色→本机真实字体族名；角色→字号 pt。
# 角色：display 封面大标题 / title 页标题 / body 正文 / caption 注释 /
#       numeric 大数字 / mono 编号与代码 / accent 引用与点缀
_ROLE_FONTS: dict = {}
_SCALE: dict = {}


def set_typography(resolved_fonts: Optional[dict] = None,
                   scale: Optional[dict] = None) -> None:
    global _ROLE_FONTS, _SCALE
    _ROLE_FONTS = dict(resolved_fonts or {})
    _SCALE = dict(scale or {})


def role_font(role: str, fallback: str = FONT) -> str:
    return _ROLE_FONTS.get(role) or fallback


def scale_size(role: str, fallback: float = 14.0) -> float:
    try:
        return float(_SCALE.get(role, fallback))
    except (TypeError, ValueError):
        return float(fallback)

# 对齐/锚点字符串 → COM 常量
_ALIGN = {"left": O.ppAlignLeft, "center": O.ppAlignCenter, "right": O.ppAlignRight,
          "justify": O.ppAlignJustify, 0: O.ppAlignLeft, 1: O.ppAlignCenter,
          2: O.ppAlignRight, 3: O.ppAlignJustify}
_ANCHOR = {"top": O.msoAnchorTop, "middle": O.msoAnchorMiddle, "bottom": O.msoAnchorBottom,
           1: O.msoAnchorTop, 3: O.msoAnchorMiddle, 4: O.msoAnchorBottom}
_SHAPE = {
    "rect": O.msoShapeRectangle, "rectangle": O.msoShapeRectangle,
    "rrect": O.msoShapeRoundedRectangle, "rounded": O.msoShapeRoundedRectangle,
    "oval": O.msoShapeOval, "ellipse": O.msoShapeOval, "circle": O.msoShapeOval,
    "diamond": O.msoShapeDiamond, "triangle": O.msoShapeIsocelesTriangle,
    "rt_triangle": O.msoShapeRightTriangle, "parallelogram": O.msoShapeParallelogram,
    "arrow": O.msoShapeRightArrow, "right_arrow": O.msoShapeRightArrow,
    "left_arrow": O.msoShapeLeftArrow, "up_arrow": O.msoShapeUpArrow,
    "down_arrow": O.msoShapeDownArrow, "chevron": O.msoShapeChevron,
    "pentagon": O.msoShapePentagon, "hexagon": O.msoShapeHexagon, "cloud": O.msoShapeCloud,
    "process": O.msoShapeFlowchartProcess, "decision": O.msoShapeFlowchartDecision,
}


# ---------------------------------------------------------------- 工具函数
def set_assets_dir(p) -> None:
    global _ASSETS_DIR
    _ASSETS_DIR = Path(p).expanduser().resolve()


def get_assets_dir() -> Optional[Path]:
    return _ASSETS_DIR


def parse_color(c, default=None):
    """把 '#RRGGBB' / '#RGB' / [r,g,b] / (r,g,b) / 'name' / int 归一成 (r,g,b) 或 None。"""
    if c is None:
        return default
    if isinstance(c, str):
        s = c.strip()
        if not s:
            return default
        low = s.lower()
        if low in _NAMED:
            return _NAMED[low]
        if s.startswith("#"):
            s = s[1:]
        if len(s) == 3:
            s = "".join(ch * 2 for ch in s)
        if len(s) == 6:
            try:
                return (int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))
            except ValueError:
                return default
        if len(s) == 8:  # #RRGGBBAA —— 忽略 alpha，Office 形状填充另走 Transparency
            try:
                return (int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))
            except ValueError:
                return default
        return default
    if isinstance(c, int):
        # 已经是 OLE 颜色（BGR）—— 反解回 RGB 以便统一处理
        return (c & 0xFF, (c >> 8) & 0xFF, (c >> 16) & 0xFF)
    if isinstance(c, (list, tuple)) and len(c) >= 3:
        return (int(c[0]), int(c[1]), int(c[2]))
    return default


def h2b(c) -> int:
    """(R,G,B) → Office COM 需要的 BGR 整数（等价 VBA 的 RGB() 函数）。"""
    r, g, b = c[0], c[1], c[2]
    return (b << 16) | (g << 8) | r


def ole(c) -> int:
    return h2b(parse_color(c) or INK)


def _box(box, default=(0, 0, 100, 40)):
    """把 [x,y,w,h] / {x,y,w,h} 归一成 4 元 float 组。"""
    if box is None:
        return tuple(float(v) for v in default)
    if isinstance(box, dict):
        return (float(box.get("x", default[0])), float(box.get("y", default[1])),
                float(box.get("w", default[2])), float(box.get("h", default[3])))
    seq = list(box)
    while len(seq) < 4:
        seq.append(default[len(seq)])
    return tuple(float(v) for v in seq[:4])


def clamp_to_canvas(box, cw: float = 960.0, ch: float = 540.0, margin: float = 0.0):
    """把元素框收进画布内（上游经验：所有元素必须满足 y+h ≤ 518）。

    Returns:
        (box, overflow: bool) —— overflow=True 表示原始框超出画布、已被裁剪。
    """
    x, y, w, h = box
    over = False
    if x < margin:
        w += x - margin
        x = margin
        over = True
    if y < margin:
        h += y - margin
        y = margin
        over = True
    if x + w > cw - margin:
        w = cw - margin - x
        over = True
    if y + h > ch - margin:
        h = ch - margin - y
        over = True
    return (x, y, max(1.0, w), max(1.0, h)), over


# ---------------------------------------------------------------- 文字
def txt(s, x, y, w, h, text, fs=13, color=INK, bold=False, align="left",
        anchor="middle", font=None, spacing=None, italic=False, wrap=True,
        role=None):
    """单色文本框。

    align: left|center|right|justify（也接受 0/1/2/3 兼容上游 JSON）
    anchor: top|middle|bottom（也接受 1/3/4）
    role:  排版角色（display/title/body/caption/numeric/mono/accent）。
           给了 role 且未显式传 font/fs 时，按主题排版系统取字体与字号。
    """
    if role and font is None:
        font = role_font(role)
    if role and fs == 13:
        fs = scale_size(role, 13)
    tb = s.Shapes.AddTextbox(O.msoTextOrientationHorizontal, x, y, w, h)
    tf = tb.TextFrame
    # 坑 3：WordWrap 必须为 True，否则居中/右对齐时文字跑出画布
    try:
        tf.WordWrap = O.msoTrue if wrap else O.msoFalse
    except Exception:  # noqa: BLE001
        pass
    try:
        tf.MarginLeft = 0
        tf.MarginRight = 0
        tf.MarginTop = 0
        tf.MarginBottom = 0
        tf.AutoSize = 0  # ppAutoSizeNone
    except Exception:  # noqa: BLE001
        pass

    text = "" if text is None else str(text)
    nlines = text.count("\n") + 1

    # 坑 4：多行 + 垂直居中在 Office 里算错位置 → 顶对齐 + 手动下移
    anchor_val = _ANCHOR.get(anchor, O.msoAnchorMiddle)
    if nlines > 1 and anchor_val == O.msoAnchorMiddle:
        anchor_val = O.msoAnchorTop
        try:
            tb.Top = y + max(0.0, (h - nlines * fs * 1.42) / 2.0)
        except Exception:  # noqa: BLE001
            pass
    try:
        tf.VerticalAnchor = anchor_val
    except Exception:  # noqa: BLE001
        pass

    tr = tf.TextRange
    tr.Text = text
    f = tr.Font
    try:
        f.Name = font or FONT
    except Exception:  # noqa: BLE001
        pass
    try:
        f.Size = float(fs)
    except Exception:  # noqa: BLE001
        pass
    try:
        f.Color.RGB = h2b(parse_color(color) or INK)
    except Exception:  # noqa: BLE001
        pass
    # 坑 2：Bold=False 在部分版本抛异常；取消加粗用 0（msoFalse）而非 Python False
    try:
        f.Bold = O.msoTrue if bold else 0
    except Exception:  # noqa: BLE001
        try:
            if bold:
                f.Bold = True
        except Exception:  # noqa: BLE001
            pass
    if italic:
        try:
            f.Italic = O.msoTrue
        except Exception:  # noqa: BLE001
            pass
    try:
        tr.ParagraphFormat.Alignment = _ALIGN.get(align, O.ppAlignLeft)
    except Exception:  # noqa: BLE001
        pass
    if spacing is not None:
        try:
            tr.ParagraphFormat.SpaceWithin = float(spacing)
        except Exception:  # noqa: BLE001
            pass
    return tb


def rich(s, x, y, w, h, parts, fs=13, align="left", anchor="middle", font=None,
         spacing=None, role=None):
    """分段着色文本。parts = [(文字, 颜色, 是否加粗), ...]，也接受 dict 形式。"""
    if role and font is None:
        font = role_font(role)
    if role and fs == 13:
        fs = scale_size(role, 13)
    norm = []
    for p in parts or []:
        if isinstance(p, dict):
            norm.append((str(p.get("text", "")), p.get("color", INK), bool(p.get("bold", False))))
        elif isinstance(p, (list, tuple)):
            t = str(p[0]) if len(p) > 0 else ""
            c = p[1] if len(p) > 1 else INK
            b = bool(p[2]) if len(p) > 2 else False
            norm.append((t, c, b))
        else:
            norm.append((str(p), INK, False))

    tb = txt(s, x, y, w, h, "".join(p[0] for p in norm), fs=fs, color=INK,
             bold=False, align=align, anchor=anchor, font=font, spacing=spacing)
    tf = tb.TextFrame
    pos = 1
    for t, c, b in norm:
        n = len(t)
        if n <= 0:
            continue
        try:
            rng = tf.TextRange.Characters(pos, n)
            rgb = parse_color(c)
            if rgb is not None:
                rng.Font.Color.RGB = h2b(rgb)
            if b:
                rng.Font.Bold = O.msoTrue
        except Exception:  # noqa: BLE001
            pass
        pos += n
    return tb


# ---------------------------------------------------------------- 图片
def resolve_image(name_or_path) -> Optional[str]:
    """把「素材名」或「路径」解析成真实存在的文件路径。"""
    if not name_or_path:
        return None
    p = Path(str(name_or_path)).expanduser()
    if p.is_absolute() and p.exists():
        return str(p)
    # 相对当前工作目录
    if p.exists():
        return str(p.resolve())
    # 相对素材目录
    if _ASSETS_DIR is not None:
        cand = _ASSETS_DIR / p.name
        if cand.exists():
            return str(cand)
    return None


def pic(s, name, x, y, w, h, fit=True, assets_dir=None):
    """插图；fit=True 时等比缩放并居中，防止拉伸变形与溢出。"""
    base = Path(assets_dir).expanduser().resolve() if assets_dir else _ASSETS_DIR
    p = None
    if base is not None:
        cand = base / str(name)
        if cand.exists():
            p = str(cand)
    if p is None:
        p = resolve_image(name)
    if p is None:
        raise FileNotFoundError(f"图片不存在：{name}（已查素材目录 {base} 与当前目录）")

    iw = ih = 0
    try:
        from PIL import Image
        with Image.open(p) as im:
            iw, ih = im.size
    except Exception:  # noqa: BLE001
        pass

    if fit and iw > 0 and ih > 0 and w > 0 and h > 0:
        sr, ir = w / h, iw / ih
        if ir > sr:
            nw, nh = w, int(w / ir)
        else:
            nh, nw = h, int(h * ir)
        x += (w - nw) / 2.0
        y += (h - nh) / 2.0
        w, h = nw, nh
    try:
        # AddPicture(FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height)
        return s.Shapes.AddPicture(str(p), O.msoFalse, O.msoTrue, x, y, w, h)
    except Exception as e:  # noqa: BLE001
        raise RuntimeError(f"插入图片失败 {p}: {e}") from e


# ---------------------------------------------------------------- 形状
def shape(s, kind, x, y, w, h, fill=None, line=None, lw=1.0, radius=None, transparency=None):
    """通用形状。kind 见 _SHAPE（rect/rrect/oval/arrow/...）。"""
    t = _SHAPE.get(str(kind).lower(), None)
    if t is None:
        try:
            t = int(kind)
        except (TypeError, ValueError):
            t = O.msoShapeRectangle
    sp = s.Shapes.AddShape(t, x, y, w, h)

    rgb = parse_color(fill)
    if rgb is not None:
        try:
            sp.Fill.Visible = O.msoTrue
            sp.Fill.ForeColor.RGB = h2b(rgb)
            if transparency is not None:
                sp.Fill.Transparency = max(0.0, min(1.0, float(transparency)))
        except Exception:  # noqa: BLE001
            pass
    else:
        try:
            sp.Fill.Visible = O.msoFalse
        except Exception:  # noqa: BLE001
            pass

    ln = parse_color(line)
    if ln is not None:
        try:
            sp.Line.Visible = O.msoTrue
            sp.Line.ForeColor.RGB = h2b(ln)
            sp.Line.Weight = float(lw)
        except Exception:  # noqa: BLE001
            pass
    else:
        try:
            sp.Line.Visible = O.msoFalse
        except Exception:  # noqa: BLE001
            pass

    # 坑 1：Adjustments 只能用下标，`.Item(1) = x` 是 Python 语法错误
    if radius is not None and t == O.msoShapeRoundedRectangle:
        try:
            sp.Adjustments[1] = max(0.0, min(0.5, float(radius)))
        except Exception:  # noqa: BLE001
            pass

    # 形状自带文字框默认有内边距与自动换行，清掉以便精确定位
    try:
        sp.TextFrame.MarginLeft = 0
        sp.TextFrame.MarginRight = 0
        sp.TextFrame.MarginTop = 0
        sp.TextFrame.MarginBottom = 0
        sp.TextFrame.WordWrap = O.msoTrue
    except Exception:  # noqa: BLE001
        pass
    return sp


def rect(s, x, y, w, h, fill=None, line=None, lw=1.0, transparency=None):
    return shape(s, "rect", x, y, w, h, fill, line, lw, transparency=transparency)


def rrect(s, x, y, w, h, fill=None, line=None, lw=1.0, radius=0.12, transparency=None):
    return shape(s, "rrect", x, y, w, h, fill, line, lw, radius=radius, transparency=transparency)


def oval(s, x, y, w, h, fill=None, line=None, lw=1.0):
    return shape(s, "oval", x, y, w, h, fill, line, lw)


def arrow(s, x, y, w, h, fill=BLUE_M, direction="right"):
    kind = {"right": "right_arrow", "left": "left_arrow",
            "up": "up_arrow", "down": "down_arrow"}.get(str(direction).lower(), "right_arrow")
    return shape(s, kind, x, y, w, h, fill=fill)


def hline(s, x, y, w, color=GRAY_LN, wt=1.0):
    ln = s.Shapes.AddLine(x, y, x + w, y)
    try:
        ln.Line.ForeColor.RGB = h2b(parse_color(color) or GRAY_LN)
        ln.Line.Weight = float(wt)
    except Exception:  # noqa: BLE001
        pass
    return ln


def vline(s, x, y, h, color=GRAY_LN, wt=1.0):
    ln = s.Shapes.AddLine(x, y, x, y + h)
    try:
        ln.Line.ForeColor.RGB = h2b(parse_color(color) or GRAY_LN)
        ln.Line.Weight = float(wt)
    except Exception:  # noqa: BLE001
        pass
    return ln


# ---------------------------------------------------------------- 组合组件
def num_circle(s, x, y, d, label, fill=BLUE_M, fs=12, color=WHITE):
    """圆形编号徽标。"""
    oval(s, x, y, d, d, fill=fill)
    return txt(s, x, y, d, d, label, fs=fs, color=color, bold=True, align="center")


def grad_bar(s, img_name, x, y, w, h, text, fs=15, color=WHITE, assets_dir=None):
    """渐变横幅图 + 居中白字（比纯色矩形更有层次）。"""
    pic(s, img_name, x, y, w, h, fit=False, assets_dir=assets_dir)
    return txt(s, x, y, w, h, text, fs=fs, color=color, bold=True, align="center")


def gray_box(s, x, y, w, h, rows, fs=14, numbered=True, fill=GRAY_BG, line=GRAY_LN):
    """浅灰说明框；rows = [[(文字,色,粗), ...], ...] 一行一个。"""
    rect(s, x, y, w, h, fill=fill, line=line, lw=0.9)
    outs = []
    rh = h / max(1, len(rows))
    for i, parts in enumerate(rows):
        ry = y + i * rh
        if numbered:
            num_circle(s, x + 14, ry + rh / 2 - 11, 22, str(i + 1), fill=(70, 122, 190), fs=12)
            txt(s, x + 44, ry + rh / 2 - 9, 18, 18, "\u25b6", fs=11, color=(60, 112, 182), align="left")
            outs.append(rich(s, x + 66, ry + 4, w - 82, rh - 8, parts, fs=fs, align="left"))
        else:
            outs.append(rich(s, x + 18, ry + 4, w - 30, rh - 8, parts, fs=fs, align="left"))
    return outs


def pill(s, x, y, w, h, text, fs=12.5, fill=RED, color=WHITE, rounded=True, radius=0.5):
    """胶囊标签。"""
    (rrect if rounded else rect)(s, x, y, w, h, fill=fill, line=None,
                                 radius=radius if rounded else None)
    return txt(s, x, y, w, h, text, fs=fs, color=color, bold=True, align="center")


def chip(s, x, y, w, h, parts, fs=12.5, fill=(249, 250, 252), line=(206, 214, 224)):
    """浅色圆角小卡片。"""
    rrect(s, x, y, w, h, fill=fill, line=line, lw=0.9)
    return rich(s, x + 8, y + 2, w - 16, h - 4, parts, fs=fs, align="center")


def header(s, sec, title, logo="", width=960, fs=None, color=BLUE,
           line_color=(96, 148, 210), accent_img="accent_bar.png", assets_dir=None,
           brand=None):
    """页面标题栏。两种几何：

    无品牌行：标题 y=12..56，分隔线 y=54/57（headerEndY≈56，内容从 80 起）
    有品牌行：标识+字标 y=14..40，标题 y=52..96，分隔线 y=96/99（headerEndY≈98，内容从 122 起）

    标题文本框给足高度并顶对齐 —— 此前 y=10 h=34 居中锚定，32pt 字形 bbox 顶到
    y=-2 被 verify 判画布溢出；顶对齐 + 44pt 框高后字形完整落在框内。
    """
    fs = fs or scale_size("h1", 32)
    has_brand = bool(brand)
    title_y = 52.0 if has_brand else 12.0
    rule_y = 96.0 if has_brand else 54.0

    if has_brand:
        mark = (brand or {}).get("mark")
        word = (brand or {}).get("wordmark", "")
        if mark:
            try:
                pic(s, mark, 48, 14, 26, 26, fit=True, assets_dir=assets_dir)
            except Exception:  # noqa: BLE001
                oval(s, 48, 14, 26, 26, fill=color)
        if word:
            txt(s, width - 48 - 200, 16, 200, 22, str(word), fs=13, color=color,
                bold=True, align="right", role="mono")

    # 左侧竖条 + 标题（顶对齐，框高 44 容纳 32pt 行高）
    rect(s, 48, title_y + 4, 6, 34, fill=color)
    parts = []
    if sec:
        parts.append((f"{sec}  ", color, True))
    parts.append((str(title), color, True))
    rich(s, 64, title_y, width - 64 - 200, 44, parts, fs=fs, align="left",
         anchor="top", role="title")

    hline(s, 48, rule_y, width - 96, line_color, 1.5)
    hline(s, 48, rule_y + 3, 120, color, 2.5)
    return rule_y


def cards(s, x, y, w, h, items, cols=3, fill=WHITE, line=GRAY_LN,
          title_fs=15, body_fs=11.5, gap=12, title_color=BLUE, body_color=GRAY_TX,
          accent=None):
    """卡片网格。items = [{"title":..,"body":..,"color":..}, ...] 或 [[title, body], ...]"""
    norm = []
    for it in items or []:
        if isinstance(it, dict):
            norm.append((str(it.get("title", "")), str(it.get("body", "")), it.get("color")))
        elif isinstance(it, (list, tuple)):
            t = str(it[0]) if len(it) > 0 else ""
            b = str(it[1]) if len(it) > 1 else ""
            c = it[2] if len(it) > 2 else None
            norm.append((t, b, c))
    n = len(norm)
    if n == 0:
        return []
    rows = (n + cols - 1) // cols
    cw = (w - gap * (cols - 1)) / cols
    chh = (h - gap * (rows - 1)) / rows
    out = []
    for i, (t, b, c) in enumerate(norm):
        r, cc = divmod(i, cols)
        cx = x + cc * (cw + gap)
        cy = y + r * (chh + gap)
        ac = parse_color(c) or parse_color(accent) or BLUE_M
        rrect(s, cx, cy, cw, chh, fill=fill, line=line, lw=0.9, radius=0.08)
        rect(s, cx, cy, 4, chh, fill=ac)  # 左侧色条
        txt(s, cx + 14, cy + 8, cw - 24, 22, t, fs=title_fs, color=ac, bold=True, align="left")
        if b:
            txt(s, cx + 14, cy + 32, cw - 24, chh - 40, b, fs=body_fs, color=body_color,
                align="left", anchor="top")
        out.append((cx, cy, cw, chh))
    return out


def stat_cards(s, x, y, w, h, items, fill=WHITE, line=GRAY_LN, gap=12,
               num_fs=30, label_fs=12, color=BLUE):
    """数字统计卡（1 行 N 列）。items = [{"value":"98%","label":"合格率"}, ...]

    大数字用 numeric 角色字体（Segoe UI/Consolas/Times），标签用 caption 角色 ——
    数字与正文不同族是「数据感」的来源（调研：numeric 单独成族）。
    """
    n = len(items or [])
    if n == 0:
        return []
    cw = (w - gap * (n - 1)) / n
    out = []
    for i, it in enumerate(items):
        if isinstance(it, dict):
            v, lb, c = str(it.get("value", "")), str(it.get("label", "")), it.get("color")
        else:
            v, lb, c = str(it[0]), (str(it[1]) if len(it) > 1 else ""), (it[2] if len(it) > 2 else None)
        cx = x + i * (cw + gap)
        ac = parse_color(c) or color
        rrect(s, cx, y, cw, h, fill=fill, line=line, lw=0.9, radius=0.1)
        txt(s, cx, y + h * 0.12, cw, h * 0.46, v, fs=num_fs, color=ac, bold=True,
            align="center", role="numeric")
        txt(s, cx, y + h * 0.62, cw, h * 0.30, lb, fs=label_fs, color=GRAY_TX,
            align="center", role="caption")
        out.append((cx, y, cw, h))
    return out


def timeline(s, x, y, w, h, items, color=BLUE_M, dot_d=18, fs=12, label_fs=11):
    """水平时间轴。items = [{"label":"2024","text":"..."}, ...]"""
    n = len(items or [])
    if n == 0:
        return []
    mid = y + h / 2.0
    hline(s, x, mid, w, color, 2.0)
    step = w / max(1, n - 1) if n > 1 else 0
    out = []
    for i, it in enumerate(items):
        if isinstance(it, dict):
            lb, tx = str(it.get("label", "")), str(it.get("text", ""))
        else:
            lb, tx = str(it[0]), (str(it[1]) if len(it) > 1 else "")
        cx = x + i * step if n > 1 else x + w / 2.0
        oval(s, cx - dot_d / 2.0, mid - dot_d / 2.0, dot_d, dot_d, fill=color)
        txt(s, cx - 60, mid - dot_d / 2.0 - 24, 120, 20, lb, fs=fs, color=color,
            bold=True, align="center")
        txt(s, cx - 70, mid + dot_d / 2.0 + 6, 140, h / 2.0, tx, fs=label_fs,
            color=GRAY_TX, align="center", anchor="top")
        out.append(cx)
    return out


def tagline_bar(s, x, y, w, h, text, fill=(238, 244, 252), color=BLUE, fs=14):
    """页面底部总结条。"""
    rrect(s, x, y, w, h, fill=fill, line=None, radius=0.25)
    return txt(s, x + 16, y, w - 32, h, text, fs=fs, color=color, bold=True, align="center")


def footer(s, width=960, height=540, left="", right="", page=None, total=None,
           color=GRAY_TX, rule=GRAY_LN, fs=10.5, y=None):
    """页脚：细分隔线 + 左 deck 名 + 右页码（调研：footer 10-12pt，source note 置底）。

    页脚是「设计过」与「拼凑出来」最便宜的分水岭 —— 全篇统一出现即显专业。
    """
    fy = float(y if y is not None else height - 26)
    hline(s, 48, fy - 6, width - 96, rule, 0.75)
    if left:
        txt(s, 48, fy, width * 0.55, 18, str(left), fs=fs, color=color,
            align="left", role="caption")
    pg = ""
    if page is not None:
        pg = f"{page:02d}" + (f" / {total:02d}" if total else "")
    elif right:
        pg = str(right)
    if pg:
        txt(s, width - 48 - 120, fy, 120, 18, pg, fs=fs, color=color,
            align="right", role="mono")
    return fy


def divider(s, width=960, height=540, sec="", title="", subtitle="",
            primary=BLUE, bg=None, text_color=None, accent=None):
    """章节分隔页（调研 Layout 8：title-only，44-64pt，≤5 词，给观众心理重置）。

    大编号 + 大标题 + 一句副题，其余留白 —— 分隔页信息越少越高级。
    """
    tc = parse_color(text_color) or (WHITE if bg else INK)
    ac = parse_color(accent) or primary
    if bg:
        rect(s, 0, 0, width, height, fill=bg)
    # 左侧超大章节号（半透明感用浅色实色近似）
    if sec:
        txt(s, 48, height * 0.20, 300, 120, str(sec), fs=96, color=ac,
            bold=True, align="left", role="numeric")
    rect(s, 52, height * 0.20 + 128, 64, 5, fill=ac)
    txt(s, 48, height * 0.20 + 152, width - 160, 70, str(title),
        fs=scale_size("display", 44), color=tc, bold=True, align="left", role="display")
    if subtitle:
        txt(s, 48, height * 0.20 + 232, width - 200, 40, str(subtitle),
            fs=scale_size("body", 18), color=tc, align="left", role="caption")
    return (48, height * 0.20, width - 96, height * 0.6)


# ---------------------------------------------------------------- 表格
def table_native(s, x, y, w, h, header_row, rows, cols=None,
                 header_fill=(32, 84, 160), header_color=WHITE, hdr_fs=13, cell_fs=11,
                 alt_fill=(238, 244, 252), line=GRAY_LN, align="center"):
    """原生 AddTable —— MS Office 上稳定，且生成的表格用户可在 PowerPoint 里正常编辑。

    rows: [[cell, cell, ...], ...]，cell 可为 str 或 {"text":..,"color":..,"bold":..,"align":..}
    """
    nrows = len(rows) + (1 if header_row else 0)
    ncols = len(header_row) if header_row else (len(rows[0]) if rows else 1)
    if nrows < 1 or ncols < 1:
        return None
    shp = s.Shapes.AddTable(nrows, ncols, x, y, w, h)
    tbl = shp.Table

    # 列宽
    if cols:
        total = float(sum(cols)) or 1.0
        for i, cw in enumerate(cols[:ncols], start=1):
            try:
                tbl.Columns(i).Width = w * (float(cw) / total)
            except Exception:  # noqa: BLE001
                pass

    def put(r, c, val, fs, color, bold, al):
        try:
            cell = tbl.Cell(r, c)
            tf = cell.Shape.TextFrame
            try:
                tf.WordWrap = O.msoTrue
                tf.MarginLeft = 4
                tf.MarginRight = 4
                tf.MarginTop = 2
                tf.MarginBottom = 2
                tf.VerticalAnchor = O.msoAnchorMiddle
            except Exception:  # noqa: BLE001
                pass
            tr = tf.TextRange
            tr.Text = "" if val is None else str(val)
            try:
                tr.Font.Size = float(fs)
                tr.Font.Name = FONT
                tr.Font.Color.RGB = h2b(parse_color(color) or INK)
                tr.Font.Bold = O.msoTrue if bold else 0
                tr.ParagraphFormat.Alignment = _ALIGN.get(al, O.ppAlignCenter)
            except Exception:  # noqa: BLE001
                pass
        except Exception:  # noqa: BLE001
            pass

    r = 1
    if header_row:
        for c, t in enumerate(header_row, start=1):
            put(r, c, t, hdr_fs, header_color, True, "center")
        try:
            for c in range(1, ncols + 1):
                tbl.Cell(r, c).Shape.Fill.ForeColor.RGB = h2b(parse_color(header_fill) or BLUE)
        except Exception:  # noqa: BLE001
            pass
        r += 1

    for i, row in enumerate(rows):
        for c in range(1, ncols + 1):
            val = row[c - 1] if c - 1 < len(row) else ""
            if isinstance(val, dict):
                put(r, c, val.get("text", ""), val.get("fs", cell_fs),
                    val.get("color", INK), bool(val.get("bold", False)),
                    val.get("align", align))
            else:
                put(r, c, val, cell_fs, INK, False, align)
            try:
                bg = alt_fill if i % 2 == 0 else WHITE
                tbl.Cell(r, c).Shape.Fill.ForeColor.RGB = h2b(parse_color(bg) or WHITE)
            except Exception:  # noqa: BLE001
                pass
        r += 1
    return shp


def table_rects(s, x, y, w, cols, header_row, rows, hdr_h=32, row_h=66,
                header_fill=(32, 84, 160), hdr_fs=14, cell_fs=11,
                alt_fill=(238, 244, 252), left_bold_color=BLUE):
    """矩形拼接表格 —— 上游为规避 WPS `AddTable` 边框 API 不稳定而写的兜底方案。

    cols        : 各列宽（pt）
    header_row  : [str, ...]
    rows        : [[第0列文本, [(文本,色,粗), ...]], ...]
    """
    rect(s, x, y, w, hdr_h, fill=header_fill)
    cx = x
    for i, t in enumerate(header_row):
        cw = cols[i] if i < len(cols) else (w / max(1, len(header_row)))
        txt(s, cx, y, cw, hdr_h, t, fs=hdr_fs, color=WHITE, bold=True, align="center")
        cx += cw
    ry = y + hdr_h
    for i, row in enumerate(rows):
        cap, body = (row[0], row[1]) if isinstance(row, (list, tuple)) and len(row) >= 2 else (str(row), [])
        rect(s, x, ry, w, row_h, fill=alt_fill if i % 2 == 0 else WHITE)
        hline(s, x, ry, w, (196, 208, 224), 0.9)
        c0 = cols[0] if cols else w * 0.25
        txt(s, x + 10, ry, c0, row_h, cap, fs=12, color=left_bold_color, bold=True, align="center")
        c1 = cols[1] if len(cols) > 1 else (w - c0)
        rich(s, x + c0 + 10, ry + 5, c1 - 20, row_h - 10, body, fs=cell_fs, align="left")
        ry += row_h
    hline(s, x, ry, w, (196, 208, 224), 0.9)


def table(s, x, y, w, h=None, header_row=None, rows=None, cols=None, mode="auto", **kw):
    """表格统一入口。mode: auto|native|rects。"""
    header_row = header_row or []
    rows = rows or []
    if mode == "auto":
        mode = "native"
    if mode == "native":
        hh = h if h else (32 + 26 * len(rows))
        return table_native(s, x, y, w, hh, header_row, rows, cols=cols, **{
            k: v for k, v in kw.items()
            if k in ("header_fill", "header_color", "hdr_fs", "cell_fs", "alt_fill", "line", "align")
        })
    return table_rects(s, x, y, w, cols or [w / max(1, len(header_row))] * len(header_row),
                       header_row, rows, **{
                           k: v for k, v in kw.items()
                           if k in ("hdr_h", "row_h", "header_fill", "hdr_fs", "cell_fs",
                                    "alt_fill", "left_bold_color")
                       })
