# -*- coding: utf-8 -*-
"""素材生成器 —— PIL 程序化绘制，不依赖外部图床。

直接沿用上游 harness-anything/pptengine/assets.py 的绘制思路（渐变斜切横幅、
竖条、圆环、箭头、页面背景、装饰圆），并补两点：

1. **主题色感知**：`ensure_theme_assets(theme)` 按主题主色生成一整套横幅/竖条，
   换主题不用改代码。
2. **中文字体探测**：上游把字体名硬编码成「微软雅黑」，这里加一层
   Windows 字体文件探测，找不到时退回 PIL 默认字体而不是崩掉。
"""
from __future__ import annotations

import os
import random
from pathlib import Path
from typing import Optional, Sequence

from PIL import Image, ImageDraw, ImageFont


def lerp(a: Sequence[int], b: Sequence[int], t: float):
    return tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(3))


# ---------------------------------------------------------------- 字体
_FONT_CANDIDATES = [
    r"C:\Windows\Fonts\msyh.ttc",      # 微软雅黑
    r"C:\Windows\Fonts\msyhbd.ttc",
    r"C:\Windows\Fonts\simhei.ttf",    # 黑体
    r"C:\Windows\Fonts\simsun.ttc",    # 宋体
    r"C:\Windows\Fonts\Deng.ttf",      # 等线
]


def load_font(size: int = 20, bold: bool = False):
    for p in _FONT_CANDIDATES:
        if os.path.exists(p):
            try:
                return ImageFont.truetype(p, size)
            except Exception:  # noqa: BLE001
                continue
    try:
        return ImageFont.load_default(size)
    except TypeError:
        return ImageFont.load_default()


# ---------------------------------------------------------------- 基础图形
def grad_bar(path, c1, c2, w=480, h=72, cut=0.22, dark_bottom=True):
    """渐变斜切横幅。"""
    im = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(im, "RGBA")
    for x in range(w):
        c = lerp(c1, c2, x / max(1, w - 1))
        d.line([(x, 0), (x, h)], fill=c + (255,))
    if dark_bottom:
        y0 = int(h * 0.78)
        for y in range(y0, h):
            t = (y - y0) / max(1, h - 1 - y0)
            for x in range(w):
                base = lerp(c1, c2, x / max(1, w - 1))
                c = lerp(base, tuple(int(v * 0.72) for v in base), t)
                d.point((x, y), fill=c + (255,))
    ch = int(h * cut)
    d.polygon([(0, 0), (ch, 0), (0, h)], fill=(0, 0, 0, 0))
    d.polygon([(w, h), (w - ch, h), (w, 0)], fill=(0, 0, 0, 0))
    im.save(path)


def vertical_accent(path, c1=(58, 116, 190), c2=(24, 62, 140), w=40, h=120):
    im = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(im, "RGBA")
    for y in range(h):
        d.line([(0, y), (w, y)], fill=lerp(c1, c2, y / max(1, h - 1)) + (255,))
    im.save(path)


def ring(path, d0=300, stroke=(150, 198, 116), fill=(255, 255, 255, 0), width=3.0):
    im = Image.new("RGBA", (d0, d0), (0, 0, 0, 0))
    dr = ImageDraw.Draw(im, "RGBA")
    dr.ellipse([0, 0, d0 - 1, d0 - 1], fill=fill, outline=stroke, width=int(width))
    im.save(path)


def mark_hexagon(path, size=96, color=(100, 160, 40), stroke=9, gap_side="right"):
    """开口六边形标识 —— 按用户提供的 品牌模板参考图左上角 mark 程序化复刻。

    画一个正六边形环，去掉一侧的边形成「C」形开口，线端做圆头。
    纯矢量线段绘制，任意尺寸清晰，不依赖外部 logo 文件。
    """
    import math
    pad = stroke + 4
    S = size + pad * 2
    im = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(im, "RGBA")
    cx = cy = S / 2.0
    r = size / 2.0
    # 顶点：平顶六边形（参考图形态）
    verts = []
    for i in range(6):
        ang = math.radians(60 * i + 0)
        verts.append((cx + r * math.cos(ang), cy + r * math.sin(ang)))
    # 边索引 → 方位（平顶六边形：边0右、边1右下、…）；去掉 gap 侧的那条边
    skip = {"right": 0, "left": 3, "top": 4, "bottom": 1}.get(gap_side, 0)
    for i in range(6):
        if i == skip:
            continue
        a, b = verts[i], verts[(i + 1) % 6]
        d.line([a, b], fill=tuple(color) + (255,), width=stroke)
        # 圆头端点
        for p in (a, b):
            d.ellipse([p[0] - stroke / 2, p[1] - stroke / 2,
                       p[0] + stroke / 2, p[1] + stroke / 2], fill=tuple(color) + (255,))
    im.save(path)


def arrow_img(path, w, h, direction="right", c1=(108, 190, 78), c2=(58, 148, 44), head=0.16):
    im = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(im, "RGBA")
    for x in range(w):
        d.line([(x, 0), (x, h)], fill=lerp(c1, c2, x / max(1, w - 1)) + (255,))
    hw, bh = int(w * head), h // 2
    if direction == "right":
        d.polygon([(w - 1, bh - h // 2), (w - 1, bh + h // 2), (w - hw, bh)], fill=c2 + (255,))
    else:
        d.polygon([(0, bh - h // 2), (0, bh + h // 2), (hw, bh)], fill=c2 + (255,))
    im.save(path)


def page_bg(path, w=960, h=540, top=(252, 253, 254), bottom=(246, 248, 250)):
    im = Image.new("RGB", (w, h))
    d = ImageDraw.Draw(im)
    for y in range(h):
        d.line([(0, y), (w, y)], fill=lerp(top, bottom, y / max(1, h - 1)))
    im.save(path)


def deco_circle(path, d=200, outer=(210, 228, 246), inner=(228, 240, 252), alpha=90):
    im = Image.new("RGBA", (d, d), (0, 0, 0, 0))
    dr = ImageDraw.Draw(im, "RGBA")
    dr.ellipse([0, 0, d - 1, d - 1], fill=outer + (alpha,))
    m = int(d * 0.12)
    dr.ellipse([m, m, d - 1 - m, d - 1 - m], fill=inner + (alpha,))
    im.save(path)


def cover_bg(path, w=960, h=540, c1=(24, 62, 140), c2=(16, 44, 110), band=True):
    """封面底图：对角渐变 + 可选装饰带。上游是每校一张手工 template_bg.png，
    这里改成程序化生成，换主题色即可复用。"""
    im = Image.new("RGB", (w, h))
    d = ImageDraw.Draw(im)
    for y in range(h):
        for x in range(0, w, 4):
            t = (x / max(1, w - 1) * 0.55 + y / max(1, h - 1) * 0.45)
            c = lerp(c1, c2, t)
            d.rectangle([x, y, x + 3, y], fill=c)
    if band:
        d.rectangle([0, int(h * 0.62), w, int(h * 0.628)], fill=(255, 255, 255))
    im.save(path)


def texture_bg(path, w=960, h=540, kind="grid", base=(252, 253, 254),
               ink=(16, 62, 140), step=48, alpha=12):
    """内容页背景纹理：极淡的网格/点阵/横线，给页面一层「设计底」。

    依据：纯白底 + 黑字是上游点名的「instant boredom」反模式；
    但背景又绝不能抢内容（学术规范：不要与内容无关的花哨背景图）。
    解法就是这种 alpha≈12/255 的结构化淡纹 —— 有质感、不干扰。
    """
    im = Image.new("RGB", (w, h), tuple(base))
    d = ImageDraw.Draw(im, "RGBA")
    kind = str(kind or "grid").lower()
    if kind == "grid":
        for x in range(step, w, step):
            d.line([(x, 0), (x, h)], fill=tuple(ink) + (alpha,))
        for y in range(step, h, step):
            d.line([(0, y), (w, y)], fill=tuple(ink) + (alpha,))
    elif kind == "dots":
        for y in range(step // 2, h, step):
            for x in range(step // 2, w, step):
                d.ellipse([x - 1, y - 1, x + 1, y + 1], fill=tuple(ink) + (alpha + 10,))
    elif kind == "rule":
        for y in range(step, h, step):
            d.line([(0, y), (w, y)], fill=tuple(ink) + (alpha,))
    elif kind == "diag":
        for x in range(-h, w, step):
            d.line([(x, 0), (x + h, h)], fill=tuple(ink) + (alpha,))
    # 顶部一条极淡的主题色渐隐带，增加层次但不抢戏
    for y in range(0, 6):
        t = y / 6.0
        d.line([(0, y), (w, y)], fill=lerp(tuple(ink), tuple(base), t) + (26,))
    im.save(path)


# ---------------------------------------------------------------- 科研示意图
def fig_heat(path, w, h, palette, cell=10, seed=None):
    rng = random.Random(seed)
    im = Image.new("RGB", (w, h), (255, 255, 255))
    d = ImageDraw.Draw(im)
    cols = [lerp(palette[0], palette[1], i / max(1, len(palette) - 1)) for i in range(len(palette))]
    for y in range(0, h, cell):
        for x in range(0, w, cell):
            d.rectangle([x, y, x + cell - 1, y + cell - 1], fill=rng.choice(cols))
    im.save(path)


def fig_scatter(path, w, h, colors, n=240, seed=None):
    rng = random.Random(seed)
    im = Image.new("RGB", (w, h), (255, 255, 255))
    d = ImageDraw.Draw(im)
    for _ in range(n):
        cx, cy = rng.uniform(10, w - 10), rng.uniform(10, h - 10)
        c = rng.choice(colors)
        r = rng.uniform(2, 6)
        d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=c)
    im.save(path)


# ---------------------------------------------------------------- 一键生成
def ensure_default_assets(target_dir=None, force=False) -> Path:
    """生成通用素材（已存在则跳过，force=True 强制重画）。"""
    base = Path(target_dir).expanduser().resolve() if target_dir else Path.cwd() / ".ppt-assets"
    base.mkdir(parents=True, exist_ok=True)

    def gen(name, fn):
        p = base / name
        if force or not p.exists():
            fn(str(p))

    gen("bg_page.png", lambda p: page_bg(p))
    gen("bar_blue.png", lambda p: grad_bar(p, (58, 116, 190), (24, 62, 140)))
    gen("bar_teal.png", lambda p: grad_bar(p, (36, 150, 160), (18, 96, 118)))
    gen("bar_red.png", lambda p: grad_bar(p, (232, 78, 62), (176, 24, 24)))
    gen("bar_dark.png", lambda p: grad_bar(p, (40, 88, 160), (16, 44, 110)))
    gen("bar_green.png", lambda p: grad_bar(p, (72, 168, 116), (24, 110, 70)))
    gen("accent_bar.png", lambda p: vertical_accent(p))
    gen("deco_circle.png", lambda p: deco_circle(p))
    gen("ring.png", lambda p: ring(p))
    gen("arrow_green.png", lambda p: arrow_img(p, 690, 34, "right"))
    gen("arrow_orange.png", lambda p: arrow_img(p, 420, 30, "left", (244, 158, 62), (226, 116, 32)))
    gen("fig_heat.png", lambda p: fig_heat(p, 360, 200, [(240, 200, 190), (180, 60, 50), (90, 20, 20)], 9, seed=1))
    gen("fig_scatter.png", lambda p: fig_scatter(p, 360, 200, [(80, 120, 200), (200, 120, 160)], seed=2))
    return base


def ensure_theme_assets(theme: dict, target_dir=None, force=False) -> Path:
    """按主题色生成一整套素材。theme 需含 primary/secondary（[r,g,b] 或 #hex）。"""
    from .elements import parse_color
    base = Path(target_dir).expanduser().resolve() if target_dir else Path.cwd() / ".ppt-assets"
    base.mkdir(parents=True, exist_ok=True)
    tid = str(theme.get("id", "custom"))
    p1 = parse_color(theme.get("primary")) or (24, 62, 140)
    p2 = parse_color(theme.get("secondary")) or tuple(max(0, int(v * 0.6)) for v in p1)
    acc = parse_color(theme.get("accent")) or (232, 126, 34)

    def gen(name, fn):
        p = base / name
        if force or not p.exists():
            fn(str(p))

    gen(f"bar_{tid}.png", lambda p: grad_bar(p, p1, p2))
    gen(f"accent_{tid}.png", lambda p: vertical_accent(p, p1, p2))
    gen(f"cover_{tid}.png", lambda p: cover_bg(p, c1=p1, c2=p2))
    gen(f"arrow_{tid}.png", lambda p: arrow_img(p, 690, 34, "right", acc, tuple(max(0, int(v * 0.7)) for v in acc)))
    gen(f"ring_{tid}.png", lambda p: ring(p, stroke=p1))
    # 品牌标识（如品牌主题的开口六边形）：主题带 brand.mark 时程序化生成
    brand = theme.get("brand") or {}
    if brand.get("mark"):
        mk = brand["mark"]
        gen(mk, lambda p: mark_hexagon(p, size=96, color=p1, stroke=10))
        # 深底版本（封面/分隔页上用浅色标识）
        gen(mk.replace(".png", "_light.png"),
            lambda p: mark_hexagon(p, size=96, color=(255, 255, 255), stroke=10))
    # 内容页背景纹理：暗色主题用深底+亮纹，浅色主题用浅底+主题色淡纹
    dark = bool((theme.get("rules") or {}).get("darkMode"))
    base_rgb = parse_color(theme.get("bg")) or ((15, 20, 35) if dark else (252, 253, 254))
    ink_rgb = p1 if not dark else parse_color(theme.get("primary")) or p1
    gen(f"bg_{tid}.png", lambda p: texture_bg(
        p, kind=(theme.get("typo") or {}).get("texture", "grid"),
        base=base_rgb, ink=ink_rgb, alpha=26 if dark else 12))
    return base


def list_assets(target_dir=None) -> list[dict]:
    base = Path(target_dir).expanduser().resolve() if target_dir else Path.cwd() / ".ppt-assets"
    if not base.exists():
        return []
    out = []
    for f in sorted(base.iterdir()):
        if f.suffix.lower() in (".png", ".jpg", ".jpeg"):
            st = f.stat()
            out.append({"name": f.name, "path": str(f), "size": st.st_size})
    return out
