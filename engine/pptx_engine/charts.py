# -*- coding: utf-8 -*-
"""matplotlib 图表生成 —— 把上游工作流里「手工写 gen_charts.py」那一步自动化。

上游 SKILL.md 的 6 步生产流程中，第 3 步是「matplotlib 生成图表：柱状图/饼图/
折线图/横向柱/气泡图」，每个项目都要手写一个 gen_charts.py。本模块把它收敛成
一个 `make_chart(kind, data)` 入口，模型只需在 JSON 里声明数据。

中文字体坑（必须处理，否则图表里全是方块）：
    matplotlib 默认 DejaVu Sans 不含 CJK 字形。这里显式把
    `font.sans-serif` 指到微软雅黑/黑体，并关掉 `axes.unicode_minus`
    （否则负号渲染成方块）。
"""
from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
from typing import Optional

from .elements import parse_color

_DEFAULT_PALETTE = [
    "#1F4E8C", "#2C6EB5", "#4E9BD8", "#E8862A", "#3AA76D",
    "#9B59B6", "#E05C5C", "#16A085", "#F1C40F", "#7F8C8D",
]

_font_ready = False


def _setup_font() -> None:
    """配置 matplotlib 中文字体（幂等）。"""
    global _font_ready
    if _font_ready:
        return
    import matplotlib
    matplotlib.use("Agg")  # 无窗口后端 —— 引擎在后台进程里跑，绝不能弹图窗
    from matplotlib import font_manager, rcParams

    cands = ["Microsoft YaHei", "微软雅黑", "SimHei", "黑体", "DengXian", "等线", "SimSun"]
    installed = {f.name for f in font_manager.fontManager.ttflist}
    picked = [c for c in cands if c in installed]
    rcParams["font.sans-serif"] = (picked + ["DejaVu Sans"]) if picked else ["DejaVu Sans"]
    rcParams["axes.unicode_minus"] = False
    rcParams["font.size"] = 11
    _font_ready = True


def _colors(spec, n: int) -> list:
    """从 spec 取色，不足则用默认色板补齐。"""
    raw = spec.get("colors") or spec.get("palette") or []
    out = []
    for c in raw:
        rgb = parse_color(c)
        if rgb:
            out.append("#%02X%02X%02X" % rgb)
    while len(out) < n:
        out.append(_DEFAULT_PALETTE[len(out) % len(_DEFAULT_PALETTE)])
    return out[:max(n, 1)]


def _series(data: dict) -> list[dict]:
    """归一化 series：支持 [{name,values}] / [[...],[...]] / 直接 values。"""
    s = data.get("series")
    if isinstance(s, list) and s:
        out = []
        for i, it in enumerate(s):
            if isinstance(it, dict):
                out.append({"name": str(it.get("name", f"S{i + 1}")),
                            "values": list(it.get("values") or it.get("data") or [])})
            else:
                out.append({"name": f"S{i + 1}", "values": list(it)})
        return out
    v = data.get("values") or data.get("data")
    if isinstance(v, list):
        return [{"name": str(data.get("name", "数值")), "values": list(v)}]
    return [{"name": "数值", "values": []}]


def _labels(data: dict, n: int) -> list:
    lb = data.get("labels") or data.get("x") or data.get("categories")
    if isinstance(lb, list) and lb:
        return [str(x) for x in lb]
    return [str(i + 1) for i in range(n)]


def _finalize(fig, out_path: str, dpi: int, transparent: bool) -> str:
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    fig.savefig(out_path, dpi=dpi, bbox_inches="tight",
                transparent=transparent, facecolor=fig.get_facecolor())
    import matplotlib.pyplot as plt
    plt.close(fig)
    return os.path.abspath(out_path)


def _style_ax(ax, spec: dict, grid: bool = True) -> None:
    from matplotlib import rcParams  # noqa: F401
    if spec.get("xlabel"):
        ax.set_xlabel(str(spec["xlabel"]), fontsize=spec.get("labelFs", 11))
    if spec.get("ylabel"):
        ax.set_ylabel(str(spec["ylabel"]), fontsize=spec.get("labelFs", 11))
    if spec.get("title"):
        ax.set_title(str(spec["title"]), fontsize=spec.get("titleFs", 14),
                     fontweight="bold", pad=10)
    if grid:
        ax.grid(True, axis=spec.get("gridAxis", "y"), linestyle="--",
                linewidth=0.6, alpha=0.45)
        ax.set_axisbelow(True)
    for sp in ("top", "right"):
        try:
            ax.spines[sp].set_visible(False)
        except Exception:  # noqa: BLE001
            pass
    if spec.get("ylim"):
        ax.set_ylim(spec["ylim"])
    if spec.get("xlim"):
        ax.set_xlim(spec["xlim"])
    if spec.get("valueLabels", True):
        _annotate(ax, spec)


def _annotate(ax, spec: dict) -> None:
    """在柱/点上标数值（汇报 PPT 常见需求）。"""
    fmt = spec.get("valueFormat", "{:.0f}")
    try:
        for cont in ax.containers:
            ax.bar_label(cont, fmt=fmt, fontsize=spec.get("valueFs", 9), padding=2)
    except Exception:  # noqa: BLE001
        pass


# ---------------------------------------------------------------- 各图种
def _bar(data, spec, figsize, dpi, out, transparent):
    import matplotlib.pyplot as plt
    import numpy as np

    ser = _series(data)
    labels = _labels(data, max(len(s["values"]) for s in ser) if ser else 0)
    colors = _colors(spec, len(ser) if len(ser) > 1 else len(labels))
    fig, ax = plt.subplots(figsize=figsize)
    x = np.arange(len(labels))
    if spec.get("stacked"):
        bottom = np.zeros(len(labels))
        for i, s in enumerate(ser):
            v = np.array(s["values"][:len(labels)], dtype=float)
            ax.bar(x, v, bottom=bottom, label=s["name"],
                   color=_colors(spec, len(ser))[i], width=0.62)
            bottom += v
        ax.legend(frameon=False, fontsize=spec.get("legendFs", 10))
    elif len(ser) > 1:
        w = 0.8 / len(ser)
        for i, s in enumerate(ser):
            v = np.array(s["values"][:len(labels)], dtype=float)
            ax.bar(x + (i - (len(ser) - 1) / 2) * w, v, w, label=s["name"], color=colors[i])
        ax.legend(frameon=False, fontsize=spec.get("legendFs", 10))
    else:
        v = np.array(ser[0]["values"][:len(labels)], dtype=float) if ser else np.array([])
        ax.bar(x, v, color=colors[:len(labels)], width=0.62)
    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=spec.get("tickFs", 10))
    _style_ax(ax, spec)
    return _finalize(fig, out, dpi, transparent)


def _hbar(data, spec, figsize, dpi, out, transparent):
    import matplotlib.pyplot as plt
    import numpy as np

    ser = _series(data)
    labels = _labels(data, len(ser[0]["values"]) if ser else 0)
    colors = _colors(spec, len(labels))
    fig, ax = plt.subplots(figsize=figsize)
    y = np.arange(len(labels))
    v = np.array(ser[0]["values"][:len(labels)], dtype=float) if ser else np.array([])
    ax.barh(y, v, color=colors[:len(labels)], height=0.62)
    ax.set_yticks(y)
    ax.set_yticklabels(labels, fontsize=spec.get("tickFs", 10))
    ax.invert_yaxis()
    if spec.get("valueLabels", True):
        try:
            ax.bar_label(ax.containers[0], fmt=spec.get("valueFormat", "{:.0f}"),
                         fontsize=spec.get("valueFs", 9), padding=3)
        except Exception:  # noqa: BLE001
            pass
    _style_ax(ax, spec, grid=True)
    ax.grid(True, axis="x", linestyle="--", linewidth=0.6, alpha=0.45)
    return _finalize(fig, out, dpi, transparent)


def _catmull_rom(x, y, per: int = 12):
    """Catmull-Rom 平滑（SPC 折线签名）。控制点 y 不做外推夹紧，
    但端点重复一次避免首尾过冲。点数 = 原点数*per。"""
    import numpy as np

    x = np.asarray(x, dtype=float)
    y = np.asarray(y, dtype=float)
    if len(x) < 3:
        return x, y
    xs = np.concatenate(([x[0]], x, [x[-1]]))
    ys = np.concatenate(([y[0]], y, [y[-1]]))
    out_x, out_y = [], []
    for i in range(1, len(xs) - 2):
        p0x, p1x, p2x, p3x = xs[i - 1], xs[i], xs[i + 1], xs[i + 2]
        p0y, p1y, p2y, p3y = ys[i - 1], ys[i], ys[i + 1], ys[i + 2]
        t = np.linspace(0, 1, per, endpoint=False)
        mx = 0.5 * ((2 * p1x) + (-p0x + p2x) * t
                    + (2 * p0x - 5 * p1x + 4 * p2x - p3x) * t ** 2
                    + (-p0x + 3 * p1x - 3 * p2x + p3x) * t ** 3)
        my = 0.5 * ((2 * p1y) + (-p0y + p2y) * t
                    + (2 * p0y - 5 * p1y + 4 * p2y - p3y) * t ** 2
                    + (-p0y + 3 * p1y - 3 * p2y + p3y) * t ** 3)
        out_x.extend(mx.tolist())
        out_y.extend(my.tolist())
    out_x.append(x[-1])
    out_y.append(y[-1])
    return np.array(out_x), np.array(out_y)


def _line(data, spec, figsize, dpi, out, transparent):
    import matplotlib.pyplot as plt
    import numpy as np

    ser = _series(data)
    labels = _labels(data, max((len(s["values"]) for s in ser), default=0))
    colors = _colors(spec, len(ser))
    style = dict(spec.get("chartStyle") or {})
    # 主题级默认（品牌主题 = 平滑+光晕+面积），spec 可覆盖
    for k in ("smooth", "glow", "area"):
        style.setdefault(k, spec.get(k))
    fig, ax = plt.subplots(figsize=figsize)
    x = np.arange(len(labels))
    for i, s in enumerate(ser):
        v = np.array(s["values"][:len(labels)], dtype=float)
        px, py = x, v
        if style.get("smooth") and len(v) >= 3:
            px, py = _catmull_rom(x, v)
        # 光晕：同色加宽半透明垫底（SPC 的 glow 层）
        if style.get("glow"):
            ax.plot(px, py, linewidth=spec.get("linewidth", 2.2) * 3.2,
                    color=colors[i], alpha=0.14, solid_capstyle="round")
        # 面积只给主序列（i==0）：控制限/对照线填面积会糊成一大片灰带
        # （2026-10-29 实测 品牌主题折线图翻车点）；需要全填用 areaAll=true
        if style.get("area") and (i == 0 or spec.get("areaAll")):
            ax.fill_between(px, py, alpha=0.13, color=colors[i], linewidth=0)
        ax.plot(px, py, marker=spec.get("marker", "o" if not style.get("smooth") else ""),
                markersize=spec.get("markerSize", 5),
                linewidth=spec.get("linewidth", 2.2), color=colors[i],
                label=s["name"], solid_capstyle="round")
        # 原始数据描点（平滑后仍标真值位置）
        if style.get("smooth"):
            ax.scatter(x, v, s=spec.get("markerSize", 5) ** 1.6, color=colors[i],
                       edgecolors="white", linewidths=0.9, zorder=5)
    if len(ser) > 1 or spec.get("legend"):
        ax.legend(frameon=False, fontsize=spec.get("legendFs", 10))
    step = max(1, len(labels) // 12)
    ax.set_xticks(x[::step])
    ax.set_xticklabels([labels[i] for i in range(0, len(labels), step)],
                       fontsize=spec.get("tickFs", 10), rotation=spec.get("rotate", 0))
    _style_ax(ax, spec)
    return _finalize(fig, out, dpi, transparent)


def _pie(data, spec, figsize, dpi, out, transparent):
    import matplotlib.pyplot as plt

    ser = _series(data)
    vals = [float(v) for v in (ser[0]["values"] if ser else [])]
    labels = _labels(data, len(vals))
    colors = _colors(spec, len(vals))
    fig, ax = plt.subplots(figsize=figsize)
    donut = spec.get("donut") or spec.get("chart") == "donut"

    # 小切片标签重叠是 matplotlib 饼图的经典坑：任一扇区占比 < 阈值时，
    # 外侧标签和百分比会互相压字。此时改用图例承载名称，百分比只标大扇区。
    total = sum(vals) or 1.0
    min_ratio = min(vals) / total if vals else 1.0
    small = min_ratio < float(spec.get("smallSliceThreshold", 0.06))

    if small:
        kw = {
            "colors": colors[:len(vals)],
            "autopct": lambda p: (f"{p:.1f}%" if p >= 6 else ""),
            "startangle": spec.get("startAngle", 90),
            "pctdistance": 0.78 if donut else 0.7,
            "wedgeprops": {"width": 0.42, "edgecolor": "w"} if donut else {"edgecolor": "w"},
            "textprops": {"fontsize": spec.get("tickFs", 10), "color": "white",
                          "fontweight": "bold"},
        }
        # 注意：带 autopct 时 ax.pie 返回 (wedges, texts, autotexts) 三元组，
        # 不带时返回二元组 —— 统一取第 0 项，别按固定长度解包。
        wedges = ax.pie(vals, **kw)[0]
        ax.legend(wedges, [f"{lb}  {v:g}（{v / total:.1%}）" for lb, v in zip(labels, vals)],
                  loc="center left", bbox_to_anchor=(1.0, 0.5), frameon=False,
                  fontsize=spec.get("legendFs", 10))
    else:
        kw = {
            "labels": labels if spec.get("showLabels", True) else None,
            "colors": colors[:len(vals)],
            "autopct": spec.get("autopct", "%1.1f%%"),
            "startangle": spec.get("startAngle", 90),
            "pctdistance": 0.72 if donut else 0.62,
            "wedgeprops": {"width": 0.42, "edgecolor": "w"} if donut else {"edgecolor": "w"},
            "textprops": {"fontsize": spec.get("tickFs", 10)},
        }
        ax.pie(vals, **kw)

    if donut and spec.get("centerText"):
        ax.text(0, 0, str(spec["centerText"]), ha="center", va="center",
                fontsize=spec.get("centerFs", 16), fontweight="bold",
                color=spec.get("centerColor", "#1F4E8C"))
    ax.set_aspect("equal")
    if spec.get("title"):
        ax.set_title(str(spec["title"]), fontsize=spec.get("titleFs", 14), fontweight="bold")
    return _finalize(fig, out, dpi, transparent)


def _scatter(data, spec, figsize, dpi, out, transparent):
    import matplotlib.pyplot as plt
    import numpy as np

    xs = data.get("x") or []
    ys = data.get("y") or []
    if not xs or not ys:
        ser = _series(data)
        ys = ser[0]["values"] if ser else []
        xs = list(range(len(ys)))
    sizes = data.get("sizes") or [spec.get("size", 60)] * len(xs)
    colors = _colors(spec, 1)
    fig, ax = plt.subplots(figsize=figsize)
    ax.scatter(np.array(xs, dtype=float), np.array(ys, dtype=float),
               s=np.array(sizes, dtype=float), c=colors[0], alpha=spec.get("alpha", 0.72),
               edgecolors="white", linewidths=0.8)
    labels = data.get("labels")
    if labels:
        for i, lb in enumerate(labels[:len(xs)]):
            ax.annotate(str(lb), (float(xs[i]), float(ys[i])),
                        fontsize=spec.get("valueFs", 9), xytext=(4, 4),
                        textcoords="offset points")
    _style_ax(ax, spec, grid=True)
    return _finalize(fig, out, dpi, transparent)


_KINDS = {
    "bar": _bar, "column": _bar, "col": _bar,
    "hbar": _hbar, "horizontal_bar": _hbar, "barh": _hbar,
    "line": _line, "area": _line, "trend": _line,
    "pie": _pie, "donut": _pie, "doughnut": _pie,
    "scatter": _scatter, "bubble": _scatter,
}


def chart_kinds() -> list[str]:
    return sorted(set(_KINDS.keys()))


def make_chart(kind: str, data: dict, out_dir=None, title: str = "",
               dpi: int = 160, theme: Optional[dict] = None,
               figsize=None, transparent: bool = False,
               out_name: Optional[str] = None) -> str:
    """生成图表 PNG，返回绝对路径。

    Args:
        kind: bar|hbar|line|area|pie|donut|scatter|bubble
        data: {"labels":[...], "series":[{"name":..,"values":[..]}], ...}
        out_dir: 输出目录，默认 <cwd>/.charts
        figsize: (w,h) 英寸，默认按画布比例 4.6x2.9
    """
    _setup_font()
    kind = str(kind or "bar").lower()
    fn = _KINDS.get(kind)
    if fn is None:
        raise ValueError(f"未知图表类型：{kind}。可用：{', '.join(chart_kinds())}")

    spec = dict(data or {})
    if theme:
        # 分类色优先级：显式 colors > 主题 chartPalette（品牌色系）> Okabe-Ito > 主题 palette
        if not spec.get("colors"):
            pal = (theme.get("typo") or {}).get("chartPalette")
            if pal:
                spec["colors"] = list(pal)
            elif theme.get("palette"):
                spec["colors"] = list(theme["palette"])
        # 折线样式签名（品牌主题 = 平滑+光晕+面积），spec 可逐项覆盖
        cs = (theme.get("typo") or {}).get("chartStyle") or {}
        if cs:
            spec["chartStyle"] = {**cs, **(spec.get("chartStyle") or {})}
        # 环心大数字颜色跟主题（品牌主题=深绿，default=藏青），不再硬编码 #1F4E8C
        if not spec.get("centerColor"):
            spec["centerColor"] = ((theme.get("semantic") or {}).get("deep")
                                   or theme.get("primary") or "#1F4E8C")
    if title and not spec.get("title"):
        spec["title"] = title
    if figsize is None:
        figsize = tuple(spec.get("figsize") or (4.6, 2.9))

    base = Path(out_dir).expanduser().resolve() if out_dir else Path.cwd() / ".charts"
    base.mkdir(parents=True, exist_ok=True)
    if out_name:
        name = str(out_name)
        if not name.lower().endswith(".png"):
            name += ".png"
    else:
        digest = hashlib.sha1(json.dumps({"k": kind, "d": data, "t": title, "f": figsize},
                                         ensure_ascii=False, sort_keys=True,
                                         default=str).encode("utf-8")).hexdigest()[:12]
        name = f"chart_{kind}_{digest}.png"
    out = str(base / name)
    return fn(spec, spec, tuple(figsize), int(dpi), out, bool(transparent))
