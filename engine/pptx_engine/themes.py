# -*- coding: utf-8 -*-
"""设计主题 —— 抽取自上游 cli_anything/wps/styles/design_presets.py 的 4 套预设，
并补 3 套（政务红、医疗绿、暗色科技），共 7 套。

上游把配色/字体/间距/规则四组信息封在 `DesignPreset` 类里，本实现改成纯 dict，
便于直接 JSON 序列化给模型看，也便于用户在 config 里覆盖。

每套主题额外带 `layout` 段 —— 这是上游 SKILL.md「关键约束 ⚠️」那一段的机器可读版：
标题栏高度、内容起始 y、画布安全边距、正文字号下限。quality.py 会拿它做自动审查。
"""
from __future__ import annotations

from typing import Optional

# 通用布局约束（上游 SKILL.md 的硬规则 + 调研修正，所有主题共享）
# 字号依据（slidesmate / k-dense scientific-slides / 清华《PPT的设计与制作》）：
#   封面标题 44-54pt · 页标题 32-40pt · 正文 24-28pt（最低 18pt）
#   注释 14-16pt · 页脚 10-12pt · 行距 1.2-1.5x · 单行 ≤50-65 字符
BASE_LAYOUT = {
    "canvas": {"w": 960, "h": 540},
    "safeBottom": 518,        # 所有元素必须满足 y + h ≤ 518
    "headerEndY": 56,         # 标题栏（含分隔线）结束于 y≈56
    "contentStartY": 80,      # 第一个内容元素必须从 y ≥ 80 开始（≥24pt 间距）
    "minGap": 24,             # 标题与内容最小间距
    "margin": 48,             # 页边距 ≈5% of 960（调研：margins 5-10% each side）
    "titleMaxChars": 12,      # 页标题≤12字；章节分隔页≤5词
    "bodyMinFs": 16,          # 正文下限（投影建议≥18；屏幕阅读型看板放宽到16）
    "bodyRecommendFs": 20,
    "titleFs": 32,
    "coverTitleFs": 46,
    "footerFs": 10.5,
    "lineHeight": 1.35,       # 行距 1.2-1.5x
    "maxLineChars": 42,       # 中文按字宽计，≈半幅页宽（调研：50-65 字符/行）
    "maxBullets": 6,          # 6×6 规则上限；推荐 3-4 条
}

THEMES: dict[str, dict] = {
    # ---------------------------------------------------------------- 默认：干净
    # 依据：slidesbrain「corporate trust」60-30-10（soft white / navy / orange）、
    # WCAG 近黑正文 #121212、slidesmate 两字体系统、k-dense 留白 40-50%。
    # 取向：干净、克制、细节到位（页脚/细线/统一边距），不堆装饰。
    "default": {
        "id": "default",
        "label": "默认 · 干净商务",
        "desc": "近白底 + 藏青 30% + 橙 10% 强调；雅黑标题 + 等线正文；细节靠页脚与细线",
        "primary": "#1F3A5F",
        "secondary": "#2B4A73",
        "accent": "#F28C28",
        "dark": "#121212",
        "light": "#E8EDF3",
        "bg": "#F7F9FC",
        "text": "#121212",
        "muted": "#5A6472",
        "palette": ["#1F3A5F", "#2B4A73", "#4E79A7", "#F28C28", "#009E73", "#CC79A7"],
        "rules": {"visualRatio": 0.55, "maxColors": 4, "colorblindSafe": True,
                  "oneIdeaPerSlide": True, "whiteSpace": 0.45, "maxBullets": 5,
                  "darkMode": False, "contentDensity": "medium"},
    },
}


def theme_names() -> list[str]:
    return list(THEMES.keys())


# ================================================================
# 排版系统（2026-10-28 审美补强）
# ----------------------------------------------------------------
# 此前 themes 里的 fontSizes 只是「写着好看」，渲染器根本没读，
# 全篇单一微软雅黑 —— 这是「工程整洁但不像设计过」的主因之一。
# 这里给每套主题配真正的：
#   fonts  —— 字体搭配（display 封面大标题 / title 章节标题 / body 正文 /
#              caption 注释 / numeric 大数字 / mono 代码与编号）
#   scale  —— 字号层级（pt），渲染器按角色取用，不再各处硬编码
#   transition —— 默认整页切换效果（克制为主，不耍杂耍）
#   texture —— 背景纹理类型（PIL 程序化，见 assets.texture_bg）
# 字体名一律给「候选列表」，按本机 C:\Windows\Fonts 实测取第一个存在的，
# 避免 PowerPoint 静默回退成宋体。
# ================================================================

_FONT_FILES = {
    "Microsoft YaHei": "msyh.ttc",
    "Microsoft YaHei Light": "msyhl.ttc",
    "微软雅黑": "msyh.ttc",
    "DengXian": "Deng.ttf",
    "DengXian Light": "Dengl.ttf",
    "等线": "Deng.ttf",
    "SimHei": "simhei.ttf",
    "黑体": "simhei.ttf",
    "SimSun": "simsun.ttc",
    "宋体": "simsun.ttc",
    "KaiTi": "simkai.ttf",
    "楷体": "simkai.ttf",
    "FangSong": "simfang.ttf",
    "仿宋": "simfang.ttf",
    "Segoe UI": "segoeui.ttf",
    "Segoe UI Light": "segoeuil.ttf",
    "Arial": "arial.ttf",
    "Arial Black": "ariblk.ttf",
    "Calibri": "calibri.ttf",
    "Consolas": "consolab.ttf",
    "Times New Roman": "times.ttf",
    "Georgia": "georgia.ttf",
    "Impact": "impact.ttf",
}

import os as _os

_FONTS_DIR = _os.path.join(_os.environ.get("SystemRoot", r"C:\Windows"), "Fonts")
_font_cache: dict[str, bool] = {}


def font_exists(family: str) -> bool:
    if family in _font_cache:
        return _font_cache[family]
    f = _FONT_FILES.get(family)
    ok = bool(f) and _os.path.exists(_os.path.join(_FONTS_DIR, f))
    _font_cache[family] = ok
    return ok


def pick_font(candidates, fallback: str = "Microsoft YaHei") -> str:
    """从候选列表取第一个本机存在的字体族名。"""
    if isinstance(candidates, str):
        candidates = [candidates]
    for c in candidates or []:
        if font_exists(c):
            return c
    return fallback


# 每套主题的排版配置。
# fonts 值为「候选列表」，get_theme 会按本机字体实测解析成真实族名。
# 搭配依据（slidesmate 两字体系统 / 清华讲义 / chengqingli 学术PPT）：
#   全篇 ≤2 个正文字体族 + 1 个数字/等宽族；标题与正文靠族或字重区分；
#   商务=雅黑标题+等线正文；学术=黑体标题+宋体正文+Times 数字+楷体引用；
#   政务=黑体标题+仿宋正文（公文惯例）；科技=雅黑Light大标题+等线正文+Consolas 数字。
# scale 依据（slidesmate/k-dense 字号表，960x540pt 画布）：
#   display 封面 44-54 · h1 页标题 32-40 · h2 卡片标题 ~20 · body 18-24(≥18)
#   small 13-14 · numeric 大数字 40-48 · caption 10-12 · 页脚 10-12
# transition/entrance 依据（verdanabold / animation-principles / Apple 0.8s cross-dissolve）：
#   整页默认 Fade 0.5-0.8s；入场只用 Fade 错落 0.15s；wipe 仅用于线/箭头的「画出」。
TYPO: dict[str, dict] = {
    "default": {
        "fonts": {
            "display": ["Microsoft YaHei"], "title": ["Microsoft YaHei"],
            "body": ["DengXian", "Microsoft YaHei"], "caption": ["DengXian Light", "DengXian"],
            "numeric": ["Segoe UI", "Arial"], "mono": ["Consolas"], "accent": ["KaiTi"],
        },
        "scale": {"display": 46, "h1": 32, "h2": 20, "body": 20, "small": 14,
                  "numeric": 44, "caption": 11},
        "transition": "fade", "entrance": "fade", "texture": "grid",
    },
}

# 图表分类色：Okabe-Ito 色盲安全调色板（Nature / k-dense 推荐）。
# 浅色底去掉低对比的 #F0E442；暗色底改用其亮色变体。绝不红绿单靠颜色区分。
OKABE_ITO_LIGHT = ["#0072B2", "#E69F00", "#009E73", "#D55E00", "#CC79A7", "#56B4E9", "#404040"]
OKABE_ITO_DARK = ["#56B4E9", "#E69F00", "#009E73", "#CC79A7", "#F0E442", "#D55E00", "#BBBBBB"]

# 整页切换效果名 → PowerPoint ppEntryEffect 常量（见 office.py）
TRANSITIONS = ("none", "fade", "push", "wipe", "split", "cover", "random")


def get_typo(theme_id: str) -> dict:
    base = TYPO.get("default")
    t = TYPO.get(str(theme_id).strip().lower()) or base
    merged = {**base, **t}
    merged["fonts"] = {**base["fonts"], **t.get("fonts", {})}
    merged["scale"] = {**base["scale"], **t.get("scale", {})}
    return merged


def get_theme(name: Optional[str] = None) -> dict:
    """取主题；name 为空或未知时回退 business。返回副本，附 layout + typo 段。"""
    key = str(name or "default").strip().lower()
    t = dict(THEMES.get(key) or THEMES["default"])
    layout = dict(BASE_LAYOUT)
    # 暗色主题的文字/背景对调，避免生成黑底黑字
    if t.get("rules", {}).get("darkMode"):
        layout["bodyMinFs"] = 13
    # 带品牌行（标识+字标）的主题，标题下移，内容起始随之下移
    if (t.get("brand") or {}).get("onContent"):
        layout["headerEndY"] = 98
        layout["contentStartY"] = 122
    t["layout"] = layout
    t["typo"] = get_typo(t.get("id", key))
    # 图表分类色：主题自带 chartPalette（如品牌主题的 品牌色系）优先，
    # 否则色盲安全 Okabe-Ito；暗色主题用亮色变体
    t["typo"]["chartPalette"] = (t["typo"].get("chartPalette")
                                 or (OKABE_ITO_DARK if t.get("rules", {}).get("darkMode")
                                     else OKABE_ITO_LIGHT))
    t["typo"].setdefault("chartStyle", {})
    t["typo"].setdefault("entrance", "fade")
    # 把字体候选解析成本机真实存在的族名，渲染器直接用，不再每次判断
    resolved = {}
    for role, cands in t["typo"]["fonts"].items():
        resolved[role] = pick_font(cands)
    t["typo"]["resolvedFonts"] = resolved
    return t


def theme_brief() -> list[dict]:
    """给模型看的精简清单（不含 layout 细节）。"""
    return [{"id": t["id"], "label": t["label"], "desc": t["desc"],
             "primary": t["primary"], "accent": t["accent"],
             "darkMode": bool(t.get("rules", {}).get("darkMode"))}
            for t in THEMES.values()]


def register_theme(spec: dict) -> dict:
    """注册自定义主题（用户在 config 里给的 overrides 也走这里）。"""
    tid = str(spec.get("id") or "custom").strip().lower()
    base = get_theme(spec.get("base") or "business")
    merged = {**base, **{k: v for k, v in spec.items() if v is not None}}
    merged["id"] = tid
    merged["layout"] = {**base["layout"], **(spec.get("layout") or {})}
    THEMES[tid] = merged
    return merged
