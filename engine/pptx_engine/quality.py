# -*- coding: utf-8 -*-
"""质量审查 —— 把上游 SKILL.md 里的「关键约束 ⚠️」从口口相传变成可执行检查。

上游 harness-anything 有两块质量资产：
  1. `styles/quality_checks.py` 的 5 维度审查清单（visual/pedagogy/proofreading/
     parity/substance，各带阈值 70/75/80/85/90）—— 但那是**给人看的清单**，
     代码里只实现了很粗的字号/密度检查。
  2. `skills/SKILL.md` 的「关键约束 ⚠️」—— 标题≤4字、间距≥24pt、y+h≤518、
     中文引号用「」、正文≥22pt、每页图表+表格≥2种 —— 这些是**真正踩过坑的规则**，
     但只写在文档里，靠模型自觉。

本模块把 (2) 全部实现成 pre-flight 静态检查（渲染前就能拦下问题，省掉一轮 COM 往返），
把 (1) 的 5 维度做成结构化评分骨架，并接上 verify.py 的 PDF 实测结果做 post-flight。
"""
from __future__ import annotations

from typing import Optional

from .elements import parse_color as _pc
from .themes import BASE_LAYOUT, get_theme


# ---------------------------------------------------------------- 对比度（WCAG）
def _rel_lum(rgb) -> float:
    def ch(c: float) -> float:
        c = c / 255.0
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    return 0.2126 * ch(rgb[0]) + 0.7152 * ch(rgb[1]) + 0.0722 * ch(rgb[2])


def contrast_ratio(a, b) -> float:
    """WCAG 相对亮度对比度。正文要求 ≥4.5:1，大字（≥24pt 或 ≥18pt 粗）≥3:1。"""
    la, lb = _rel_lum(a), _rel_lum(b)
    hi, lo = max(la, lb), min(la, lb)
    return (hi + 0.05) / (lo + 0.05)

# 上游 slide-excellence 的 5 维度审查清单（阈值原样保留）
REVIEW_DIMENSIONS = {
    "visual": {
        "name": "视觉审查",
        "checks": ["字体层级", "颜色对比度", "留白比例", "布局一致性", "图片质量"],
        "threshold": 70,
    },
    "pedagogy": {
        "name": "叙事结构审查",
        "checks": ["叙事弧完整性", "预备知识清晰", "示例充分", "符号一致性", "逻辑流畅度"],
        "threshold": 75,
    },
    "proofreading": {
        "name": "校对审查",
        "checks": ["拼写", "语法", "术语一致性", "标点", "字体溢出"],
        "threshold": 80,
    },
    "parity": {
        "name": "格式一致性",
        "checks": ["PPTX vs PDF 一致", "字体嵌入", "图片不丢失", "动画兼容"],
        "threshold": 85,
    },
    "substance": {
        "name": "内容实质",
        "checks": ["数据准确性", "引用完整性", "结论支撑", "方法可复现"],
        "threshold": 90,
    },
}

# 元素类型归类：哪些算「视觉元素」，哪些算「文本元素」
_VISUAL_TYPES = {"pic", "image", "chart", "table", "rect", "rrect", "rounded", "oval",
                 "circle", "shape", "arrow", "grad_bar", "cards", "cards_2x3",
                 "cards_1x4", "cards_1x4_info", "stat_cards", "timeline",
                 "timeline_horiz", "num_circle", "kpi", "hline", "vline", "cover", "toc"}
_TEXT_TYPES = {"text", "txt", "rich", "bullets", "pill", "chip", "tagline_bar",
               "gray_box", "quote", "header", "card_list_wide"}
_TITLE_TYPES = {"header", "cover"}


def _box_of(spec: dict) -> tuple:
    b = spec.get("box")
    if isinstance(b, dict):
        return (float(b.get("x", 0)), float(b.get("y", 0)),
                float(b.get("w", 0)), float(b.get("h", 0)))
    if isinstance(b, (list, tuple)) and len(b) >= 4:
        return tuple(float(v) for v in b[:4])
    return (float(spec.get("x", 0) or 0), float(spec.get("y", 0) or 0),
            float(spec.get("w", 0) or 0), float(spec.get("h", 0) or 0))


def _texts_of(spec: dict) -> list[str]:
    """把一个元素里所有可见文本抽出来（用于引号/长度/术语检查）。"""
    out = []
    for k in ("text", "title", "label", "value", "subtitle", "meta", "author", "note", "sec"):
        v = spec.get(k)
        if isinstance(v, str) and v.strip():
            out.append(v)
    for k in ("parts", "rows", "items"):
        v = spec.get(k)
        if isinstance(v, list):
            for it in v:
                if isinstance(it, str):
                    out.append(it)
                elif isinstance(it, dict):
                    for kk in ("text", "title", "body", "label", "value", "sub"):
                        if isinstance(it.get(kk), str):
                            out.append(it[kk])
                elif isinstance(it, (list, tuple)):
                    for sub in it:
                        if isinstance(sub, str):
                            out.append(sub)
                        elif isinstance(sub, (list, tuple)):
                            for s2 in sub:
                                if isinstance(s2, str):
                                    out.append(s2)
    hdr = spec.get("header") or spec.get("headerRow")
    if isinstance(hdr, list):
        out.extend([str(x) for x in hdr if isinstance(x, str)])
    return [t for t in out if t.strip()]


def check_slide(elements: list, layout: Optional[dict] = None,
                theme: Optional[dict] = None) -> dict:
    """对单页的**元素 spec 列表**做 pre-flight 静态审查。

    Returns:
        {"score": 0-100, "pass": bool, "issues": [{level,rule,msg}, ...]}
        level: error（必须改）| warn（建议改）| info
    """
    lay = {**BASE_LAYOUT, **(layout or {})}
    th = theme or get_theme()
    cw = float((th.get("layout") or lay).get("canvas", {}).get("w", lay["canvas"]["w"]))
    ch = float((th.get("layout") or lay).get("canvas", {}).get("h", lay["canvas"]["h"]))
    safe_bottom = float(lay["safeBottom"])
    content_start = float(lay["contentStartY"])
    min_gap = float(lay["minGap"])
    body_min = float(lay["bodyMinFs"])
    title_max_chars = int(lay["titleMaxChars"])

    issues: list[dict] = []
    score = 100

    def add(level, rule, msg, penalty):
        nonlocal score
        issues.append({"level": level, "rule": rule, "msg": msg})
        score -= penalty

    texts = [e for e in elements if str(e.get("type", "")).lower() in _TEXT_TYPES]
    visuals = [e for e in elements if str(e.get("type", "")).lower() in _VISUAL_TYPES]
    titles = [e for e in elements if str(e.get("type", "")).lower() in _TITLE_TYPES]

    # ---- 规则 1：所有元素不出画布（上游硬规则：y+h ≤ 518）----
    for e in elements:
        x, y, w, h = _box_of(e)
        if w <= 0 and h <= 0:
            continue  # header/hline 这类用 x/y/w 单独给的，跳过
        if y + h > safe_bottom + 0.5:
            add("error", "canvas-overflow",
                f"{e.get('type')} 超出安全底线：y+h={y + h:.0f} > {safe_bottom:.0f}", 8)
        if x + w > cw + 0.5:
            add("error", "canvas-overflow",
                f"{e.get('type')} 超出右边界：x+w={x + w:.0f} > {cw:.0f}", 8)
        if x < -0.5 or y < -0.5:
            add("warn", "canvas-negative", f"{e.get('type')} 坐标为负：({x:.0f},{y:.0f})", 3)

    # ---- 规则 2：标题与内容间距 ≥ 24pt（上游最常翻车的一条）----
    non_title_content = [e for e in elements
                         if str(e.get("type", "")).lower() not in _TITLE_TYPES]
    if titles and non_title_content:
        first_y = min((_box_of(e)[1] for e in non_title_content
                       if _box_of(e)[3] > 0), default=None)
        if first_y is not None and first_y < content_start:
            gap = first_y - float(lay["headerEndY"])
            if gap < min_gap:
                add("error", "title-gap",
                    f"首个内容元素 y={first_y:.0f}，与标题栏间距 {gap:.0f}pt < {min_gap:.0f}pt"
                    f"（内容应从 y≥{content_start:.0f} 开始）", 10)

    # ---- 规则 3：一页一个主题（标题元素不超过 1 个）----
    if len(titles) > 1:
        add("warn", "one-title", f"检测到 {len(titles)} 个标题元素，建议一页一主题", 6)

    # ---- 规则 4：标题长度（上游要求 ≤4 字，这里按主题放宽）----
    for t in titles:
        tt = str(t.get("title") or t.get("text") or "")
        if tt and len(tt) > title_max_chars:
            add("warn", "title-length",
                f"标题「{tt[:20]}」{len(tt)} 字 > {title_max_chars} 字，投影场景建议更短", 3)

    # ---- 规则 5：正文字号下限 ----
    for e in texts:
        fs = e.get("fs")
        if fs is None:
            continue
        try:
            fs = float(fs)
        except (TypeError, ValueError):
            continue
        typ = str(e.get("type", "")).lower()
        if typ in _TITLE_TYPES:
            continue
        if fs < body_min:
            add("warn", "font-too-small",
                f"{typ} 字号 {fs:g}pt < 下限 {body_min:g}pt（投影建议 ≥{lay['bodyRecommendFs']}pt）", 3)

    # ---- 规则 5b：WCAG 对比度（调研硬指标：正文 4.5:1，大字 3:1，优选 7:1）----
    bg_rgb = _pc(th.get("bg")) or (255, 255, 255)
    accent_rgb = _pc(th.get("accent"))
    for e in texts:
        c = _pc(e.get("color"))
        if c is None:
            continue
        try:
            fs = float(e.get("fs") or 0)
        except (TypeError, ValueError):
            fs = 0
        large = fs >= 24 or (fs >= 18 and bool(e.get("bold")))
        req = 3.0 if large else 4.5
        ratio = contrast_ratio(c, bg_rgb)
        if ratio < req:
            lvl = "error" if ratio < 3.0 else "warn"
            add(lvl, "contrast",
                f"{e.get('type')} 文字对比度 {ratio:.1f}:1 < 要求 {req:.1f}:1"
                f"（{'大字' if large else '正文'}，WCAG）", 6 if lvl == "error" else 3)
        # 强调色对比不足时只能用于形状/大数字，不能用于正文（调研：orange on white 2.3:1 失败案例）
        if accent_rgb and c == accent_rgb and ratio < 4.5:
            add("warn", "accent-misuse",
                f"强调色用于正文但对比度仅 {ratio:.1f}:1 —— 强调色应只用于形状/大数字/单条图表系列", 3)

    # ---- 规则 6：中文引号（上游硬规则：JSON 文本内引用用「」代替 ""）----
    for e in elements:
        for t in _texts_of(e):
            if '"' in t or '"' in t or '"' in t:
                add("warn", "cjk-quote",
                    f"文本含英文双引号，中文排版建议改用「」：{t[:30]}", 2)
                break

    # ---- 规则 7：内容密度（按主题 rules.maxBullets）----
    max_bullets = int((th.get("rules") or {}).get("maxBullets", 6))
    if len(texts) > max_bullets * 2:
        add("warn", "density",
            f"文本元素 {len(texts)} 个，超过主题上限 {max_bullets} 的 2 倍，考虑拆页", 6)

    # ---- 规则 8：视觉占比（上游 rules.visualRatio）----
    total = len(texts) + len(visuals)
    target_ratio = float((th.get("rules") or {}).get("visualRatio", 0.5))
    if total >= 4:
        ratio = len(visuals) / total
        if ratio < target_ratio - 0.25:
            add("info", "visual-ratio",
                f"视觉元素占比 {ratio:.0%} < 目标 {target_ratio:.0%}，纯文字页观感偏弱", 2)

    # ---- 规则 9：上游「每页图表+表格 ≥2 种」（仅对内容页，非强制）----
    kinds = {str(e.get("type", "")).lower() for e in visuals}
    if len(elements) >= 6 and not (kinds & {"chart", "table", "pic", "image"}):
        add("info", "no-visual-anchor", "本页无图表/表格/图片，数据类汇报建议至少一个", 2)

    score = max(0, min(100, score))
    errors = [i for i in issues if i["level"] == "error"]
    return {"score": score, "pass": not errors and score >= 70,
            "errors": len(errors), "warnings": len([i for i in issues if i["level"] == "warn"]),
            "issues": issues}


def check_deck(deck_spec: dict, theme: Optional[dict] = None) -> dict:
    """整份 pre-flight 审查。deck_spec 结构同 Deck.build() 的入参。"""
    th = theme or get_theme(deck_spec.get("theme"))
    lay = {**BASE_LAYOUT, **(th.get("layout") or {})}
    canvas = deck_spec.get("canvas") or {}
    if canvas:
        lay = {**lay, "canvas": {"w": float(canvas.get("w", lay["canvas"]["w"])),
                                 "h": float(canvas.get("h", lay["canvas"]["h"]))}}
    slides = deck_spec.get("slides") or []
    per = []
    for i, sd in enumerate(slides, start=1):
        r = check_slide(sd.get("elements") or [], layout=lay, theme=th)
        r["slide"] = i
        per.append(r)

    overall = sum(r["score"] for r in per) / max(1, len(per))
    all_issues = [dict(i, slide=r["slide"]) for r in per for i in r["issues"]]
    errors = [i for i in all_issues if i["level"] == "error"]

    # ---- 全篇级检查 ----
    deck_issues = []
    if len(slides) < 3:
        deck_issues.append({"level": "info", "rule": "deck-length",
                            "msg": f"仅 {len(slides)} 页，完整汇报通常 ≥5 页"})
    if len(slides) > 40:
        deck_issues.append({"level": "warn", "rule": "deck-length",
                            "msg": f"{len(slides)} 页偏多，建议拆分或精简"})
    # 封面/结尾
    types_first = {str(e.get("type", "")).lower() for e in (slides[0].get("elements") or [])} if slides else set()
    if slides and not (types_first & {"cover"}):
        deck_issues.append({"level": "info", "rule": "no-cover",
                            "msg": "首页未使用 cover 元素，建议加封面页"})
    # 标准页序提示（上游：S1 封面 → S2 目录 → S3-S10 内容 → S11 总结 → S12 致谢）
    if len(slides) >= 6:
        types_second = {str(e.get("type", "")).lower() for e in (slides[1].get("elements") or [])}
        if not (types_second & {"toc", "card_list_wide"}):
            deck_issues.append({"level": "info", "rule": "no-toc",
                                "msg": "第 2 页未使用 toc/card_list_wide，建议加目录页"})

    all_issues.extend(deck_issues)
    return {
        "overallScore": round(overall, 1),
        "pass": not errors and overall >= 70,
        "slides": len(per),
        "errors": len(errors),
        "warnings": len([i for i in all_issues if i["level"] == "warn"]),
        "infos": len([i for i in all_issues if i["level"] == "info"]),
        "perSlide": per,
        "issues": all_issues,
        "dimensions": REVIEW_DIMENSIONS,
        "summary": f"整体 {overall:.0f} 分 / {len(per)} 页 / "
                   f"{len(errors)} 个必须修 / {len([i for i in all_issues if i['level'] == 'warn'])} 个建议修",
    }


def merge_postflight(pre: dict, pdf_verify: Optional[dict] = None,
                     pptx_verify: Optional[dict] = None) -> dict:
    """把 verify.py 的 PDF 实测结果并进 pre-flight 报告，形成 5 维度终评。

    parity 维度靠「pptx 页数 == pdf 页数」判定；proofreading 靠 PDF 溢出/重叠实测。
    """
    dims = {}
    pre_score = float(pre.get("overallScore", 100))

    dims["visual"] = {"score": round(pre_score, 1),
                      "threshold": REVIEW_DIMENSIONS["visual"]["threshold"],
                      "note": "pre-flight 静态审查（画布/间距/字号/密度/视觉占比）"}

    proof = pre_score
    parity = 100.0
    if pdf_verify and pdf_verify.get("available"):
        ov = int(pdf_verify.get("overflow", 0))
        ol = int(pdf_verify.get("overlaps", 0))
        empty = len(pdf_verify.get("emptyPages") or [])
        proof -= min(40, ov * 6 + ol * 4 + empty * 8)
        dims["proofreading"] = {
            "score": round(max(0.0, proof), 1),
            "threshold": REVIEW_DIMENSIONS["proofreading"]["threshold"],
            "note": f"PDF 实测：溢出 {ov} / 重叠 {ol} / 空页 {empty}"}
        if pptx_verify and pptx_verify.get("available"):
            ps = int(pptx_verify.get("slides", 0))
            fs = int(pdf_verify.get("pages", 0))
            parity = 100.0 if ps == fs else max(0.0, 100.0 - abs(ps - fs) * 20)
            dims["parity"] = {"score": round(parity, 1),
                              "threshold": REVIEW_DIMENSIONS["parity"]["threshold"],
                              "note": f"PPTX {ps} 页 vs PDF {fs} 页"}
    else:
        dims["proofreading"] = {"score": round(proof, 1),
                                "threshold": REVIEW_DIMENSIONS["proofreading"]["threshold"],
                                "note": "未做 PDF 实测（缺 PyMuPDF 或未导 PDF）"}
        dims["parity"] = {"score": 0.0, "threshold": REVIEW_DIMENSIONS["parity"]["threshold"],
                          "note": "未导出 PDF，无法比对"}

    dims["pedagogy"] = {"score": round(pre_score, 1),
                        "threshold": REVIEW_DIMENSIONS["pedagogy"]["threshold"],
                        "note": "叙事结构需人工/模型判断，此处沿用静态分"}
    dims["substance"] = {"score": 0.0, "threshold": REVIEW_DIMENSIONS["substance"]["threshold"],
                         "note": "数据准确性与引用完整性无法自动判定，需人工确认"}

    passed = [k for k, v in dims.items() if v["score"] >= v["threshold"]]
    return {
        "dimensions": dims,
        "passed": passed,
        "failed": [k for k in dims if k not in passed],
        "overallScore": round(sum(v["score"] for v in dims.values()) / max(1, len(dims)), 1),
        "preFlight": pre,
        "pdfVerify": pdf_verify,
        "pptxVerify": pptx_verify,
    }
