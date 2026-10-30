# -*- coding: utf-8 -*-
"""dsh-ppt 引擎 —— 用真实 Office COM 做 PPT，全程可见。

精华来源：harness-anything（yb2460，MIT）的 pptengine + cli_anything/zotero，
针对 Microsoft Office 实测重写并补齐实时预览能力。

模块地图::

    office.py     PowerPoint COM 连接（MS Office 优先，WPS 兜底）
    deck.py       Deck 会话 + 实时可见三件套（GotoSlide / PNG 快照 / 操作流水）
    elements.py   绘图原语与组合组件（含上游 7 个 COM 坑的规避）
    renderer.py   声明式 JSON spec → 绘图调用的类型路由表
    assets.py     PIL 程序化素材（渐变横幅/竖条/圆环/箭头/背景/封面）
    charts.py     matplotlib 图表（中文字体已配好）
    verify.py     PyMuPDF 落盘校验（溢出/重叠/空页）+ 逐页渲染
    themes.py     7 套设计主题 + 机器可读的布局硬约束
    quality.py    pre-flight 静态审查 + 5 维度终评
    zotero.py     Zotero 只读检索 + GB/T 7714 / APA / IEEE 引用渲染
    live.py       实时预览 HTTP 侧车（只读，绝不触碰 COM）
    mcp_stdio.py  零依赖 MCP stdio 服务端
"""
from __future__ import annotations

__version__ = "1.0.0"
__all__ = ["__version__"]
