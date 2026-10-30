# -*- coding: utf-8 -*-
"""Deck —— 一份演示文稿的会话封装，内置「实时可见」三件套。

实时观看的三个层次（本类负责前两个，第三个在 live.py）：

1. **真实 PowerPoint 窗口**：`visible=True` 时以 `WithWindow=msoTrue` 建文稿，
   每次加页/加元素后调 `GotoSlide()`，用户能看着页面在眼前一页页长出来。
2. **当前页 PNG 快照**：每次变更后把当前页导出到 `.live/slide.png`，
   DSH 客户端面板轮询取图，不用切窗口也能看。
3. **操作流水**：每个动作记一条 op（时间/类型/目标页/摘要），
   面板按时间倒序展示，形成「AI 正在做什么」的实时字幕。

上游 harness-anything 的 `Deck` 只有第 1 层的一半（GotoSlide），
且每次连接前无条件 taskkill；本实现补齐 2、3 层并去掉危险清理。
"""
from __future__ import annotations

import json
import os
import threading
import time
from pathlib import Path
from typing import Optional

from . import office as O
from . import elements as el


class Deck:
    """一个演示文稿会话。

    ::

        deck = Deck(theme="business", visible=True)
        deck.add_slide()
        deck.render({"type": "header", "sec": "1.1", "title": "背景"})
        deck.save("out/demo.pptx")
    """

    def __init__(self, width: int = 960, height: int = 540,
                 workdir: Optional[str] = None, visible: bool = True,
                 bg: Optional[str] = None, font: str = "微软雅黑",
                 prefer_backend: str = "auto", live: bool = True,
                 snapshot_scale: int = 960, theme: Optional[dict] = None,
                 open_path: Optional[str] = None):
        """
        Args:
            open_path: 传了则打开已有 pptx（画布尺寸从文件读回，忽略 width/height），
                       不传则新建空白文稿。
        """
        self.open_path = open_path
        self.width = int(width)
        self.height = int(height)
        self.font = font
        self.visible = bool(visible)
        self.bg = bg
        self.theme = theme or {}
        self.live = bool(live)
        self.snapshot_scale = int(snapshot_scale)

        self.workdir = Path(workdir).expanduser().resolve() if workdir else Path.cwd()
        self.live_dir = self.workdir / ".live"
        self.live_dir.mkdir(parents=True, exist_ok=True)
        self.snapshot_path = self.live_dir / "slide.png"
        self.state_path = self.live_dir / "state.json"

        el.FONT = font
        el.TITLE_FONT = font
        assets = self.theme.get("assetsDir")
        if assets:
            el.set_assets_dir(assets)
        # 排版系统注入：角色字体（本机实测解析后的族名）+ 字号层级
        typo = self.theme.get("typo") or {}
        el.set_typography(typo.get("resolvedFonts"), typo.get("scale"))

        # 动效与页脚默认值（主题可覆盖，单页可再覆盖）
        self.transition = self.theme.get("transition") or typo.get("transition") or "fade"
        self.transition_duration = float(self.theme.get("transitionDuration") or 0.6)
        self.entrance = self.theme.get("entrance") or typo.get("entrance") or "fade"
        self.entrance_stagger = float(self.theme.get("entranceStagger") or 0.3)
        self.footer_enabled = bool(self.theme.get("footer", True))
        self.footer_left = str(self.theme.get("footerText") or self.theme.get("label") or "")
        self._last_had_footer = False

        self._lock = threading.RLock()
        self._ops: list[dict] = []
        self._op_seq = 0
        self._current_slide = 0
        self._created_at = time.time()
        self._saved_to: Optional[str] = None

        self.app = O.PowerPointApp(visible=self.visible, prefer=prefer_backend)
        if open_path:
            self.pres = self.app.open_deck(open_path)
            # 画布尺寸以文件为准，否则后续元素的坐标校验会用错基准
            try:
                self.width = int(self.pres.PageSetup.SlideWidth)
                self.height = int(self.pres.PageSetup.SlideHeight)
            except Exception:  # noqa: BLE001
                pass
            self._saved_to = os.path.abspath(open_path)
            self._record("open_deck", self.slide_count,
                         f"打开 {os.path.basename(open_path)}（{self.slide_count} 页，"
                         f"{self.width}x{self.height}），后端={self.app.backend['label']}")
            self._current_slide = max(1, self.slide_count)
        else:
            self.pres = self.app.new_deck(self.width, self.height)
            self._record("new_deck", 0,
                         f"新建演示文稿 {self.width}x{self.height}，后端={self.app.backend['label']}")
        self._refresh_live(self._current_slide or 1)

    # ------------------------------------------------------------ 状态
    @property
    def slide_count(self) -> int:
        try:
            return int(self.pres.Slides.Count)
        except Exception:  # noqa: BLE001
            return 0

    @property
    def current_slide(self) -> int:
        return self._current_slide or self.slide_count

    def _record(self, kind: str, slide: int, detail: str = "") -> dict:
        """记一条操作流水（实时字幕的数据源）。"""
        with self._lock:
            self._op_seq += 1
            op = {
                "seq": self._op_seq,
                "t": time.time(),
                "kind": kind,
                "slide": int(slide),
                "detail": (detail or "")[:200],
            }
            self._ops.append(op)
            if len(self._ops) > 200:
                del self._ops[: len(self._ops) - 200]
        return op

    def _refresh_live(self, slide_index: Optional[int] = None) -> None:
        """刷新实时预览：跳到该页 + 导出 PNG + 落状态文件。

        全部包在 try 里 —— 预览失败绝不能影响 PPT 生成本身。
        """
        if not self.live:
            return
        try:
            idx = int(slide_index or self.current_slide)
            if idx < 1 or idx > self.slide_count:
                return  # 空文稿（0 页）时 Slides.Item(1) 会抛 out of range
            self._current_slide = idx
            self.app.goto_slide(idx)
            self.app.export_slide_png(idx, str(self.snapshot_path), self.snapshot_scale)
            self._write_state()
        except Exception:  # noqa: BLE001
            pass

    def _write_state(self) -> None:
        with self._lock:
            state = {
                "deck": {
                    "width": self.width,
                    "height": self.height,
                    "slideCount": self.slide_count,
                    "currentSlide": self._current_slide,
                    "visible": self.visible,
                    "savedTo": self._saved_to,
                    "createdAt": self._created_at,
                    "updatedAt": time.time(),
                    "theme": self.theme.get("id", ""),
                },
                "backend": self.app.info(),
                "ops": list(reversed(self._ops[-40:])),
                "snapshot": {
                    "path": str(self.snapshot_path),
                    "exists": self.snapshot_path.exists(),
                    "mtime": (self.snapshot_path.stat().st_mtime if self.snapshot_path.exists() else 0),
                    "slide": self._current_slide,
                },
            }
        tmp = self.state_path.with_suffix(".json.tmp")
        try:
            tmp.write_text(json.dumps(state, ensure_ascii=False), encoding="utf-8")
            os.replace(str(tmp), str(self.state_path))
        except Exception:  # noqa: BLE001
            pass

    def state(self) -> dict:
        """给 MCP 工具 / HTTP 面用的完整状态。"""
        with self._lock:
            return {
                "width": self.width,
                "height": self.height,
                "slideCount": self.slide_count,
                "currentSlide": self._current_slide,
                "visible": self.visible,
                "savedTo": self._saved_to,
                "theme": self.theme.get("id", ""),
                "workdir": str(self.workdir),
                "ops": list(reversed(self._ops[-40:])),
                "backend": self.app.info(),
            }

    # ------------------------------------------------------------ 页面
    def add_slide(self, bg: Optional[str] = None, layout: int = O.ppLayoutBlank,
                  goto: bool = True, transition: Optional[str] = None,
                  footer: Optional[bool] = None):
        """新增一页（默认空白版式）并铺底图 + 整页切换 + 自动页脚。"""
        idx = self.slide_count + 1
        s = self.pres.Slides.Add(idx, layout)
        bgp = bg or self.bg
        if bgp:
            p = el.resolve_image(bgp)
            if p:
                try:
                    s.FollowMasterBackground = O.msoFalse
                except Exception:  # noqa: BLE001
                    pass
                try:
                    s.Background.Fill.UserPicture(p)
                except Exception as e:  # noqa: BLE001
                    self._record("warn", idx, f"铺底图失败：{e}")
            else:
                self._record("warn", idx, f"底图不存在：{bgp}")

        # 整页切换（调研：默认 Fade 0.5-0.8s；无切换会像「仓促拼凑的 PDF」）
        tr_name = transition if transition is not None else self.transition
        tr = O.apply_transition(s, tr_name, self.transition_duration)

        # 自动页脚（封面/分隔页可传 footer=False 关掉）。
        # 页脚作为第 0 个形状画在最前，入场动画时跳过它。
        want_footer = self.footer_enabled if footer is None else bool(footer)
        self._last_had_footer = False
        if want_footer:
            try:
                el.footer(s, width=self.width, height=self.height,
                          left=self.footer_left, page=idx,
                          color=el.parse_color(self.theme.get("muted")) or el.GRAY_TX,
                          rule=el.parse_color(self.theme.get("light")) or el.GRAY_LN,
                          fs=el.scale_size("caption", 11))
                self._last_had_footer = True
            except Exception as e:  # noqa: BLE001
                self._record("warn", idx, f"页脚绘制失败：{e}")

        self._current_slide = idx
        self._record("add_slide", idx,
                     f"新增第 {idx} 页（切换={tr.get('transition')}，页脚={'有' if self._last_had_footer else '无'}）")
        if goto:
            self._refresh_live(idx)
        return s

    def slide(self, index: Optional[int] = None):
        """取指定页（默认当前页）。"""
        idx = int(index or self.current_slide)
        if idx < 1 or idx > self.slide_count:
            raise IndexError(f"页码越界：{idx}（共 {self.slide_count} 页）")
        return self.pres.Slides(idx)

    def remove_slide(self, index: int) -> None:
        idx = int(index)
        self.slide(idx).Delete()
        self._record("remove_slide", idx, f"删除第 {idx} 页")
        self._refresh_live(min(idx, max(1, self.slide_count)))

    def goto(self, index: int) -> None:
        """仅切换可见窗口到某页（不改内容）。"""
        idx = max(1, min(int(index), max(1, self.slide_count)))
        self._current_slide = idx
        self._refresh_live(idx)

    # ------------------------------------------------------------ 渲染
    def render(self, spec: dict, slide_index: Optional[int] = None):
        """按声明式 spec 在当前页（或指定页）画一个元素。

        这是模型侧最主要的入口：一个 dict 描述一个元素，
        具体路由见 renderer.py 的 ELEMENT_ROUTES。
        """
        from . import renderer
        s = self.slide(slide_index)
        idx = int(slide_index or self.current_slide)
        out = renderer.render_element(s, spec, self)
        kind = str(spec.get("type", "?"))
        label = str(spec.get("text") or spec.get("title") or spec.get("label") or "")[:40]
        self._record("element", idx, f"{kind} {label}".strip())
        self._refresh_live(idx)
        return out

    def render_many(self, specs: list, slide_index: Optional[int] = None) -> list:
        """批量画元素，只在最后刷新一次预览（避免每个元素都导 PNG 拖慢速度）。"""
        from . import renderer
        s = self.slide(slide_index)
        idx = int(slide_index or self.current_slide)
        out = []
        for spec in specs or []:
            try:
                out.append({"ok": True, "type": spec.get("type"),
                            "result": renderer.render_element(s, spec, self)})
            except Exception as e:  # noqa: BLE001
                out.append({"ok": False, "type": spec.get("type"),
                            "error": f"{type(e).__name__}: {e}"})
                self._record("error", idx, f"{spec.get('type')} 失败：{e}")
        self._record("elements", idx, f"批量绘制 {len(specs or [])} 个元素")
        # 入场动画（调研：只用 Fade 错落 0.15-0.3s；wipe 留给线/箭头；不耍杂耍）
        if self.entrance and str(self.entrance).lower() != "none":
            try:
                shapes = list(s.Shapes)
                skip = 1 if self._last_had_footer else 0
                res_anim = O.apply_entrance(s, shapes[skip:], effect=self.entrance,
                                            stagger=self.entrance_stagger)
                if res_anim.get("applied"):
                    self._record("animate", idx,
                                 f"入场动画 {res_anim['applied']} 个元素（{self.entrance}）")
            except Exception as e:  # noqa: BLE001
                self._record("warn", idx, f"入场动画失败（不影响内容）：{e}")
        self._refresh_live(idx)
        return out

    def build(self, deck_spec: dict) -> dict:
        """JSON 驱动整份构建 —— 上游 SKILL.md 里标注的「生产环境核心工作流」。

        deck_spec::

            {
              "canvas": {"w": 960, "h": 540},
              "theme": "business",
              "bg": "bg_page.png",
              "slides": [
                {"bg": "...", "elements": [ {...}, {...} ]},
                ...
              ]
            }
        """
        slides = deck_spec.get("slides") or []
        report = {"slides": [], "ok": 0, "failed": 0}
        for i, sd in enumerate(slides, start=1):
            self.add_slide(bg=sd.get("bg"), transition=sd.get("transition"),
                           footer=sd.get("footer"))
            res = self.render_many(sd.get("elements") or [], slide_index=i)
            bad = [r for r in res if not r.get("ok")]
            report["slides"].append({"index": i, "elements": len(res), "failed": len(bad),
                                     "errors": [b.get("error") for b in bad][:5]})
            report["ok"] += len(res) - len(bad)
            report["failed"] += len(bad)
        self._record("build", self.slide_count,
                     f"JSON 构建完成：{len(slides)} 页，成功 {report['ok']}，失败 {report['failed']}")
        self._refresh_live(self.slide_count)
        return report

    # ------------------------------------------------------------ 保存
    def save(self, out_path: str, pdf: bool = True, pdf_path: Optional[str] = None) -> dict:
        res = self.app.save(out_path, pdf=pdf, pdf_path=pdf_path)
        self._saved_to = res.get("pptx")
        self._record("save", self.slide_count,
                     f"已保存 {os.path.basename(res.get('pptx') or out_path)}"
                     + (f" + PDF" if res.get("pdf") else ""))
        self._write_state()
        return res

    def snapshot(self, slide_index: Optional[int] = None, out_path: Optional[str] = None) -> Optional[str]:
        """导出某页 PNG。默认写实时预览缓存路径。"""
        idx = int(slide_index or self.current_slide)
        p = out_path or str(self.snapshot_path)
        got = self.app.export_slide_png(idx, p, self.snapshot_scale)
        if got:
            self._record("snapshot", idx, f"导出第 {idx} 页 PNG")
        return got

    def close(self, save_path: Optional[str] = None, quit_app: bool = False) -> None:
        self.app.close_deck(save_path=save_path)
        self.app.shutdown(quit_app=quit_app)
        self._record("close", 0, "已关闭演示文稿")
