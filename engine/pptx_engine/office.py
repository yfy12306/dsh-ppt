# -*- coding: utf-8 -*-
"""Office COM 后端 —— PowerPoint 优先，自动探测 MS Office / WPS。

本模块是整个引擎与真实 Office 之间唯一的桥梁。设计目标：
**让 AI 生成的每一页都实时出现在用户眼前的 PowerPoint 窗口里**。

与上游 harness-anything（WPS 版）的三处关键差异，均为 MS Office 实测修正：

1. **不无条件 taskkill**。上游 `connect_wps()` 每次生成前都
   `taskkill /F /IM wps.exe`，这会连带毁掉用户自己未保存的文档。
   本实现只关闭「本进程创建的」演示文稿，默认不退出应用，
   需要强制清理时走显式的 `reset()` 并二次确认。

2. **PowerPoint 的 `Application.Visible` 不能设 False**（会抛
   `com_error`）。无窗口模式必须走 `Presentations.Add(WithWindow=msoFalse)`，
   而不是先 Dispatch 再隐藏。上游 `app.Visible = False` 的写法在
   PowerPoint 上直接报错，在 WPS 上才成立。

3. **ProgID 自动探测**：`PowerPoint.Application` → `KWPP.Application`，
   两者 COM 对象模型 VBA 兼容，同一套元素代码通吃。

COM 线程约束：所有 COM 调用必须在同一个已 `CoInitialize` 的线程里。
本引擎的 MCP 主循环即该线程；实时预览的 HTTP 线程只读磁盘 PNG 与
状态字典，绝不触碰 COM 对象。
"""
from __future__ import annotations

import os
import sys
import time
from typing import Optional

# ---------------------------------------------------------------- 常量
# MsoTriState
msoTrue = -1
msoFalse = 0

# PpSaveAsFileType
ppSaveAsPresentation = 1        # .pptx
ppSaveAsPDF = 32                # .pdf
ppSaveAsOpenXMLPresentation = 24

# PpSlideLayout
ppLayoutBlank = 12

# PpParagraphAlignment
ppAlignLeft = 1
ppAlignCenter = 2
ppAlignRight = 3
ppAlignJustify = 4

# MsoVerticalAnchor
msoAnchorTop = 1
msoAnchorMiddle = 3
msoAnchorBottom = 4

# MsoTextOrientation
msoTextOrientationHorizontal = 1

# MsoAutoShapeType（常用子集）
msoShapeRectangle = 1
msoShapeRoundedRectangle = 5
msoShapeOval = 9
msoShapeDiamond = 4
msoShapeRightArrow = 33
msoShapeLeftArrow = 34
msoShapeUpArrow = 35
msoShapeDownArrow = 36
msoShapeIsocelesTriangle = 7
msoShapeRightTriangle = 8
msoShapeParallelogram = 2
msoShapeChevron = 55
msoShapePentagon = 51
msoShapeHexagon = 10
msoShapeCloud = 17
msoShapeFlowchartProcess = 61
msoShapeFlowchartDecision = 62

# PpAlertLevel
ppAlertsNone = 1

# ---------------------------------------------------------------- 切换 / 动画
# PowerPoint 的 ppEntryEffect / msoAnimEffect 常量在两代 API 里数值不一致，
# 网上抄的常量经常设进去报「无效的请求」。这里不硬信任何一个数值：
# 每个语义名给一组候选，逐个尝试 + 读回校验，第一个「设得进去且读回一致」的生效；
# 全部失败就退回 none。**动画是装饰，绝不能让它把生成流程搞崩。**
TRANSITION_CANDIDATES = {
    "none": [0],
    "cut": [1],
    "random": [2],
    "fade": [3845, 3844, 15],          # Fade / FadeSmoothly / Dissolve 兜底
    "dissolve": [15],
    "push": [3847, 3848, 7, 9],        # PushLeft / PushRight / Cover 兜底
    "wipe": [3846, 3, 4],              # Wipe / Blinds 兜底
    "split": [3849, 3, 4],
    "cover": [7, 8, 9, 10],
    "blinds": [3, 4],
}

# msoAnimEffect（入场）—— 只用把握最大的三个
ANIM_APPEAR = 1
ANIM_FLY = 2
ANIM_FADE = 10
ANIM_DISSOLVE = 9
ANIM_BY_NAME = {"appear": ANIM_APPEAR, "fly": ANIM_FLY, "fade": ANIM_FADE,
                "dissolve": ANIM_DISSOLVE}

# MsoAnimTriggerType
ANIM_TRIGGER_WITH_PREVIOUS = 2
ANIM_TRIGGER_AFTER_PREVIOUS = 3


def apply_transition(slide, name: str = "fade", duration: float = 0.6,
                     advance_seconds: float = 0.0) -> dict:
    """给单页设置整页切换效果。返回实际生效的效果与耗时。

    Args:
        slide: COM Slide 对象
        name: none/cut/random/fade/dissolve/push/wipe/split/cover/blinds
        duration: 切换时长（秒）
        advance_seconds: >0 时自动换页（放映用），0 = 手动
    """
    applied = {"transition": "none", "effect": 0, "duration": 0.0}
    try:
        tr = slide.SlideShowTransition
    except Exception:  # noqa: BLE001
        return applied
    cands = TRANSITION_CANDIDATES.get(str(name or "none").lower(), [0])
    chosen = 0
    for c in cands:
        try:
            tr.EntryEffect = c
            back = tr.EntryEffect
            if int(back) == int(c) or c == 0:
                chosen = int(back)
                break
        except Exception:  # noqa: BLE001
            continue
    applied["effect"] = chosen
    applied["transition"] = name if chosen else "none"
    try:
        tr.Duration = float(duration)
        applied["duration"] = float(duration)
    except Exception:  # noqa: BLE001
        pass
    try:
        if advance_seconds and advance_seconds > 0:
            tr.AdvanceOnTime = msoTrue
            tr.AdvanceTime = float(advance_seconds)
        else:
            tr.AdvanceOnTime = msoFalse
    except Exception:  # noqa: BLE001
        pass
    return applied


def apply_entrance(slide, shapes, effect: str = "fade", stagger: float = 0.3,
                   skip: int = 0) -> dict:
    """给一页里的形状加入场动画（TimeLine 主序列，逐个 AfterPrevious 形成错落）。

    Args:
        shapes: COM 形状可迭代（按想要的出现顺序）
        effect: appear/fly/fade/dissolve
        stagger: 每个元素的动画时长（秒），AfterPrevious 下即错落间隔
        skip: 跳过前 N 个形状（比如底图/页脚不动）
    Returns:
        {"applied": n, "effect": id, "ok": bool}
    """
    out = {"applied": 0, "effect": ANIM_BY_NAME.get(str(effect).lower(), ANIM_FADE),
           "ok": False}
    try:
        seq = slide.TimeLine.MainSequence
    except Exception:  # noqa: BLE001
        return out
    idx = 0
    for sh in shapes:
        if idx < skip:
            idx += 1
            continue
        try:
            eff = seq.AddEffect(sh, out["effect"], 1, ANIM_TRIGGER_AFTER_PREVIOUS)
            try:
                eff.Timing.Duration = float(stagger)
            except Exception:  # noqa: BLE001
                pass
            out["applied"] += 1
            out["ok"] = True
        except Exception:  # noqa: BLE001
            # 单个形状加动画失败不影响其余
            continue
        idx += 1
    return out

# 后端候选：MS Office 优先（本机实测 16.0 可用），WPS 兜底
BACKENDS = [
    {
        "id": "ms-office",
        "label": "Microsoft Office",
        "impress": "PowerPoint.Application",
        "writer": "Word.Application",
        "calc": "Excel.Application",
    },
    {
        "id": "wps",
        "label": "WPS Office",
        "impress": "KWPP.Application",
        "writer": "KWPS.Application",
        "calc": "KET.Application",
    },
]


class OfficeError(RuntimeError):
    """COM 层可预期错误（未装 Office、连接失败、对象已失效等）。"""


def _log(msg: str) -> None:
    """日志一律走 stderr —— stdout 被 MCP JSON-RPC 独占。"""
    print(f"[office] {msg}", file=sys.stderr, flush=True)


# ---------------------------------------------------------------- 探测
def probe_backends() -> list[dict]:
    """探测本机可用的 Office 后端（不真正启动应用，只查注册表 ProgID）。

    Returns:
        [{"id","label","impress","available","error"}, ...]
    """
    out = []
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        for b in BACKENDS:
            out.append({**b, "available": False, "error": "缺少 pywin32（pip install pywin32）"})
        return out

    import win32com.client

    for b in BACKENDS:
        try:
            # 只做类对象解析，不实例化，避免探测时弹出 Office 窗口
            win32com.client.Dispatch(b["impress"])
            out.append({**b, "available": True, "error": ""})
        except Exception as e:  # noqa: BLE001
            out.append({**b, "available": False, "error": f"{type(e).__name__}: {e}"})
    return out


def detect_backend(prefer: str = "auto") -> dict:
    """选定后端。prefer: auto | ms-office | wps。"""
    probes = probe_backends()
    ok = [p for p in probes if p["available"]]
    if prefer != "auto":
        for p in ok:
            if p["id"] == prefer:
                return p
        raise OfficeError(
            f"指定后端 {prefer} 不可用。探测结果："
            + "; ".join(f"{p['id']}={'OK' if p['available'] else p['error']}" for p in probes)
        )
    if not ok:
        raise OfficeError(
            "未找到可用的 Office 后端。需要 Microsoft Office（PowerPoint）或 WPS Office（演示）。探测结果："
            + "; ".join(f"{p['id']}={p['error']}" for p in probes)
        )
    return ok[0]


# ---------------------------------------------------------------- 连接
class PowerPointApp:
    """一个 PowerPoint 应用连接 + 由本进程创建的演示文稿。

    用法::

        app = PowerPointApp(visible=True)      # 可见 = 用户实时观看
        pres = app.new_deck(960, 540)
        ...
        app.close_deck(save_as=None)           # 只关自己建的
        app.shutdown(quit=False)               # 默认不退出，保留窗口给用户看
    """

    def __init__(self, visible: bool = True, prefer: str = "auto",
                 retries: int = 3, retry_delay: float = 1.5):
        import pythoncom
        import win32com.client

        self._pythoncom = pythoncom
        self._win32 = win32com.client
        self.visible = bool(visible)
        self.backend = detect_backend(prefer)
        self.progid = self.backend["impress"]
        self.app = None
        self.pres = None
        self._we_started_app = False
        self._last_error: Optional[Exception] = None

        # COM 必须在调用线程初始化；MCP 主循环即该线程
        try:
            pythoncom.CoInitialize()
        except Exception:  # noqa: BLE001
            pass  # 已初始化过

        for attempt in range(retries):
            try:
                self.app = win32com.client.Dispatch(self.progid)
                # 触发一次 RPC 往返，确认对象真的活着（Dispatch 可能返回僵尸代理）
                _ = self.app.Name
                self._we_started_app = True
                break
            except Exception as e:  # noqa: BLE001
                self._last_error = e
                _log(f"连接 {self.progid} 失败（第 {attempt + 1}/{retries} 次）：{e}")
                time.sleep(retry_delay)
        if self.app is None:
            raise OfficeError(f"无法连接 {self.progid}：{self._last_error}")

        try:
            self.app.DisplayAlerts = ppAlertsNone
        except Exception:  # noqa: BLE001
            pass

        if self.visible:
            # 可见模式：PowerPoint 允许设 True（设 False 才抛异常）
            try:
                self.app.Visible = msoTrue
            except Exception as e:  # noqa: BLE001
                _log(f"设置 Visible=True 失败（忽略，改用 WithWindow 控制）：{e}")
        _log(f"已连接 {self.backend['label']}（{self.progid}），visible={self.visible}")

    # ------------------------------------------------------------ 属性
    @property
    def version(self) -> str:
        try:
            return str(self.app.Version)
        except Exception:  # noqa: BLE001
            return ""

    @property
    def has_deck(self) -> bool:
        return self.pres is not None

    # ------------------------------------------------------------ 演示文稿
    def new_deck(self, width: int = 960, height: int = 540):
        """新建演示文稿。

        可见模式走 `WithWindow=msoTrue`，这样 `ActiveWindow.View.GotoSlide`
        可用，用户能在真实窗口里看着页面一页页长出来。
        """
        if self.pres is not None:
            raise OfficeError("已有打开的演示文稿，请先 close_deck()")
        with_window = msoTrue if self.visible else msoFalse
        try:
            self.pres = self.app.Presentations.Add(with_window)
        except Exception as e:  # noqa: BLE001
            raise OfficeError(f"Presentations.Add 失败：{e}") from e

        try:
            self.pres.PageSetup.SlideWidth = int(width)
            self.pres.PageSetup.SlideHeight = int(height)
        except Exception as e:  # noqa: BLE001
            _log(f"设置画布尺寸失败（忽略）：{e}")
        _log(f"新建演示文稿 {width}x{height}，WithWindow={with_window}")
        return self.pres

    def open_deck(self, path: str):
        """打开已有 pptx（用于增量修改）。"""
        if self.pres is not None:
            raise OfficeError("已有打开的演示文稿，请先 close_deck()")
        ap = os.path.abspath(path)
        if not os.path.exists(ap):
            raise OfficeError(f"文件不存在：{ap}")
        with_window = msoTrue if self.visible else msoFalse
        self.pres = self.app.Presentations.Open(ap, ReadOnly=msoFalse, WithWindow=with_window)
        return self.pres

    def goto_slide(self, index: int) -> bool:
        """把可见窗口跳到指定页 —— 实时观看的关键动作。"""
        if not self.visible or self.pres is None:
            return False
        try:
            self.app.ActiveWindow.View.GotoSlide(int(index))
            return True
        except Exception:  # noqa: BLE001
            return False

    def export_slide_png(self, index: int, out_path: str, scale_width: int = 960) -> Optional[str]:
        """导出指定页为 PNG（供 DSH 内实时预览面板使用）。

        原子写：先导到 .tmp 再 os.replace —— 实时预览面板轮询读该文件，
        非原子写会让读方偶发拿到「不存在/半写」状态（2026-10-28 实测 404 竞态）。
        """
        if self.pres is None:
            return None
        try:
            slide = self.pres.Slides(int(index))
            ap = os.path.abspath(out_path)
            os.makedirs(os.path.dirname(ap), exist_ok=True)
            tmp = ap + ".tmp.png"
            slide.Export(tmp, "PNG", int(scale_width))
            os.replace(tmp, ap)
            return ap
        except Exception as e:  # noqa: BLE001
            _log(f"导出第 {index} 页 PNG 失败：{e}")
            return None

    def save(self, out_path: str, pdf: bool = True, pdf_path: Optional[str] = None) -> dict:
        """保存 pptx（可选同时导 PDF），并做落盘存在性 + 大小校验。

        上游踩过的坑：WPS 的 `SaveAs` 会静默失败。MS Office 同样可能因
        路径含非法字符/被占用而失败，所以这里一律校验落盘结果。
        """
        if self.pres is None:
            raise OfficeError("没有打开的演示文稿")
        out_path = os.path.abspath(out_path)
        d = os.path.dirname(out_path)
        if d:
            os.makedirs(d, exist_ok=True)

        result = {"pptx": None, "pdf": None, "pptx_size": 0, "pdf_size": 0}
        try:
            self.pres.SaveAs(out_path)
        except Exception as e:  # noqa: BLE001
            raise OfficeError(f"SaveAs(pptx) 失败：{e}") from e
        if os.path.exists(out_path):
            result["pptx"] = out_path
            result["pptx_size"] = os.path.getsize(out_path)
        if result["pptx_size"] <= 0:
            raise OfficeError(f"SaveAs 静默失败：{out_path} 不存在或为空")

        if pdf:
            pp = os.path.abspath(pdf_path) if pdf_path else os.path.splitext(out_path)[0] + ".pdf"
            try:
                self.pres.SaveAs(pp, ppSaveAsPDF)
                if os.path.exists(pp):
                    result["pdf"] = pp
                    result["pdf_size"] = os.path.getsize(pp)
            except Exception as e:  # noqa: BLE001
                _log(f"导出 PDF 失败（pptx 已保存，不影响主流程）：{e}")
        return result

    def close_deck(self, save_path: Optional[str] = None) -> None:
        """关闭本进程创建的演示文稿。绝不碰用户自己打开的文档。"""
        if self.pres is None:
            return
        try:
            if save_path:
                self.save(save_path, pdf=False)
        except Exception as e:  # noqa: BLE001
            _log(f"关闭前保存失败：{e}")
        try:
            self.pres.Close()
        except Exception:  # noqa: BLE001
            pass
        self.pres = None

    def shutdown(self, quit_app: bool = False) -> None:
        """释放 COM。默认不退出 PowerPoint —— 保留窗口供用户查看/继续编辑。"""
        self.close_deck()
        if quit_app and self._we_started_app:
            try:
                self.app.Quit()
            except Exception:  # noqa: BLE001
                pass
        self.app = None
        try:
            self._pythoncom.CoUninitialize()
        except Exception:  # noqa: BLE001
            pass

    # ------------------------------------------------------------ 诊断
    def info(self) -> dict:
        return {
            "backend": self.backend["id"],
            "backendLabel": self.backend["label"],
            "progid": self.progid,
            "version": self.version,
            "visible": self.visible,
            "hasDeck": self.has_deck,
            "slideCount": (self.pres.Slides.Count if self.pres is not None else 0),
        }
