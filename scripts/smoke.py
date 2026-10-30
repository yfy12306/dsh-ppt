# -*- coding: utf-8 -*-
"""dsh-ppt 引擎独立冒烟测试 —— 不经过 DSH/MCP，直接驱动真实 PowerPoint COM。

跑法::

    python scripts/smoke.py                 # 可见模式，能看着 PPT 一页页长出来
    python scripts/smoke.py --headless      # 无窗口
    python scripts/smoke.py --keep-open     # 跑完不关 PowerPoint，留着人工看

覆盖：后端探测 → 主题 → 素材生成 → JSON 建整份（封面/目录/图表/表格/卡片/时间轴/
参考文献）→ pre-flight 质量审查 → 保存 pptx+pdf → PDF 溢出/重叠校验 →
实时预览 HTTP 侧车 → 逐页渲染 PNG。
"""
from __future__ import annotations

import argparse
import json
import os
import sys
import time
import urllib.request
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "engine"))

from pptx_engine import assets as A          # noqa: E402
from pptx_engine import charts as C          # noqa: E402
from pptx_engine import office as O          # noqa: E402
from pptx_engine import quality as Q         # noqa: E402
from pptx_engine import renderer as R        # noqa: E402
from pptx_engine import themes as T          # noqa: E402
from pptx_engine import verify as V          # noqa: E402
from pptx_engine import zotero as Z          # noqa: E402
from pptx_engine.deck import Deck            # noqa: E402
from pptx_engine.live import LiveServer      # noqa: E402

PASS, FAIL = "[PASS]", "[FAIL]"
_results: list[tuple[str, str, str]] = []


def check(name: str, fn):
    t0 = time.time()
    try:
        detail = fn()
        _results.append(("PASS", name, str(detail or "")))
        print(f"{PASS} {name}  ({time.time() - t0:.2f}s) {detail or ''}")
        return True
    except Exception as e:  # noqa: BLE001
        import traceback
        _results.append(("FAIL", name, f"{type(e).__name__}: {e}"))
        print(f"{FAIL} {name}  ({time.time() - t0:.2f}s) {type(e).__name__}: {e}")
        traceback.print_exc(limit=3)
        return False


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--headless", action="store_true", help="无窗口模式")
    ap.add_argument("--keep-open", action="store_true", help="跑完不关 PowerPoint")
    ap.add_argument("--theme", default="default", help="default | 自定义主题id")
    ap.add_argument("--out", default="")
    args = ap.parse_args()

    visible = not args.headless
    outdir = Path(args.out) if args.out else (Path.home() / "Desktop" / "dsh-ppt-out" / "smoke")
    outdir.mkdir(parents=True, exist_ok=True)
    assets_dir = outdir / ".ppt-assets"
    print(f"=== dsh-ppt 冒烟测试 ===\n输出目录：{outdir}\n可见模式：{visible}\n")

    # ---- 1. 后端探测 ----
    backend = {}

    def _probe():
        bs = O.probe_backends()
        for b in bs:
            print(f"    {b['id']:12s} {'OK' if b['available'] else 'X'}  {b.get('error', '')[:70]}")
        ok = [b for b in bs if b["available"]]
        if not ok:
            raise RuntimeError("没有可用 Office 后端")
        backend.update(ok[0])
        return f"选定 {ok[0]['label']}（{ok[0]['impress']}）"

    if not check("Office 后端探测", _probe):
        print("\n后端不可用，后续 COM 测试全部跳过。")
        return _summary()

    # ---- 2. 主题 ----
    theme = {}

    def _theme():
        th = T.get_theme(args.theme)
        theme.update(th)
        assert th.get("primary"), "主题缺 primary"
        assert th.get("layout", {}).get("safeBottom") == 518, "布局硬约束丢失"
        return f"{th['id']} / {th['label']} / primary={th['primary']} / 共 {len(T.theme_names())} 套"

    check("主题加载", _theme)

    # ---- 3. 素材生成 ----
    def _assets():
        A.ensure_default_assets(assets_dir, force=True)
        A.ensure_theme_assets(theme, assets_dir, force=True)
        items = A.list_assets(assets_dir)
        names = [i["name"] for i in items]
        for must in ("bg_page.png", "accent_bar.png", f"bar_{theme['id']}.png"):
            assert must in names, f"缺素材 {must}"
        return f"{len(items)} 个素材，共 {sum(i['size'] for i in items) // 1024} KB"

    check("PIL 素材生成", _assets)

    # ---- 4. 图表生成 ----
    chart_png = {}

    def _chart():
        sem = theme.get("semantic") or {}
        pie_colors = [sem.get("good"), sem.get("warn"), sem.get("bad")] if sem else None
        p = C.make_chart("bar", {
            "labels": ["Q1", "Q2", "Q3", "Q4"],
            "series": [{"name": "产量", "values": [120, 168, 142, 195]},
                       {"name": "目标", "values": [130, 150, 160, 180]}],
            "ylabel": "千台", "title": "季度产量 vs 目标",
        }, out_dir=outdir / ".charts", theme=theme)
        p2 = C.make_chart("pie", {
            "labels": ["合格", "返修", "报废"], "values": [94.2, 4.6, 1.2],
            "donut": True, "centerText": "94.2%",
            **({"colors": [c for c in pie_colors if c]} if pie_colors else {}),
        }, out_dir=outdir / ".charts", theme=theme)
        p3 = C.make_chart("line", {
            "labels": ["W1", "W2", "W3", "W4", "W5", "W6", "W7", "W8"],
            "series": [{"name": "外倾角均值", "values": [-0.12, -0.08, -0.15, -0.05,
                                                         -0.02, -0.06, -0.01, -0.03]},
                       {"name": "控制上限", "values": [0.10] * 8}],
            "ylabel": "deg", "title": "外倾角 8 周趋势",
        }, out_dir=outdir / ".charts", theme=theme)
        chart_png["bar"] = p
        chart_png["pie"] = p2
        chart_png["line"] = p3
        assert os.path.getsize(p) > 3000, "图表 PNG 太小，可能渲染失败"
        return f"bar={os.path.getsize(p) // 1024}KB pie={os.path.getsize(p2) // 1024}KB line={os.path.getsize(p3) // 1024}KB"

    check("matplotlib 图表（中文字体）", _chart)

    # ---- 5. 元素路由表 ----
    def _routes():
        names = R.route_names()
        for must in ("text", "rich", "header", "table", "chart", "cards", "cover",
                     "toc", "bullets", "timeline", "tagline_bar", "pic"):
            assert must in names, f"缺元素类型 {must}"
        return f"{len(names)} 种：{', '.join(names)}"

    check("元素类型路由表", _routes)

    # ---- 6. pre-flight 质量审查（先审一个故意有错的 spec）----
    def _quality_bad():
        bad = {"slides": [{"elements": [
            {"type": "header", "sec": "1.1", "title": "背景"},
            {"type": "text", "box": [16, 60, 928, 40], "text": "离标题太近", "fs": 9},
            {"type": "rect", "box": [16, 480, 928, 100], "fill": "#ccc"},  # y+h=580 > 518
        ]}]}
        rep = Q.check_deck(bad, theme=theme)
        rules = {i["rule"] for i in rep["issues"]}
        assert "canvas-overflow" in rules, f"没抓到画布溢出：{rules}"
        assert "title-gap" in rules, f"没抓到标题间距：{rules}"
        assert "font-too-small" in rules, f"没抓到字号过小：{rules}"
        assert not rep["pass"], "有 error 却判 pass"
        return f"抓到 {rep['errors']} error / {rep['warnings']} warn，规则={sorted(rules)}"

    check("质量审查（应拦下 3 类问题）", _quality_bad)

    # ---- 7. 真实 COM 建整份 PPT ----
    deck_holder = {}
    # 内容起始 y 随主题几何（品牌主题有品牌行 → 122；default → 80），演示稿自适应
    cy = float(theme["layout"]["contentStartY"])
    deck_spec = {
        "slides": [
            {"footer": False, "elements": [
                {"type": "cover", "title": "产线质量月报", "subtitle": "2026 年 10 月 · 数据驱动",
                 "meta": "汇报人：于飞洋  |  DES 数据组", "accent": theme["primary"]},
            ]},
            {"elements": [
                {"type": "header", "sec": "", "title": "目录", "color": theme["primary"]},
                {"type": "toc", "box": [80, cy + 30, 800, 518 - cy - 60], "cols": 2,
                 "items": ["整体合格率", "缺陷分布", "停线分析", "外倾角趋势",
                           "制动力监控", "改进行动", "下月计划"],
                 "accent": theme["primary"]},
            ]},
            {"footer": False, "elements": [
                {"type": "divider", "sec": "01", "title": "质量现状",
                 "subtitle": "合格率 · 缺陷 · 停线", "color": theme["primary"],
                 "accent": theme["secondary"]},
            ]},
            {"elements": [
                {"type": "header", "sec": "1.1",
                 "title": "合格率连续三月上升", "color": theme["primary"]},
                {"type": "stat_cards", "box": [48, cy, 864, 104],
                 "items": [{"value": "94.2%", "label": "一次合格率", "color": theme["primary"]},
                           {"value": "+1.8pt", "label": "环比", "color": "#009E73"},
                           {"value": "312", "label": "停线分钟", "color": "#D55E00"},
                           {"value": "17", "label": "改善项", "color": "#E69F00"}]},
                {"type": "pic", "box": [48, cy + 120, 430, 518 - cy - 176], "src": chart_png.get("bar", "")},
                {"type": "pic", "box": [494, cy + 120, 418, 518 - cy - 176], "src": chart_png.get("pie", "")},
                {"type": "tagline_bar", "box": [48, 468, 864, 36],
                 "text": "报废率首次降到 1.2% 以下",
                 "color": theme["primary"]},
            ]},
            {"elements": [
                {"type": "header", "sec": "2.1",
                 "title": "螺簧档错档是最大缺陷源", "color": theme["primary"]},
                {"type": "table", "box": [48, cy, 864, 250],
                 "header": ["缺陷类型", "发生数", "占比", "主要工位", "状态"],
                 "cols": [2.2, 1, 1, 2, 1.2],
                 "rows": [
                     ["螺簧档错档", "42", "28.4%", "D13 / D21 T1", "已定位"],
                     ["外倾角超差", "31", "20.9%", "四轮检测线", "监控中"],
                     ["制动力偏差", "26", "17.6%", "制动检测线", "待业务确认"],
                     ["漆面瑕疵", "18", "12.2%", "涂装线", "已改善"],
                     ["其他", "31", "20.9%", "—", "—"]]},
                {"type": "cards", "box": [48, cy + 262, 864, 518 - cy - 274], "cols": 3,
                 "items": [{"title": "根因", "body": "vehicle_config 字段缺失"},
                           {"title": "影响", "body": "D13/D21 T1 螺簧档错档"},
                           {"title": "解法", "body": "恢复小时级管道字段"}]},
            ]},
            {"elements": [
                {"type": "header", "sec": "3.1",
                 "title": "三步走闭环改善", "color": theme["primary"]},
                {"type": "timeline", "box": [60, cy + 40, 840, 160],
                 "items": [{"label": "10-08", "text": "发现错档"},
                           {"label": "10-15", "text": "定位字段空洞"},
                           {"label": "10-22", "text": "实锤全量缺失"},
                           {"label": "10-28", "text": "推动管道修复"}],
                 "color": theme["primary"]},
                {"type": "bullets", "box": [70, cy + 216, 820, 518 - cy - 230],
                 "items": ["短期：按「不分档参考值」解读，避免误判",
                           "中期：小时级管道恢复 vehicle_config 字段",
                           "长期：字段级监控告警，缺失即熔断"],
                 "color": theme["text"]},
            ]},
            {"elements": [
                {"type": "header", "sec": "4.1",
                 "title": "外倾角趋势已收敛", "color": theme["primary"]},
                {"type": "pic", "box": [48, cy, 864, 518 - cy - 60],
                 "src": chart_png.get("line", "")},
                {"type": "tagline_bar", "box": [48, 468, 864, 36],
                 "text": "均值收敛至 ±0.05° 控制带内",
                 "color": theme["primary"]},
            ]},
        ]
    }

    def _build():
        d = Deck(width=960, height=540, workdir=str(outdir), visible=visible,
                 font="微软雅黑",
                 prefer_backend="auto", live=True, theme={**theme, "assetsDir": str(assets_dir)})
        deck_holder["deck"] = d
        rep = d.build(deck_spec)
        assert d.slide_count == len(deck_spec["slides"]), \
            f"页数不符：{d.slide_count} != {len(deck_spec['slides'])}"
        assert rep["failed"] == 0, f"有元素绘制失败：{rep}"
        return f"{d.slide_count} 页，元素成功 {rep['ok']} / 失败 {rep['failed']}"

    if not check("COM 建整份 PPT（实时可见）", _build):
        return _summary()

    d = deck_holder["deck"]

    # ---- 8. 实时预览侧车 ----
    def _live():
        srv = LiveServer(lambda: d.live_dir, lambda: {"engine": "smoke"}, port=8790)
        res = srv.start()
        assert res.get("ok"), f"侧车启动失败：{res}"
        time.sleep(0.4)
        with urllib.request.urlopen(f"{res['url']}/health", timeout=5) as r:
            health = json.loads(r.read().decode("utf-8"))
        with urllib.request.urlopen(f"{res['url']}/state", timeout=5) as r:
            state = json.loads(r.read().decode("utf-8"))
        with urllib.request.urlopen(f"{res['url']}/snapshot.png", timeout=5) as r:
            png = r.read()
        with urllib.request.urlopen(f"{res['url']}/ops?since=0", timeout=5) as r:
            ops = json.loads(r.read().decode("utf-8"))
        assert health.get("ok"), "health 不 ok"
        assert state.get("deck", {}).get("slideCount") == d.slide_count, "state 页数不符"
        assert len(png) > 5000, f"快照 PNG 太小：{len(png)}"
        assert png[:8] == b"\x89PNG\r\n\x1a\n", "不是合法 PNG"
        assert len(ops.get("ops") or []) > 0, "操作流水为空"
        deck_holder["live"] = srv
        return (f"port={res['port']} 快照={len(png) // 1024}KB "
                f"ops={len(ops['ops'])} 条 state.slideCount={state['deck']['slideCount']}")

    check("实时预览 HTTP 侧车", _live)

    # ---- 9. 跳页 + 单页快照（实时观看的核心动作）----
    def _goto():
        d.goto(2)
        assert d.current_slide == 2, f"跳页失败：{d.current_slide}"
        p = d.snapshot(2, str(outdir / "slide2.png"))
        assert p and os.path.getsize(p) > 3000, "单页快照失败"
        d.goto(d.slide_count)
        return f"跳到第 2 页并导出 {os.path.getsize(p) // 1024}KB，再跳到第 {d.slide_count} 页"

    check("跳页 + 单页快照", _goto)

    # ---- 10. 保存 pptx + pdf ----
    saved = {}

    def _save():
        res = d.save(str(outdir / "smoke-deck.pptx"), pdf=True)
        saved.update(res)
        assert res.get("pptx") and res["pptx_size"] > 10000, f"pptx 落盘异常：{res}"
        assert res.get("pdf") and res["pdf_size"] > 5000, f"pdf 落盘异常：{res}"
        return f"pptx={res['pptx_size'] // 1024}KB pdf={res['pdf_size'] // 1024}KB"

    if not check("保存 pptx + pdf（含落盘校验）", _save):
        return _summary()

    # ---- 11. PDF 溢出/重叠校验 ----
    def _verify():
        rep = V.verify_pdf(saved["pdf"], width=960, height=540, verbose=True)
        assert rep["available"], f"PyMuPDF 不可用：{rep.get('error')}"
        assert rep["pages"] == d.slide_count, f"PDF 页数 {rep['pages']} != {d.slide_count}"
        xv = V.verify_pptx(saved["pptx"])
        review = Q.merge_postflight({"overallScore": 100.0, "issues": []},
                                    pdf_verify=rep, pptx_verify=xv)
        deck_holder["verify"] = rep
        return (f"页数={rep['pages']} 溢出={rep['overflow']} 重叠={rep['overlaps']} "
                f"空页={len(rep['emptyPages'])} | pptx 读回 {xv.get('slides')} 页 "
                f"{xv.get('shapes')} 形状 | parity={review['dimensions']['parity']['score']}")

    check("PDF 落盘校验（溢出/重叠/空页/parity）", _verify)

    # ---- 12. 逐页渲染 PNG ----
    def _render():
        files = V.render_pages(saved["pdf"], str(outdir / "pages"), dpi=100)
        assert len(files) == d.slide_count, f"渲染页数 {len(files)} != {d.slide_count}"
        return f"{len(files)} 页 → {outdir / 'pages'}"

    check("逐页渲染 PNG", _render)

    # ---- 13. Zotero（未安装则验证降级路径）----
    def _zotero():
        p = Z.probe()
        if not p["sqliteExists"]:
            # 降级路径本身也是要验的：不能抛异常，必须给出可操作的提示
            assert p["installed"] is False
            assert "zotero.sqlite" in p["error"], f"降级提示不可操作：{p['error']}"
            r = Z.search("test")
            assert r["available"] is False and r["items"] == [], "未安装时 search 应优雅降级"
            return f"未安装，降级正常（提示：{p['error'][:50]}…）"
        r = Z.search(limit=3)
        return f"已安装，条目 {p['itemCount']}，检索到 {r['count']} 条"

    check("Zotero 探测与降级", _zotero)

    # ---- 14. 引用渲染（不依赖 Zotero 数据，用构造样本验格式）----
    def _cite():
        sample = {"title": "基于数据驱动的整车外倾角监控方法",
                  "year": "2026", "date": "2026-03",
                  "publication": "汽车工程", "volume": "48", "issue": "3",
                  "pages": "112-120", "doi": "10.19562/j.chinasae.qcgc.2026.03.012",
                  "type": "journalArticle",
                  "creators": [{"last": "于", "first": "飞洋", "typeId": 1, "index": 0},
                               {"last": "李", "first": "赛赛", "typeId": 1, "index": 1}]}
        gbt = Z.render_citation(sample, "gbt7714")
        apa = Z.render_citation(sample, "apa")
        ieee = Z.render_citation(sample, "ieee")
        assert "[J]" in gbt, f"GB/T 7714 缺文献类型标识：{gbt}"
        assert "等" not in gbt, "2 位作者不该出现「等」"
        assert "于飞洋" in gbt, f"中文姓名应连写全名，不该缩写成首字母：{gbt}"
        assert "于飞洋" in apa, f"APA 中文姓名同样应连写：{apa}"
        assert "(2026)" in apa, f"APA 年份格式错：{apa}"
        assert '"' in ieee, f"IEEE 标题该带引号：{ieee}"
        # 西文姓名仍须走缩写规则，别被 CJK 分支带偏
        west = {"title": "Deep Residual Learning", "year": "2016", "date": "2016",
                "publication": "CVPR", "type": "conferencePaper", "doi": "",
                "volume": "", "issue": "", "pages": "770-778", "publisher": "",
                "creators": [{"last": "He", "first": "Kaiming", "typeId": 1, "index": 0}]}
        wg = Z.render_citation(west, "gbt7714")
        assert "He K" in wg, f"西文姓名应缩写为「He K」：{wg}"
        slides = Z.refs_slide_spec([sample, sample], style="gbt7714", theme=theme)
        assert len(slides) == 1 and slides[0]["elements"], "参考文献页 spec 生成失败"
        return f"GBT={gbt[:46]}… | APA={apa[:40]}… | refs页 {len(slides)} 页"

    check("引用渲染（GB/T 7714 / APA / IEEE）", _cite)

    # ---- 收尾 ----
    try:
        if deck_holder.get("live"):
            deck_holder["live"].stop()
    except Exception:  # noqa: BLE001
        pass
    try:
        d.close(quit_app=not args.keep_open)
    except Exception:  # noqa: BLE001
        pass

    print(f"\n产物目录：{outdir}")
    return _summary()


def _summary() -> int:
    npass = sum(1 for r in _results if r[0] == "PASS")
    nfail = sum(1 for r in _results if r[0] == "FAIL")
    print("\n" + "=" * 62)
    print(f"冒烟测试结果：{npass} 通过 / {nfail} 失败 / 共 {len(_results)} 项")
    print("=" * 62)
    for status, name, detail in _results:
        mark = "✓" if status == "PASS" else "✗"
        print(f"  {mark} {name}")
        if status == "FAIL":
            print(f"      {detail}")
    return 0 if nfail == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
