/**
 * dsh-ppt（client 半）：输入框右侧「PPT 实时预览」芯片 + 弹层面板。
 *
 * 实时性设计（这是本插件区别于普通 PPT 工具的关键体验）：
 *   - 芯片状态点：绿 = 引擎在线且有文稿在写；灰 = 空闲；黄 = 引擎在线但无文稿
 *   - 弹层打开后每 1.2s 轮询 /ppt/state（小 JSON），比对 snapshot.mtime /
 *     slideCount / 最新 op.seq，只有真的变了才换 <img> 的 src（带 mtime 防缓存）
 *     —— 既不浪费带宽，又能做到「AI 每画一笔，缩略图立刻跟着变」
 *   - 操作流水按时间倒序展示，形成「AI 正在做什么」的实时字幕
 *   - 页面隐藏时暂停轮询；弹层关闭时降回 60s 只刷芯片
 *
 * 全部走同源路由 /ppt/*（host 半代理到引擎只读侧车），与 dsh-copilot 同模式。
 * 颜色全用主题 CSS 变量（带兜底值），适配 moon-flavor 等主题。
 */
window.__ModuleLoader__.load({
  id: "dsh-ppt",
  factory: function (require) {
    var module = { exports: {} };
    var exports = module.exports;
    Object.defineProperty(exports, Symbol.toStringTag, { value: "Module" });

    var React = require("react");
    var h = React.createElement;

    // ---------- 样式 ----------
    if (typeof document !== "undefined") {
      var tagId = "dsh-ppt/style";
      if (document.querySelector('style[data-plugin-css="' + tagId + '"]') === null) {
        var styleTag = document.createElement("style");
        styleTag.setAttribute("data-plugin-css", tagId);
        styleTag.setAttribute("data-plugin", "dsh-ppt");
        styleTag.textContent = [
          ".dsh-ppt-wrap { position: relative; display: inline-flex; }",
          ".dsh-ppt-btn {",
          "  display: inline-flex; align-items: center; gap: 5px;",
          "  height: 28px; margin: 0 2px; padding: 0 7px;",
          "  border: none; border-radius: 8px; background: transparent;",
          "  color: var(--dsw-alias-label-secondary, #9a9aa2);",
          "  cursor: pointer; transition: background 120ms ease, color 120ms ease; flex: none;",
          "}",
          ".dsh-ppt-btn:hover { background: var(--dsw-alias-fill-quaternary, rgba(120,120,130,0.12)); color: var(--dsw-alias-label-primary, currentColor); }",
          ".dsh-ppt-btn svg { display: block; }",
          ".dsh-ppt-dot { width: 7px; height: 7px; border-radius: 50%; background: var(--dsw-alias-label-quaternary, #55555c); transition: background 200ms ease; flex: none; }",
          ".dsh-ppt-dot-on { background: #30a46c; }",
          ".dsh-ppt-dot-idle { background: #f5a623; }",
          ".dsh-ppt-dot-live { background: #30a46c; animation: dsh-ppt-pulse 1.4s ease-in-out infinite; }",
          "@keyframes dsh-ppt-pulse { 0%,100% { box-shadow: 0 0 0 0 rgba(48,164,108,0.5); } 50% { box-shadow: 0 0 0 4px rgba(48,164,108,0); } }",
          ".dsh-ppt-pop {",
          "  position: absolute; bottom: 34px; right: 0; width: 400px; z-index: 9999;",
          "  padding: 12px; border-radius: 12px;",
          "  background: var(--dsw-alias-bg-elevated, var(--dsw-alias-bg-secondary, #26262b));",
          "  border: 1px solid var(--dsw-alias-border-l2, rgba(120,120,130,0.35));",
          "  box-shadow: 0 8px 28px rgba(0,0,0,0.35);",
          "  color: var(--dsw-alias-label-primary, #eaeaf0);",
          "  font-size: 12px; line-height: 1.5;",
          "}",
          ".dsh-ppt-pop h4 { margin: 0 0 8px; font-size: 13px; font-weight: 600; display: flex; align-items: center; gap: 6px; }",
          ".dsh-ppt-stage {",
          "  position: relative; width: 100%; aspect-ratio: 16/9; border-radius: 8px; overflow: hidden;",
          "  background: #101014; border: 1px solid var(--dsw-alias-border-l2, rgba(120,120,130,0.35));",
          "  display: flex; align-items: center; justify-content: center;",
          "}",
          ".dsh-ppt-stage img { width: 100%; height: 100%; object-fit: contain; display: block; }",
          ".dsh-ppt-stage .ph { color: var(--dsw-alias-label-tertiary, #77777f); font-size: 12px; text-align: center; padding: 0 16px; }",
          ".dsh-ppt-badge {",
          "  position: absolute; top: 6px; left: 6px; padding: 2px 8px; border-radius: 6px;",
          "  background: rgba(0,0,0,0.55); color: #fff; font-size: 11px; backdrop-filter: blur(4px);",
          "}",
          ".dsh-ppt-live-tag {",
          "  position: absolute; top: 6px; right: 6px; padding: 2px 8px; border-radius: 6px;",
          "  background: rgba(48,164,108,0.85); color: #fff; font-size: 11px; font-weight: 600;",
          "}",
          ".dsh-ppt-meta { display: flex; flex-wrap: wrap; gap: 4px 12px; margin-top: 8px; color: var(--dsw-alias-label-secondary, #9a9aa2); }",
          ".dsh-ppt-meta b { color: var(--dsw-alias-label-primary, #eaeaf0); font-weight: 600; }",
          ".dsh-ppt-ops { margin-top: 8px; max-height: 132px; overflow-y: auto; border-top: 1px solid var(--dsw-alias-border-l2, rgba(120,120,130,0.25)); padding-top: 6px; }",
          ".dsh-ppt-op { display: flex; gap: 6px; padding: 1.5px 0; font-size: 11.5px; color: var(--dsw-alias-label-secondary, #9a9aa2); }",
          ".dsh-ppt-op .k { flex: none; color: var(--dsw-alias-label-tertiary, #77777f); font-family: ui-monospace, Consolas, monospace; }",
          ".dsh-ppt-op .p { flex: none; color: #4e9bd8; font-family: ui-monospace, Consolas, monospace; }",
          ".dsh-ppt-op .d { word-break: break-all; }",
          ".dsh-ppt-acts { display: flex; flex-wrap: wrap; gap: 6px; margin-top: 10px; }",
          ".dsh-ppt-act {",
          "  padding: 4px 10px; border-radius: 8px; cursor: pointer; font-size: 12px;",
          "  border: 1px solid var(--dsw-alias-border-l2, rgba(120,120,130,0.35));",
          "  background: var(--dsw-alias-fill-quaternary, rgba(120,120,130,0.10));",
          "  color: var(--dsw-alias-label-primary, inherit);",
          "}",
          ".dsh-ppt-act:hover { background: var(--dsw-alias-fill-tertiary, rgba(120,120,130,0.20)); }",
          ".dsh-ppt-err { margin-top: 8px; color: #e5484d; word-break: break-all; }",
          ".dsh-ppt-hint { margin-top: 8px; color: var(--dsw-alias-label-tertiary, #77777f); }"
        ].join("\n");
        document.head.appendChild(styleTag);
      }
    }

    // ---------- 路由 ----------
    var ROUTE_STATUS = "/ppt/status";
    var ROUTE_STATE = "/ppt/state";
    var ROUTE_SHOT = "/ppt/snapshot.png";

    function fmtTime(ts) {
      try {
        var d = new Date(ts * 1000);
        var p = function (n) { return n < 10 ? "0" + n : "" + n; };
        return p(d.getHours()) + ":" + p(d.getMinutes()) + ":" + p(d.getSeconds());
      } catch (e) { return ""; }
    }

    function PptIcon() {
      return h("svg", { viewBox: "0 0 24 24", width: "15", height: "15", fill: "none",
        stroke: "currentColor", strokeWidth: "2", "aria-hidden": "true" },
        h("rect", { x: "3", y: "4", width: "18", height: "13", rx: "1.5" }),
        h("path", { d: "M8 21h8M12 17v4" }),
        h("path", { d: "M7.5 12.5l3-3.5 2.5 2.5 3.5-4" })
      );
    }

    function makeChip() {
      return function PptChip() {
        var s1 = React.useState(null); var status = s1[0]; var setStatus = s1[1];
        var s2 = React.useState(false); var open = s2[0]; var setOpen = s2[1];
        var s3 = React.useState(null); var live = s3[0]; var setLive = s3[1];
        var s4 = React.useState(""); var error = s4[0]; var setError = s4[1];
        var s5 = React.useState(0); var shotMtime = s5[0]; var setShotMtime = s5[1];
        var s6 = React.useState(typeof document !== "undefined" && document.hidden); var hidden = s6[0]; var setHidden = s6[1];
        var rootRef = React.useRef(null);

        var pollMs = 60000;
        var popupMs = 1200;
        if (status && status.ui) {
          if (status.ui.pollIntervalMs) pollMs = status.ui.pollIntervalMs;
          if (status.ui.popupPollIntervalMs) popupMs = status.ui.popupPollIntervalMs;
        }

        function refreshStatus() {
          return fetch(ROUTE_STATUS, { method: "GET" })
            .then(function (r) { return r.json(); })
            .then(function (j) { setStatus(j); if (j && j.error) setError(""); })
            .catch(function (e) { setError("状态获取失败：" + (e && e.message ? e.message : e)); });
        }

        function refreshLive() {
          return fetch(ROUTE_STATE, { method: "GET" })
            .then(function (r) { return r.ok ? r.json() : Promise.reject(new Error("HTTP " + r.status)); })
            .then(function (j) {
              setLive(j);
              var m = j && j.snapshot ? j.snapshot.mtime : 0;
              if (m && m !== shotMtime) setShotMtime(m);
            })
            .catch(function (e) { setError("实时状态获取失败：" + (e && e.message ? e.message : e)); });
        }

        // 芯片轮询（低频）
        React.useEffect(function () {
          if (hidden) return undefined;
          refreshStatus();
          var t = setInterval(refreshStatus, pollMs);
          return function () { clearInterval(t); };
        }, [pollMs, hidden]);

        // 弹层轮询（高频 = 实时）
        React.useEffect(function () {
          if (!open || hidden) return undefined;
          refreshLive();
          var t = setInterval(refreshLive, popupMs);
          return function () { clearInterval(t); };
        }, [open, popupMs, hidden]);

        React.useEffect(function () {
          var onVis = function () { setHidden(document.hidden); };
          document.addEventListener("visibilitychange", onVis);
          var onDown = function (ev) {
            if (rootRef.current && !rootRef.current.contains(ev.target)) setOpen(false);
          };
          document.addEventListener("mousedown", onDown);
          return function () {
            document.removeEventListener("visibilitychange", onVis);
            document.removeEventListener("mousedown", onDown);
          };
        }, []);

        var liveUp = Boolean(status && status.live && status.live.up);
        var deck = (live && live.deck) || (status && status.deck) || null;
        var writing = Boolean(deck && deck.slideCount > 0);
        var dotClass = "dsh-ppt-dot" + (liveUp && writing ? " dsh-ppt-dot-live" : (liveUp ? " dsh-ppt-dot-idle" : ""));
        var title = liveUp
          ? (writing ? ("PPT 引擎在线 · " + deck.slideCount + " 页 · 点击看实时预览") : "PPT 引擎在线 · 无文稿")
          : "PPT 引擎离线";

        var ops = (live && live.deck && live.deck.ops) || (deck && deck.ops) || [];
        var shotSrc = shotMtime ? (ROUTE_SHOT + "?m=" + shotMtime) : "";

        var pop = open ? h("div", { className: "dsh-ppt-pop" },
          h("h4", null, h(PptIcon), "PPT 实时预览",
            liveUp ? h("span", { className: "dsh-ppt-dot dsh-ppt-dot-live" }) : null),
          h("div", { className: "dsh-ppt-stage" },
            shotSrc
              ? h("img", { src: shotSrc, alt: "当前页实时预览" })
              : h("div", { className: "ph" }, liveUp
                ? "引擎在线，暂无页面快照。让模型调用 ppt_new / ppt_build 开始，这里会实时跟着变。"
                : "引擎离线。检查 /ppt 命令输出，或确认已安装 pywin32 + Pillow。"),
            deck ? h("span", { className: "dsh-ppt-badge" },
              "第 " + (deck.currentSlide || 1) + " / " + deck.slideCount + " 页") : null,
            liveUp && writing ? h("span", { className: "dsh-ppt-live-tag" }, "LIVE") : null
          ),
          deck ? h("div", { className: "dsh-ppt-meta" },
            h("span", null, "主题 ", h("b", null, deck.theme || "-")),
            h("span", null, "画布 ", h("b", null, deck.width + "×" + deck.height)),
            h("span", null, "可见 ", h("b", null, deck.visible ? "是" : "否")),
            h("span", null, "已保存 ", h("b", null, deck.savedTo ? "是" : "否"))
          ) : null,
          ops.length ? h("div", { className: "dsh-ppt-ops" },
            ops.slice(0, 14).map(function (o) {
              return h("div", { className: "dsh-ppt-op", key: o.seq },
                h("span", { className: "k" }, fmtTime(o.t)),
                h("span", { className: "p" }, "p" + o.slide),
                h("span", { className: "d" }, o.kind + (o.detail ? " · " + o.detail : ""))
              );
            })
          ) : null,
          h("div", { className: "dsh-ppt-acts" },
            h("button", { className: "dsh-ppt-act", onClick: function () { refreshLive(); refreshStatus(); } }, "刷新"),
            h("button", { className: "dsh-ppt-act", onClick: function () { setOpen(false); } }, "收起")
          ),
          error ? h("div", { className: "dsh-ppt-err" }, error) : null,
          h("div", { className: "dsh-ppt-hint" },
            "模型侧工具：ppt_*（建/画/存/校验）· zotero_*（文献）；斜杠命令 /ppt 自检。" +
            "PowerPoint 窗口本身也在同步实时绘制。")
        ) : null;

        return h("span", { className: "dsh-ppt-wrap", ref: rootRef },
          h("button", {
            type: "button", className: "dsh-ppt-btn", title: title,
            "aria-label": "dsh-ppt 实时预览",
            onClick: function () { setOpen(!open); }
          }, h(PptIcon), h("span", { className: dotClass })),
          pop
        );
      };
    }

    function apply(ctx) {
      // 防御性外壳：client 半跑在渲染进程里，任何异常都只降级为「不显示芯片」，
      // 绝不能冒出去影响整个页面（渲染进程崩溃会触发 watchdog 杀页面）。
      try {
        return applyInner(ctx);
      } catch (e) {
        try { console.warn("[dsh-ppt] client apply 异常，已降级为不显示芯片：" + e); } catch (e2) { /* 忽略 */ }
        return function () {};
      }
    }

    function applyInner(ctx) {
      var disposers = [];
      disposers.push(
        ctx.slots.inject("conversation.input.right", function () {
          return ctx.slots.register(
            {
              name: "conversation.input.right",
              id: "dsh-ppt",
              order: 110,
              label: "dsh-ppt 实时预览"
            },
            makeChip()
          );
        })
      );
      return function () {
        for (var i = 0; i < disposers.length; i++) {
          try { if (typeof disposers[i] === "function") disposers[i](); } catch (e) { /* 忽略 */ }
        }
      };
    }

    var inject = ["slots"];
    exports.apply = apply;
    exports.inject = inject;
    exports.name = "dsh-ppt";
    return module.exports;
  }
});
