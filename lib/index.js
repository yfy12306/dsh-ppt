/**
 * dsh-ppt（host 半）：把「真实 PowerPoint 做 PPT」的引擎挂进 DSH。
 *
 * 两块能力面：
 *   1) MCP 工具面 ppt_* / zotero_*（经 @deepseek-ai/dsh-mcp-client，stdio 起 Python 引擎）
 *      —— 引擎是常驻进程：COM 对象不能跨进程，「新建→逐页画→保存」必须同一进程完成。
 *   2) 实时预览面 /ppt/*（同源路由，仅回环）→ 代理到引擎内的只读 HTTP 侧车
 *      —— 客户端面板轮询它，形成「PowerPoint 窗口 + DSH 缩略图 + 操作字幕」三路实时反馈。
 *
 * 与 dsh-copilot 的挂载方式一致（ctx.plugin(mcpClient, ...) + ctx.webServer.register +
 * ctx.commands.register + skills.register），零改 DSH 核心。
 *
 * 刻意不做的事（相对上游 harness-anything 的安全修正）：
 *   - 不 taskkill PowerPoint —— 那会毁掉用户未保存的文档；引擎只关自己建的文稿。
 *   - 不在 host 侧缓存任何 PPT 内容 —— 状态全部来自引擎侧车，host 只做代理。
 */
import { createRequire } from "node:module";
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import net from "node:net";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import z from "@deepseek-ai/schemastery";
import * as mcpClient from "@deepseek-ai/dsh-mcp-client";

const name = "dsh-ppt";
const inject = ["webServer", "commands"];

const DEFAULTS = {
  enabled: true,
  pythonExe: "",
  serverScript: "",
  livePort: 8790,
  workdir: "",
  visible: true,
  backend: "auto",
  theme: "default",
  autoPdf: true,
  snapshotScale: 960,
  quitOnExit: false,
  skillEnabled: true,
  routePrefix: "/ppt",
  toolCallTimeoutMs: 300000,   // 建整份 PPT 可能几十秒，给足超时
  liveProbeMs: 4000
};

const Config = z.object({
  enabled: z.boolean().default(true),
  pythonExe: z.string().default(""),
  serverScript: z.string().default(""),
  livePort: z.number().default(8790),
  workdir: z.string().default(""),
  visible: z.boolean().default(true),
  backend: z.string().default("auto"),
  theme: z.string().default("business"),
  autoPdf: z.boolean().default(true),
  snapshotScale: z.number().default(960),
  quitOnExit: z.boolean().default(false),
  skillEnabled: z.boolean().default(true),
  routePrefix: z.string().default("/ppt"),
  toolCallTimeoutMs: z.number().default(300000),
  liveProbeMs: z.number().default(4000)
});

// ---------------------------------------------------------------- 探测
function detectPythonExe() {
  // 优先级：显式配置 > 插件自带 venv > hermes venv（本机实测带全套依赖）> 基础 Python312 > PATH
  const pluginDir = path.dirname(fs.realpathSync(fileURLToPath(import.meta.url)));
  const cands = [
    path.join(pluginDir, "..", ".venv", "Scripts", "python.exe"),
    path.join(os.homedir(), "AppData", "Local", "hermes", "venv", "Scripts", "python.exe"),
    path.join(os.homedir(), "AppData", "Local", "Programs", "Python", "Python312", "python.exe"),
    path.join(os.homedir(), "AppData", "Local", "Programs", "Python", "Python313", "python.exe")
  ];
  for (const c of cands) {
    try { if (fs.existsSync(c)) return path.resolve(c); } catch (e) { /* 忽略 */ }
  }
  try {
    const r = spawnSync("where.exe", ["python"], { encoding: "utf8", windowsHide: true, timeout: 8000 });
    const first = String(r.stdout || "").split(/\r?\n/)[0].trim();
    if (first && fs.existsSync(first)) return first;
  } catch (e) { /* 忽略 */ }
  return "python";
}

/** 校验解释器是否带引擎必需依赖（pywin32 + Pillow）。缺了给出可操作提示而不是静默失败。 */
function probePythonDeps(exe) {
  try {
    const r = spawnSync(exe, ["-c",
      "import sys;ok=[];miss=[]\n" +
      "for m in ('win32com.client','PIL'):\n" +
      "    (ok if __import__('importlib').util.find_spec(m.split('.')[0]) else miss).append(m)\n" +
      "import json;print(json.dumps({'python':sys.version.split()[0],'ok':ok,'miss':miss}))"
    ], { encoding: "utf8", windowsHide: true, timeout: 20000 });
    const line = String(r.stdout || "").split(/\r?\n/).filter(Boolean).pop();
    if (line) return JSON.parse(line);
  } catch (e) { /* 忽略 */ }
  return { python: "", ok: [], miss: ["win32com.client", "PIL"], error: "探测失败" };
}

function portAlive(port, ms) {
  return new Promise(function (resolve) {
    const sock = net.connect({ port: port, host: "127.0.0.1" });
    const done = function (v) { try { sock.destroy(); } catch (e) { /* 忽略 */ } resolve(v); };
    sock.setTimeout(ms);
    sock.once("connect", function () { done(true); });
    sock.once("timeout", function () { done(false); });
    sock.once("error", function () { done(false); });
  });
}

async function fetchTimeout(url, ms, options) {
  const ac = new AbortController();
  const t = setTimeout(function () { ac.abort(); }, ms);
  try {
    return await fetch(url, Object.assign({ signal: ac.signal }, options || {}));
  } finally {
    clearTimeout(t);
  }
}

function isLoopbackAddr(addr) {
  if (!addr) return false;
  const a = String(addr);
  return a === "127.0.0.1" || a === "::1" || a === "::ffff:127.0.0.1" || a.startsWith("127.");
}

function sendJson(res, code, obj) {
  const body = Buffer.from(JSON.stringify(obj), "utf8");
  res.writeHead(code, {
    "content-type": "application/json; charset=utf-8",
    "content-length": String(body.length),
    "cache-control": "no-store"
  });
  res.end(body);
}

function sendBytes(res, code, buf, ctype) {
  res.writeHead(code, {
    "content-type": ctype,
    "content-length": String(buf.length),
    "cache-control": "no-store, max-age=0"
  });
  res.end(buf);
}

// ---------------------------------------------------------------- 主逻辑
async function apply(ctx, config) {
  // 防御性外壳（2026-10-28 教训）：插件自身任何异常都只记日志、降级为「不挂载」，
  // 绝不让异常冒到 profile-composition 阶段把整个 DSH 启动拖进恢复模式。
  try {
    return await applyInner(ctx, config);
  } catch (e) {
    try {
      console.error("[dsh-ppt] apply 异常，已降级为不挂载（不影响 DSH 其它功能）：" +
        String(e && e.stack ? e.stack : e));
    } catch (e2) { /* 忽略 */ }
  }
}

async function applyInner(ctx, config) {
  const cfg = Object.assign({}, DEFAULTS, config || {});
  const pluginDir = path.dirname(fs.realpathSync(fileURLToPath(import.meta.url)));
  const serverScript = cfg.serverScript
    ? path.resolve(cfg.serverScript)
    : path.join(pluginDir, "..", "engine", "server.py");
  const pythonExe = cfg.pythonExe ? path.resolve(cfg.pythonExe) : detectPythonExe();
  const deps = probePythonDeps(pythonExe);
  const workdir = cfg.workdir
    ? path.resolve(cfg.workdir)
    : path.join(os.homedir(), "Desktop", "dsh-ppt-out");

  const state = {
    faceMounted: false,
    liveUp: false,
    livePort: cfg.livePort,
    lastProbeAt: 0,
    lastError: ""
  };

  const serverFound = fs.existsSync(serverScript);
  const depsOk = deps.miss.length === 0;

  // ---- 1) MCP 工具面 ----
  if (cfg.enabled && serverFound && depsOk) {
    try {
      ctx.plugin(mcpClient, {
        transport: "stdio",
        serverName: "ppt",
        command: pythonExe,
        args: [serverScript],
        env: {
          PYTHONIOENCODING: "utf-8",
          PYTHONUTF8: "1",
          DSH_PPT_WORKDIR: workdir,
          DSH_PPT_LIVE_PORT: String(cfg.livePort),
          DSH_PPT_VISIBLE: cfg.visible ? "1" : "0",
          DSH_PPT_BACKEND: cfg.backend,
          DSH_PPT_AUTO_PDF: cfg.autoPdf ? "1" : "0",
          DSH_PPT_SNAPSHOT_SCALE: String(cfg.snapshotScale),
          DSH_PPT_QUIT_ON_EXIT: cfg.quitOnExit ? "1" : "0"
        },
        toolCallTimeoutMs: cfg.toolCallTimeoutMs,
        failOnStartupError: false
      });
      state.faceMounted = true;
      console.info("[dsh-ppt] MCP 面已挂载：python=" + pythonExe + " server=" + serverScript);
    } catch (e) {
      state.lastError = String(e && e.message ? e.message : e);
      console.warn("[dsh-ppt] MCP 面挂载失败（不影响 DSH 其它功能）：" + state.lastError);
    }
  } else if (cfg.enabled) {
    state.lastError = !serverFound
      ? "引擎脚本不存在：" + serverScript
      : "Python 缺依赖：" + deps.miss.join(", ") + "（pip install pywin32 Pillow）";
    console.warn("[dsh-ppt] 未挂载：" + state.lastError);
  }

  // ---- 2) 实时预览侧车探活（周期）----
  async function probeLive() {
    state.lastProbeAt = Date.now();
    for (const p of [cfg.livePort, cfg.livePort + 1, cfg.livePort + 2]) {
      if (await portAlive(p, 700)) {
        try {
          const r = await fetchTimeout("http://127.0.0.1:" + p + "/health", 1200);
          if (r.ok) {
            state.liveUp = true;
            state.livePort = p;
            return true;
          }
        } catch (e) { /* 换下一个端口 */ }
      }
    }
    state.liveUp = false;
    return false;
  }

  let probeTimer = null;
  if (cfg.enabled) {
    probeTimer = setInterval(function () {
      probeLive().catch(function () { /* 忽略 */ });
    }, Math.max(1500, cfg.liveProbeMs));
    if (probeTimer.unref) probeTimer.unref();
    probeLive().catch(function () { /* 忽略 */ });
  }

  // ---- 3) 同源路由 /ppt/*（仅回环）→ 代理到引擎侧车 ----
  const guard = function (req, res) {
    if (!isLoopbackAddr(req.socket && req.socket.remoteAddress)) {
      sendJson(res, 403, { error: "loopback only" });
      return false;
    }
    return true;
  };

  async function proxyToLive(req, res, subPath) {
    if (!guard(req, res)) return;
    const alive = state.liveUp || (await probeLive());
    if (!alive) {
      sendJson(res, 503, {
        ok: false, error: "实时预览服务未就绪（引擎进程未启动或端口未开）",
        faceMounted: state.faceMounted, python: pythonExe, deps: deps,
        expectPort: cfg.livePort, lastError: state.lastError
      });
      return;
    }
    const url = "http://127.0.0.1:" + state.livePort + subPath;
    try {
      const r = await fetchTimeout(url, 8000);
      const ctype = r.headers.get("content-type") || "application/octet-stream";
      const buf = Buffer.from(await r.arrayBuffer());
      sendBytes(res, r.status, buf, ctype);
    } catch (e) {
      sendJson(res, 502, { ok: false, error: "代理到引擎侧车失败：" + String(e && e.message ? e.message : e) });
    }
  }

  const statusRoute = async function (req, res) {
    if (!guard(req, res)) return;
    const alive = state.liveUp || (await probeLive());
    let deck = null;
    if (alive) {
      try {
        const r = await fetchTimeout("http://127.0.0.1:" + state.livePort + "/state", 2500);
        if (r.ok) deck = (await r.json()).deck || null;
      } catch (e) { /* 忽略 */ }
    }
    sendJson(res, 200, {
      ok: true,
      plugin: { name: name, version: "1.0.0" },
      face: { mounted: state.faceMounted, serverScript: serverScript, serverFound: serverFound },
      python: { exe: pythonExe, version: deps.python || "", depsOk: depsOk, missing: deps.miss },
      live: { up: alive, port: state.livePort, expectPort: cfg.livePort,
              routes: { state: cfg.routePrefix + "/state", snapshot: cfg.routePrefix + "/snapshot.png",
                        ops: cfg.routePrefix + "/ops", combined: cfg.routePrefix + "/live" } },
      config: { visible: cfg.visible, backend: cfg.backend, theme: cfg.theme,
                workdir: workdir, autoPdf: cfg.autoPdf, snapshotScale: cfg.snapshotScale },
      deck: deck,
      lastError: state.lastError,
      ui: { pollIntervalMs: 60000, popupPollIntervalMs: 1200 }
    });
  };

  ctx.effect(function* () {
    const disposers = [
      ctx.webServer.register({ kind: "exact", path: cfg.routePrefix + "/status", handler: statusRoute }),
      ctx.webServer.register({
        kind: "prefix", path: cfg.routePrefix + "/",
        handler: async function (req, res) {
          const sub = String(req.url || "").replace(cfg.routePrefix, "") || "/";
          await proxyToLive(req, res, sub);
        }
      })
    ];
    yield function () {
      for (const d of disposers) { try { if (typeof d === "function") d(); } catch (e) { /* 忽略 */ } }
    };
  }, "dsh-ppt: routes");

  // ---- 4) /ppt 斜杠命令 ----
  ctx.effect(function* () {
    yield ctx.commands.register({
      name: "ppt",
      description: "PPT 引擎状态自检（dsh-ppt：Office 后端 / 实时预览 / 当前文稿）",
      handler: async function () {
        try {
          const alive = state.liveUp || (await probeLive());
          const mark = function (v) { return v ? "✓" : "✗"; };
          const lines = [
            "dsh-ppt 状态：",
            "  MCP 面   " + mark(state.faceMounted) + " " +
              (state.faceMounted ? "ppt_* / zotero_* 工具已挂载" : "未挂载（" + (state.lastError || "disabled") + "）"),
            "  Python   " + mark(depsOk) + " " + pythonExe +
              (deps.python ? "（" + deps.python + "）" : "") +
              (depsOk ? "" : "  缺：" + deps.miss.join(", ")),
            "  实时预览 " + mark(alive) + " http://127.0.0.1:" + state.livePort +
              "（面板路由 " + cfg.routePrefix + "/live）",
            "  可见模式 " + mark(cfg.visible) + (cfg.visible ? " PowerPoint 窗口实时可见" : " 后台生成"),
            "  后端     " + cfg.backend + "　主题 " + cfg.theme + "　工作目录 " + workdir
          ];
          if (alive) {
            try {
              const r = await fetchTimeout("http://127.0.0.1:" + state.livePort + "/state", 2500);
              if (r.ok) {
                const st = (await r.json()).deck;
                if (st) {
                  lines.push("  当前文稿 " + st.slideCount + " 页，第 " + st.currentSlide +
                    " 页，已保存：" + (st.savedTo || "否"));
                  const ops = (st.ops || []).slice(0, 3);
                  for (const o of ops) {
                    lines.push("    · [" + o.kind + "] p" + o.slide + " " + o.detail);
                  }
                } else {
                  lines.push("  当前文稿 无（先调 ppt_new）");
                }
              }
            } catch (e) { /* 忽略 */ }
          }
          return { kind: "success", text: lines.join("\n") };
        } catch (error) {
          return { kind: "error", text: "状态自检失败：" + String(error && error.message ? error.message : error) };
        }
      }
    });
  }, "dsh-ppt: command");

  // ---- 5) Skill 注册（让模型知道何时用这套工具）----
  const SKILL_CONTENT = [
    "# dsh-ppt（真实 PowerPoint 做 PPT + 实时可见）",
    "",
    "用本机真实 Microsoft PowerPoint（或 WPS）COM 生成 PPT。用户在 PowerPoint 窗口和 DSH 预览面板里能实时看到每一步。",
    "",
    "## 标准流程（按此顺序）",
    "1. `ppt_status` —— 看后端/主题/元素类型/实时预览地址。任何 PPT 任务先调它。",
    "2. `ppt_new` —— 选风格建文稿：default（干净商务，默认）。",
    "3. `ppt_build` —— 传 slides[] JSON 一次建整份（主力）。每页 elements[] 用声明式元素。",
    "   或 `ppt_slide` + `ppt_elements` 逐页精修。",
    "4. `ppt_quality` —— pre-flight 审查（画布溢出/标题间距/字号/中文引号/密度）。有 error 先修再存。",
    "5. `ppt_save` —— 出 pptx + pdf，自动做落盘校验。",
    "6. `ppt_verify` —— PDF 实测溢出/重叠/空页 + PPTX/PDF 页数一致性。",
    "",
    "## 元素类型（ppt_build 的 elements[].type）",
    "结构：cover / toc / header / divider(章节分隔页) / footer / tagline_bar / bullets / quote / kpi",
    "数据：stat_cards / cards / cards_2x3 / card_list_wide / timeline / table / chart",
    "基础：text / rich / rect / rrect / oval / shape / arrow / hline / vline / pic / pill / chip / num_circle / grad_bar / gray_box",
    "",
    "## 设计系统（已内置，遵循即专业；依据 WCAG / 60-30-10 / 咨询与学术版式规范）",
    "- 排版：每主题两套字体族（标题/正文）+ 数字族 + 等宽族，字号走 scale",
    "  （display 44-48 / h1 32 / h2 20 / body 18-20 / small 13-14 / numeric 40-46 / caption 10-12）。",
    "  学术=黑体标题+宋体正文+Times 数字；政务=黑体+仿宋；商务=雅黑+等线。不要手动乱改字号。",
    "- 配色：60% 背景 / 30% 主题色（标题条、卡片、图表主体）/ 10% 强调色（只给大数字、",
    "  单条图表系列、高亮箭头）。图表分类色默认 Okabe-Ito 色盲安全板，勿红绿单靠颜色区分。",
    "- 对比度：正文 ≥4.5:1、大字 ≥3:1（WCAG）。ppt_quality 会实测并拦下不达标组合。",
    "- 版式：一页一个观点；正文左对齐（只有标题可居中）；页边距≈5%；留白 40-50%；",
    "  数据页让图表占 70-80%；标题写结论（action title）而不是「Q4 收入」这类描述。",
    "- 页脚：默认自动加（左 deck 名 + 右页码，10-12pt）。封面与 divider 页传 footer=false。",
    "- 章节分隔页：用 divider，大编号+≤5词大标题+一句副题，其余留白。",
    "",
    "## 动画（克制为原则）",
    "- 整页切换默认 Fade 0.6s（部分主题 push）；无切换会像「仓促拼凑的 PDF」。",
    "- 元素入场默认 Fade 错落 0.3s/个；线/箭头想「画出来」用 wipe。",
    "- 不要复杂动画。需要关闭：ppt_new(entrance=\"none\")。",
    "",
    "## 布局硬约束（违反会被 ppt_quality 拦下）",
    "- 画布 960x540，所有元素 y+h ≤ 518",
    "- 标题栏结束于 y≈54，第一个内容元素必须 y ≥ 76（间距 ≥24pt）",
    "- 一页一个 header；正文 ≥16pt（投影建议 ≥18-20pt）",
    "- 中文文本内引用用「」，不要用英文双引号",
    "",
    "## 图表",
    "`ppt_chart` 支持 bar/hbar/line/area/pie/donut/scatter/bubble。data={labels:[],series:[{name,values:[]}]}。",
    "小扇区饼图会自动改用图例防标签重叠。中文字体已配好不会出方块。",
    "",
    "## 文献（Zotero，可选）",
    "`zotero_search` 检索 → `zotero_cite` 渲染 GB/T 7714/APA/IEEE → `zotero_refs_slide` 直接生成参考文献页。",
    "Zotero 未安装时这些工具返回 available=false，不影响 PPT 主流程。",
    "",
    "## 实时预览",
    "引擎内置只读 HTTP 侧车；DSH 面板路由 /ppt/live（缩略图+操作流水）。",
    "想让用户看某一页：`ppt_goto` 跳页 + `ppt_snapshot` 出图。",
    "",
    "## 安全边界",
    "- 绝不 taskkill PowerPoint；`ppt_close` 只关本插件建的文稿，默认不退出应用。",
    "- 新建文稿前若有未保存的旧文稿，会自动存到 <workdir>/recovery/ 再关。"
  ].join("\n");

  if (cfg.skillEnabled) {
    try {
      ctx.inject(["skills"], function (skillCtx) {
        try {
          skillCtx.effect(function* () {
            yield skillCtx.skills.register({
              name: "dsh-ppt",
              source: "runtime",
              description: "用真实 Microsoft PowerPoint/WPS 生成 PPT：JSON 数据驱动建整份、default 风格（支持 register_theme 自定义主题）、排版/动画/配色按 WCAG 与 60-30-10 等调研规范、matplotlib 图表、表格/卡片/时间轴/章节页/页脚、PDF 落盘校验、Zotero 参考文献页。当任务需要制作/修改演示文稿、汇报 PPT、答辩 PPT 时加载。",
              whenToUse: "需要制作、修改、导出 PPT/演示文稿，或需要把数据/文献排成幻灯片时。",
              content: SKILL_CONTENT
            });
          }, "dsh-ppt: skill");
          console.info("[dsh-ppt] Skill 已注册：dsh-ppt");
          Promise.resolve(skillCtx.skills.get("dsh-ppt")).then(function (def) {
            if (def === undefined) console.warn("[dsh-ppt] Skill 自检：get() 未返回定义");
            else console.info("[dsh-ppt] Skill 自检通过");
          }).catch(function (e) {
            console.warn("[dsh-ppt] Skill 自检失败：" + String(e && e.message ? e.message : e));
          });
        } catch (e) {
          console.warn("[dsh-ppt] Skill 注册失败（不影响工具面）：" + String(e && e.message ? e.message : e));
        }
      });
    } catch (e) {
      console.warn("[dsh-ppt] skills 服务注入失败（跳过 Skill 注册）：" + String(e && e.message ? e.message : e));
    }
  }

  // ---- 6) 清理 ----
  ctx.effect(function* () {
    yield function () {
      if (probeTimer) { clearInterval(probeTimer); probeTimer = null; }
      // 引擎进程由 mcp-client 随面卸载回收；这里不主动杀 Python，
      // 避免把用户正在看的 PowerPoint 一起带走（引擎 atexit 会自行收尾）。
    };
  }, "dsh-ppt: cleanup");

  console.info(
    "[dsh-ppt] 就绪：face=" + (state.faceMounted ? "mounted" : "off") +
    " python=" + (depsOk ? "ok" : "missing(" + deps.miss.join(",") + ")") +
    " live=" + cfg.livePort + " visible=" + cfg.visible +
    " theme=" + cfg.theme + " workdir=" + workdir
  );
}

export { Config, apply, inject, name };
