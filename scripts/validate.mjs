/**
 * dsh-ppt 端到端校验 —— 用官方 @modelcontextprotocol/sdk 当客户端，
 * 验证自研零依赖 MCP stdio 服务端（pptx_engine/mcp_stdio.py）协议正确，
 * 并跑通「新建 → JSON 建整份 → 保存 → 校验 → 关闭」全链路 + 实时预览侧车。
 *
 * 跑法：node scripts/validate.mjs [--headless]
 * 依赖解析：@modelcontextprotocol/sdk 在 profile 的 node_modules 里（dsh-copilot 装的）。
 */
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const PLUGIN = path.resolve(HERE, "..");
const SERVER = path.join(PLUGIN, "engine", "server.py");
const HEADLESS = process.argv.includes("--headless");

const results = [];
function record(ok, name, detail) {
  results.push({ ok, name, detail: String(detail || "") });
  console.log(`${ok ? "[PASS]" : "[FAIL]"} ${name}${detail ? "  " + detail : ""}`);
}
async function check(name, fn) {
  const t0 = Date.now();
  try {
    const d = await fn();
    record(true, name, `${d || ""} (${((Date.now() - t0) / 1000).toFixed(1)}s)`);
    return true;
  } catch (e) {
    record(false, name, `${e && e.message ? e.message : e}`);
    return false;
  }
}

function detectPython() {
  const cands = [
    path.join(os.homedir(), "AppData", "Local", "hermes", "venv", "Scripts", "python.exe"),
    path.join(os.homedir(), "AppData", "Local", "Programs", "Python", "Python312", "python.exe")
  ];
  for (const c of cands) if (fs.existsSync(c)) return c;
  const r = spawnSync("where.exe", ["python"], { encoding: "utf8", windowsHide: true });
  return String(r.stdout || "").split(/\r?\n/)[0].trim() || "python";
}

const PYTHON = detectPython();
const WORKDIR = path.join(os.homedir(), "Desktop", "dsh-ppt-out", "validate");
fs.mkdirSync(WORKDIR, { recursive: true });

let client = null;
let transport = null;
let livePort = 8790;

async function call(name, args) {
  const r = await client.callTool({ name, arguments: args || {} });
  const text = (r.content || []).map((c) => c.text || "").join("\n");
  let json = null;
  try { json = JSON.parse(text); } catch (e) { /* 非 JSON 文本 */ }
  if (r.isError) throw new Error(`工具 ${name} 返回 isError：${text.slice(0, 300)}`);
  return json !== null ? json : { text };
}

async function main() {
  console.log(`=== dsh-ppt MCP 校验 ===\npython=${PYTHON}\nserver=${SERVER}\nheadless=${HEADLESS}\nworkdir=${WORKDIR}\n`);

  // ---- 1. 握手 ----
  const okHandshake = await check("MCP 握手（initialize + tools/list）", async () => {
    transport = new StdioClientTransport({
      command: PYTHON,
      args: [SERVER],
      env: {
        PYTHONIOENCODING: "utf-8",
        PYTHONUTF8: "1",
        DSH_PPT_WORKDIR: WORKDIR,
        DSH_PPT_LIVE_PORT: String(livePort),
        DSH_PPT_VISIBLE: HEADLESS ? "0" : "1",
        DSH_PPT_BACKEND: "auto",
        DSH_PPT_AUTO_PDF: "1",
        DSH_PPT_QUIT_ON_EXIT: "0"
      },
      stderr: "pipe"
    });
    client = new Client({ name: "dsh-ppt-validate", version: "1.0.0" }, { capabilities: {} });
    await client.connect(transport);
    const tools = await client.listTools();
    const names = tools.tools.map((t) => t.name);
    for (const must of ["ppt_status", "ppt_new", "ppt_build", "ppt_elements", "ppt_chart",
      "ppt_table", "ppt_save", "ppt_verify", "ppt_quality", "ppt_snapshot", "ppt_close",
      "zotero_status", "zotero_search", "zotero_cite", "zotero_refs_slide"]) {
      if (!names.includes(must)) throw new Error(`缺工具 ${must}`);
    }
    return `${names.length} 个工具`;
  });
  if (!okHandshake) return finish();

  // ---- 2. ppt_status ----
  let status = null;
  await check("ppt_status", async () => {
    status = await call("ppt_status", {});
    if (!status.ok) throw new Error("status.ok=false");
    const avail = (status.backends || []).filter((b) => b.available);
    if (!avail.length) throw new Error("无可用 Office 后端");
    livePort = (status.live && status.live.port) || livePort;
    return `后端=${avail[0].id} 元素=${status.elementTypes.length} 主题=${status.themes.length} live=${status.live.url}`;
  });

  // ---- 3. 实时预览侧车 ----
  await check("实时预览侧车（/health /state /snapshot）", async () => {
    const h = await fetch(`http://127.0.0.1:${livePort}/health`, { signal: AbortSignal.timeout(4000) });
    if (!h.ok) throw new Error(`/health ${h.status}`);
    const hj = await h.json();
    if (!hj.ok) throw new Error("health.ok=false");
    const s = await fetch(`http://127.0.0.1:${livePort}/state`, { signal: AbortSignal.timeout(4000) });
    const sj = await s.json();
    if (!("deck" in sj)) throw new Error("/state 缺 deck 字段");
    return `port=${livePort} engine=${hj.engine || "?"}`;
  });

  // ---- 4. 新建 ----
  await check("ppt_new（主题 default）", async () => {
    const r = await call("ppt_new", { theme: "default", visible: !HEADLESS, workdir: WORKDIR });
    if (!r.ok) throw new Error(r.error || "new 失败");
    return `deck=${r.deck.slideCount} 页 theme=${r.deck.theme}`;
  });

  // ---- 5. pre-flight 质量审查 ----
  await check("ppt_quality（pre-flight 拦错）", async () => {
    const r = await call("ppt_quality", {
      deckSpec: { slides: [{ elements: [
        { type: "header", sec: "1", title: "测试" },
        { type: "rect", box: [16, 500, 928, 80], fill: "#ccc" }
      ] }] }
    });
    const rules = new Set((r.report.issues || []).map((i) => i.rule));
    if (!rules.has("canvas-overflow")) throw new Error(`没抓到溢出：${[...rules]}`);
    return `score=${r.report.overallScore} rules=${[...rules].join(",")}`;
  });

  // ---- 6. JSON 建整份 ----
  await check("ppt_build（2 页：封面 + 图表/表格/卡片）", async () => {
    const r = await call("ppt_build", {
      slides: [
        { elements: [{ type: "cover", title: "校验演示", subtitle: "dsh-ppt validate",
          meta: "2026-10-28", titleFs: 40 }] },
        { elements: [
          { type: "header", sec: "1.1", title: "数据总览" },
          { type: "stat_cards", box: [16, 84, 928, 96], items: [
            { value: "98.1%", label: "合格率" }, { value: "12", label: "改善项" }] },
          { type: "chart", chart: "bar", box: [24, 196, 440, 260],
            data: { labels: ["A", "B", "C"], series: [{ name: "v", values: [3, 7, 5] }] } },
          { type: "table", box: [488, 196, 448, 200], header: ["项", "值"],
            rows: [["甲", "1"], ["乙", "2"], ["丙", "3"]] },
          { type: "tagline_bar", box: [16, 478, 928, 34], text: "校验通过" }
        ] }
      ]
    });
    if (!r.ok) throw new Error("build 失败");
    if (r.built.failed > 0) throw new Error(`元素失败 ${r.built.failed}：${JSON.stringify(r.built)}`);
    if (r.slideCount !== 2) throw new Error(`页数 ${r.slideCount} != 2`);
    return `2 页 / 元素成功 ${r.built.ok} / preflight=${r.preflight ? r.preflight.overallScore : "n/a"}`;
  });

  // ---- 7. 快照（实时预览的数据源）----
  await check("ppt_snapshot + 侧车快照一致", async () => {
    const r = await call("ppt_snapshot", { slide: 2 });
    if (!r.ok || !fs.existsSync(r.path)) throw new Error("快照未落盘");
    const sz = fs.statSync(r.path).size;
    if (sz < 3000) throw new Error(`快照太小 ${sz}`);
    const resp = await fetch(`http://127.0.0.1:${livePort}/snapshot.png`, { signal: AbortSignal.timeout(4000) });
    const buf = Buffer.from(await resp.arrayBuffer());
    if (buf.slice(0, 8).toString("hex") !== "89504e470d0a1a0a") throw new Error("侧车返回的不是 PNG");
    return `快照=${(sz / 1024).toFixed(0)}KB 侧车=${(buf.length / 1024).toFixed(0)}KB`;
  });

  // ---- 8. 操作流水 ----
  await check("操作流水（/ppt/ops 实时字幕）", async () => {
    const r = await fetch(`http://127.0.0.1:${livePort}/ops?since=0`, { signal: AbortSignal.timeout(4000) });
    const j = await r.json();
    if (!j.ops || j.ops.length < 3) throw new Error(`流水太少：${(j.ops || []).length}`);
    const kinds = [...new Set(j.ops.map((o) => o.kind))];
    return `${j.ops.length} 条，类型=${kinds.join("/")}`;
  });

  // ---- 9. 保存 + 校验 ----
  let savedPdf = null;
  await check("ppt_save（pptx+pdf+落盘校验）", async () => {
    const r = await call("ppt_save", { path: path.join(WORKDIR, "validate-deck.pptx"), verify: true });
    if (!r.ok || !r.pptx || r.pptx_size < 5000) throw new Error(`pptx 异常：${JSON.stringify(r).slice(0, 200)}`);
    if (!r.pdf || r.pdf_size < 2000) throw new Error(`pdf 异常：${JSON.stringify(r).slice(0, 200)}`);
    savedPdf = r.pdf;
    const ov = r.verify ? r.verify.overflow : -1;
    return `pptx=${(r.pptx_size / 1024).toFixed(0)}KB pdf=${(r.pdf_size / 1024).toFixed(0)}KB 溢出=${ov}`;
  });

  await check("ppt_verify（PDF 实测 + parity）", async () => {
    const r = await call("ppt_verify", { pdf: savedPdf });
    if (!r.ok) throw new Error("verify 失败");
    if (r.pdf && r.pdf.pages !== 2) throw new Error(`PDF 页数 ${r.pdf.pages} != 2`);
    return `溢出=${r.pdf.overflow} 重叠=${r.pdf.overlaps} parity=${r.review.dimensions.parity.score}`;
  });

  // ---- 10. Zotero 降级 ----
  await check("zotero_status（未安装时优雅降级）", async () => {
    const r = await call("zotero_status", {});
    if (r.sqliteExists) return `已安装，条目 ${r.itemCount}`;
    if (!r.error || !String(r.error).includes("zotero.sqlite")) throw new Error(`降级提示不可操作：${r.error}`);
    const s = await call("zotero_search", { query: "x" });
    if (s.available !== false) throw new Error("未安装时 search 应 available=false");
    return "未安装，降级正常";
  });

  // ---- 11. 关闭 ----
  await check("ppt_close（不杀 PowerPoint）", async () => {
    const r = await call("ppt_close", { quitApp: false });
    if (!r.ok) throw new Error("close 失败");
    return `关闭 ${r.closedSlides} 页`;
  });

  return finish();
}

function finish() {
  const npass = results.filter((r) => r.ok).length;
  const nfail = results.length - npass;
  console.log("\n" + "=".repeat(60));
  console.log(`MCP 校验结果：${npass} 通过 / ${nfail} 失败 / 共 ${results.length} 项`);
  console.log("=".repeat(60));
  for (const r of results) if (!r.ok) console.log(`  ✗ ${r.name}: ${r.detail}`);
  return nfail === 0 ? 0 : 1;
}

let code = 1;
try {
  code = await main();
} catch (e) {
  record(false, "未捕获异常", e && e.stack ? e.stack.split("\n").slice(0, 3).join(" | ") : e);
  code = 1;
} finally {
  try { if (client) await client.close(); } catch (e) { /* 忽略 */ }
  try { if (transport) await transport.close(); } catch (e) { /* 忽略 */ }
}
process.exit(code);
