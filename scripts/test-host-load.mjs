/**
 * dsh-ppt 宿主半隔离加载测试 —— 不重启 DSH、不碰真实宿主。
 *
 * 做法：把 lib/index.js 复制进临时目录，在临时目录里放 @deepseek-ai/* 的桩包，
 * 用假 ctx 调 apply()，断言：
 *   1) apply 不抛异常（2026-10-28 的教训：插件异常绝不能冒到 profile-composition）
 *   2) MCP 面被以正确的 command/args/env 挂载
 *   3) /ppt/* 路由、/ppt 命令、dsh-ppt Skill 都注册成功
 *   4) 路由 handler 在侧车离线时返回 503 而不是抛异常
 *   5) 故意让 ctx.plugin 抛异常时，apply 仍不冒异常（防御外壳生效）
 *
 * 跑法：node scripts/test-host-load.mjs
 */
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const PLUGIN = path.resolve(HERE, "..");

const results = [];
function record(ok, name, detail) {
  results.push({ ok, name, detail: String(detail || "") });
  console.log(`${ok ? "[PASS]" : "[FAIL]"} ${name}${detail ? "  " + detail : ""}`);
}
async function check(name, fn) {
  try { const d = await fn(); record(true, name, d || ""); return true; }
  catch (e) { record(false, name, e && e.stack ? e.stack.split("\n").slice(0, 2).join(" | ") : e); return false; }
}

// ---------- 桩：@deepseek-ai/schemastery ----------
const chainable = new Proxy(function () {}, {
  get(_t, prop) {
    if (prop === "then") return undefined;           // 别被当 Promise
    return chainable;
  },
  apply() { return chainable; }
});
const schemasteryStub = { default: { object: () => chainable, boolean: () => chainable, string: () => chainable, number: () => chainable }, z: chainable };

// ---------- 桩：@deepseek-ai/dsh-mcp-client ----------
let mcpConfig = null;
let mcpShouldThrow = false;
const mcpClientStub = {
  __setThrow(v) { mcpShouldThrow = v; },
  apply(_ctx, cfg) {
    if (mcpShouldThrow) throw new Error("simulated mcp mount failure");
    mcpConfig = cfg;
  }
};

function makeTempWorkspace() {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "dsh-ppt-hostload-"));
  const nm = path.join(tmp, "node_modules", "@deepseek-ai");
  fs.mkdirSync(path.join(nm, "schemastery"), { recursive: true });
  fs.mkdirSync(path.join(nm, "dsh-mcp-client"), { recursive: true });
  fs.writeFileSync(path.join(nm, "schemastery", "package.json"),
    JSON.stringify({ name: "@deepseek-ai/schemastery", version: "0.0.0", type: "module", main: "index.js", exports: { ".": "./index.js" } }));
  fs.writeFileSync(path.join(nm, "schemastery", "index.js"),
    "const c=new Proxy(function(){},{get:(t,p)=>p==='then'?undefined:c,apply:()=>c});export default {object:()=>c,boolean:()=>c,string:()=>c,number:()=>c};export const z=c;");
  fs.writeFileSync(path.join(nm, "dsh-mcp-client", "package.json"),
    JSON.stringify({ name: "@deepseek-ai/dsh-mcp-client", version: "0.0.0", type: "module", main: "index.js", exports: { ".": "./index.js" } }));
  fs.writeFileSync(path.join(nm, "dsh-mcp-client", "index.js"),
    "export let last=null;export function apply(ctx,cfg){ if(globalThis.__mcpThrow) throw new Error('simulated mcp mount failure'); last=cfg; globalThis.__mcpLast=cfg; }");
  // 复制被测文件
  fs.copyFileSync(path.join(PLUGIN, "lib", "index.js"), path.join(tmp, "index.js"));
  return tmp;
}

function makeFakeCtx() {
  const reg = { routes: [], commands: [], skills: [], effects: [], plugins: [] };
  const ctx = {
    _reg: reg,
    plugin(mod, cfg) { reg.plugins.push({ mod, cfg }); if (globalThis.__mcpThrow) throw new Error("simulated mcp mount failure"); },
    webServer: { register(r) { reg.routes.push(r); return () => {}; } },
    commands: { register(c) { reg.commands.push(c); return () => {}; } },
    inject(_names, fn) { try { fn(ctx); } catch (e) { /* 与真实宿主一致：注入失败不冒 */ } },
    skills: {
      register(s) { reg.skills.push(s); return Promise.resolve(); },
      get(n) { return Promise.resolve(reg.skills.find((s) => s.name === n)); }
    },
    effect(fn, _label) {
      // 真实宿主会跑 generator；这里同步跑一遍以触发注册
      try {
        const g = fn();
        let r = g.next();
        while (!r.done) r = g.next(typeof r.value === "function" ? r.value : undefined);
      } catch (e) { reg.effectError = String(e); }
      return () => {};
    },
    on() { return () => {}; },
    credentials: {}
  };
  return ctx;
}

async function main() {
  const tmp = makeTempWorkspace();
  const mod = await import(pathToFileURL(path.join(tmp, "index.js")).href);

  await check("模块可导入且导出齐全", async () => {
    for (const k of ["apply", "name", "inject", "Config"]) {
      if (!(k in mod)) throw new Error(`缺导出 ${k}`);
    }
    return `name=${mod.name} inject=${mod.inject.join(",")}`;
  });

  let ctx = makeFakeCtx();
  // 注意：index.js 被复制进了临时目录，按自身位置推导的 server.py 不存在；
  // 这里显式传 serverScript 指向真实引擎，模拟宿主配置里指定路径的情形。
  const baseCfg = { enabled: true, livePort: 8890, visible: false,
    serverScript: path.join(PLUGIN, "engine", "server.py") };
  await check("apply() 不抛异常（正常路径）", async () => {
    await mod.apply(ctx, baseCfg);
    return `routes=${ctx._reg.routes.length} commands=${ctx._reg.commands.length} skills=${ctx._reg.skills.length}`;
  });

  await check("MCP 面以正确 command/args/env 挂载", async () => {
    const p = ctx._reg.plugins[0];
    if (!p) throw new Error("没有挂载 MCP 面");
    const c = p.cfg;
    if (c.transport !== "stdio") throw new Error("transport 非 stdio");
    if (c.serverName !== "ppt") throw new Error(`serverName=${c.serverName}`);
    if (!/python/.test(c.command)) throw new Error(`command 不是 python：${c.command}`);
    if (!c.args[0].endsWith("server.py")) throw new Error(`args[0]=${c.args[0]}`);
    if (!fs.existsSync(c.args[0])) throw new Error("server.py 不存在");
    if (c.env.DSH_PPT_LIVE_PORT !== "8890") throw new Error("livePort 未透传");
    return `command=${path.basename(c.command)} args=${c.args.join(" ")} env.LIVE=${c.env.DSH_PPT_LIVE_PORT}`;
  });

  await check("路由 /ppt/status + /ppt/ 前缀代理已注册", async () => {
    const paths = ctx._reg.routes.map((r) => r.path);
    if (!paths.includes("/ppt/status")) throw new Error(`缺 /ppt/status：${paths}`);
    if (!paths.some((p) => p.startsWith("/ppt/"))) throw new Error("缺前缀代理");
    return paths.join(", ");
  });

  await check("/ppt 命令 + dsh-ppt Skill 已注册", async () => {
    if (!ctx._reg.commands.some((c) => c.name === "ppt")) throw new Error("缺 /ppt 命令");
    const sk = ctx._reg.skills.find((s) => s.name === "dsh-ppt");
    if (!sk) throw new Error("缺 Skill");
    if (typeof sk.content !== "string" || sk.content.length < 200) throw new Error("Skill content 过短");
    if (typeof sk.source !== "string") throw new Error("Skill 缺 source（加载校验会拒）");
    return `skill content=${sk.content.length} 字符`;
  });

  await check("路由 handler 在侧车离线时返回 503 而非抛异常", async () => {
    const prefixRoute = ctx._reg.routes.find((r) => r.kind === "prefix");
    const fakeReq = { url: "/ppt/state", socket: { remoteAddress: "127.0.0.1" }, method: "GET" };
    let code = 0; let body = "";
    const fakeRes = {
      writeHead(c, _h) { code = c; },
      end(b) { body = b ? b.toString("utf8") : ""; }
    };
    await prefixRoute.handler(fakeReq, fakeRes);
    if (code !== 503 && code !== 200) throw new Error(`意外状态码 ${code}`);
    const j = JSON.parse(body);
    if (code === 503 && j.ok !== false) throw new Error("503 但 ok 不为 false");
    return `status=${code} ok=${j.ok}`;
  });

  await check("status 路由在侧车离线时返回 200 + 诊断信息", async () => {
    const st = ctx._reg.routes.find((r) => r.path === "/ppt/status");
    let code = 0; let body = "";
    const fakeRes = { writeHead(c) { code = c; }, end(b) { body = b.toString("utf8"); } };
    await st.handler({ socket: { remoteAddress: "127.0.0.1" } }, fakeRes);
    if (code !== 200) throw new Error(`status=${code}`);
    const j = JSON.parse(body);
    if (!j.python || !j.live) throw new Error("缺 python/live 诊断字段");
    return `python.depsOk=${j.python.depsOk} live.up=${j.live.up}`;
  });

  await check("非回环请求被 403 拒绝", async () => {
    const st = ctx._reg.routes.find((r) => r.path === "/ppt/status");
    let code = 0;
    const fakeRes = { writeHead(c) { code = c; }, end() {} };
    await st.handler({ socket: { remoteAddress: "203.0.113.7" } }, fakeRes);
    if (code !== 403) throw new Error(`期望 403，实际 ${code}`);
    return "403";
  });

  // ---- 防御外壳：让 ctx.plugin 抛异常，apply 仍不得冒异常 ----
  await check("ctx.plugin 抛异常时 apply 仍不冒异常（防御外壳）", async () => {
    globalThis.__mcpThrow = true;
    const ctx2 = makeFakeCtx();
    await mod.apply(ctx2, { enabled: true, livePort: 8891,
      serverScript: path.join(PLUGIN, "engine", "server.py") });   // 不应抛
    delete globalThis.__mcpThrow;
    // 面挂载失败但路由/命令仍应注册（降级而非全灭）
    if (ctx2._reg.routes.length === 0) throw new Error("降级后路由也没注册");
    return `routes=${ctx2._reg.routes.length}（面挂载失败但已降级）`;
  });

  // ---- 清理 ----
  try { fs.rmSync(tmp, { recursive: true, force: true }); } catch (e) { /* 忽略 */ }

  const npass = results.filter((r) => r.ok).length;
  const nfail = results.length - npass;
  console.log("\n" + "=".repeat(58));
  console.log(`宿主半隔离加载：${npass} 通过 / ${nfail} 失败 / 共 ${results.length} 项`);
  console.log("=".repeat(58));
  for (const r of results) if (!r.ok) console.log(`  ✗ ${r.name}: ${r.detail}`);
  return nfail === 0 ? 0 : 1;
}

process.exit(await main());
