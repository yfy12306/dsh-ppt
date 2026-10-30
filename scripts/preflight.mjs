/**
 * dsh-ppt 启动前一致性守卫 —— 防止「配置引用了未安装的包」再次把 DSH 拖进恢复模式。
 *
 * 背景（2026-10-28 事故）：profile 的 bundles / cordis.patch.yml 里写了 dsh-ppt，
 * 但 pnpm install 因 dsh-purge tarball 校验和漂移而中止，node_modules/dsh-ppt 不存在，
 * DSH 在 profile-composition 阶段抛 PackageOverlayNotFoundError 硬失败 → 恢复模式。
 *
 * 解析根与 DSH 一致（错误原文："from the Desktop installation or active Profile"）：
 *   R1 = <profile>/node_modules
 *   R2 = <DSH 安装目录>/resources/app/node_modules
 *
 * 检查项：
 *   1) package.json 的 dsh.profile.bundles 每个包名在 R1 或 R2 可解析
 *   2) package.json 的 link: 依赖在 R1 存在
 *   3) cordis.patch.yml 顶层 patch 条目（- id: 与紧随的 name:，缩进配对）中，
 *      非相对路径、非 cordis:* 、非含空格的显示名 的 name 在 R1 或 R2 可解析
 *   4) dsh-ppt 插件自身关键文件齐备、package.json 合法
 * 任何一项不满足 → 退出码 1。**不要在这种状态下重启 DSH。**
 *
 * 跑法：node scripts/preflight.mjs [profileDir] [dshInstallDir]
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const PROFILE = path.resolve(process.argv[2] || path.join(HERE, "..", ".."));
const DSH_INSTALL = path.resolve(process.argv[3] ||
  path.join(process.env.LOCALAPPDATA || "C:\\Users\\user\\AppData\\Local",
    "Programs", "DSH Desktop"));

const ROOTS = [
  path.join(PROFILE, "node_modules"),
  path.join(DSH_INSTALL, "resources", "app", "node_modules")
];

const problems = [];
const notes = [];

function resolvable(name) {
  if (name.startsWith("./") || name.startsWith(".\\") || name.startsWith("../")) {
    const ok = fs.existsSync(path.join(PROFILE, name));
    return { kind: "relative", ok };
  }
  // 子路径导出（pkg/sub）只校验包根是否存在 —— 包在就不会触发 overlay 硬失败
  let pkgName = name;
  if (name.includes("/")) {
    const parts = name.split("/");
    pkgName = name.startsWith("@") ? parts.slice(0, 2).join("/") : parts[0];
  }
  for (const root of ROOTS) {
    if (fs.existsSync(path.join(root, pkgName, "package.json"))) {
      return { kind: "node_modules", ok: true, root };
    }
  }
  return { kind: "node_modules", ok: false };
}

// 基线白名单：宿主内建 / disabled 条目等「DSH 自己能处理、但不在 node_modules」的名字。
// 守卫只对「不在白名单里的新不可解析项」报错 —— 这正是 2026-10-28 事故的形态。
const BASELINE = new Set([
  "dsh-plugin-desktop",                              // 宿主内建桌面壳，按名解析不走 node_modules
  "@deepseek-ai/dsh-tool-plugin-manager/tools"      // 配置里 disabled: true，宿主跳过解析
]);

function isPackageRef(name) {
  if (!name) return false;
  if (name.startsWith(".")) return true;              // 相对路径也要查
  if (name.startsWith("cordis:")) return false;      // 内建伪名
  if (/\s/.test(name)) return false;                  // 模型显示名等
  if (/\.(mjs|js|yml|yaml)$/.test(name)) return true; // 文件名引用
  return true;
}

// ---- 1) bundles + link 依赖 ----
try {
  const pj = JSON.parse(fs.readFileSync(path.join(PROFILE, "package.json"), "utf8"));
  const bundles = (pj.dsh && pj.dsh.profile && pj.dsh.profile.bundles) || [];
  notes.push(`bundles 共 ${bundles.length} 个`);
  for (const b of bundles) {
    const r = resolvable(b);
    if (!r.ok) {
      if (BASELINE.has(b)) { notes.push(`bundle "${b}" 不在 node_modules（基线白名单，宿主自行处理）`); continue; }
      problems.push(`bundle "${b}" 在两个解析根都找不到。` +
        (b === "dsh-ppt"
          ? " 修复：profile 目录执行 mklink /J node_modules\\dsh-ppt dsh-ppt，或跑通 pnpm install。"
          : " 修复：跑 pnpm install，或从 bundles 移除。"));
    }
  }
  for (const [dep, spec] of Object.entries(pj.dependencies || {})) {
    if (String(spec).startsWith("link:")) {
      const r = resolvable(dep);
      if (!r.ok) problems.push(`link 依赖 "${dep}" 在 profile node_modules 不存在。修复：建 junction 或 pnpm install。`);
    }
  }
} catch (e) {
  problems.push(`package.json 读取/解析失败：${e.message}`);
}

// ---- 2) cordis.patch.yml 顶层条目 ----
try {
  const lines = fs.readFileSync(path.join(PROFILE, "cordis.patch.yml"), "utf8").split(/\r?\n/);
  let pending = null;   // {id, indent}
  let checked = 0;
  for (const line of lines) {
    const idm = /^(\s*)-\s+id:\s*(\S+)\s*$/.exec(line);
    if (idm) {
      // 只看顶层 patch 条目（缩进 0 = 覆盖项，缩进 4 = insert 项）；
      // 更深层的 - id:（preset 内部插件、模型条目）不参与包解析
      const ind = idm[1].length;
      pending = (ind === 0 || ind === 4) ? { id: idm[2], indent: ind } : null;
      continue;
    }
    const nm = /^(\s*)name:\s*['"]?([^'"\n]+?)['"]?\s*$/.exec(line);
    if (nm && pending && nm[1].length === pending.indent + 2) {
      const name = nm[2];
      if (isPackageRef(name)) {
        checked++;
        const r = resolvable(name);
        if (!r.ok) {
          if (BASELINE.has(name)) notes.push(`cordis id=${pending.id} name="${name}" 基线白名单`);
          else problems.push(`cordis.patch.yml id=${pending.id} name="${name}" 无法解析。`);
        }
      }
      pending = null;
    }
  }
  notes.push(`cordis.patch.yml 检查了 ${checked} 个包引用`);
} catch (e) {
  problems.push(`cordis.patch.yml 读取失败：${e.message}`);
}

// ---- 3) dsh-ppt 自身齐备 ----
const pluginDir = path.join(PROFILE, "dsh-ppt");
if (fs.existsSync(pluginDir)) {
  for (const f of ["package.json", "cordis.patch.yml", "lib/index.js", "lib/client.js", "engine/server.py"]) {
    if (!fs.existsSync(path.join(pluginDir, f))) problems.push(`dsh-ppt 缺文件：${f}`);
  }
  try {
    JSON.parse(fs.readFileSync(path.join(pluginDir, "package.json"), "utf8"));
  } catch (e) {
    problems.push(`dsh-ppt/package.json 非法 JSON：${e.message}`);
  }
  notes.push("dsh-ppt 插件目录齐备");
} else {
  notes.push("profile 下没有 dsh-ppt 目录（未安装，跳过自身检查）");
}

// ---- 输出 ----
console.log("=== dsh-ppt 启动前一致性守卫 ===");
console.log(`profile : ${PROFILE}`);
console.log(`install : ${DSH_INSTALL}`);
for (const n of notes) console.log(`  · ${n}`);
if (problems.length === 0) {
  console.log("\n[PASS] 配置与解析根一致，可以安全重启 DSH。");
  process.exit(0);
}
console.log(`\n[FAIL] ${problems.length} 个不一致 —— 现在重启 DSH 会进恢复模式：`);
for (const p of problems) console.log("  ✗ " + p);
process.exit(1);
