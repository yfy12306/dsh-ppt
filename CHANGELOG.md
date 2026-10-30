# Changelog

## 1.1.1 — 2026-10-29（品牌主题图表全面回归 品牌色系）

用户反馈：品牌主题里的图表线体/序列色还在用 Okabe 的蓝橙粉，不像 品牌 家族。

- **品牌主题图表序列色**改为 品牌色系（`typo.chartPalette`）：
  `#36A36E` primary-500 主序列 → `#6B7568` 中性灰绿对照序列 → `#094828` primary-700 第三序列
- **新增 `theme.semantic`**（good/warn/bad/info/deep/muted，取自 品牌语义色），
  饼图语义切片示范：合格=good 绿、返修=warn 橙、报废=bad 红。
- **折线加 SPC 签名**（designmd-demo v2 定案的视觉层）：Catmull-Rom 平滑 +
  同色加宽半透明光晕 + 淡面积 + 白边描点。**面积只给主序列** —— 控制限/对照线填面积
  会糊成一大片灰带（实测翻车点），需要全填用 `areaAll=true`。
- 环心大数字、总结条底色改为跟主题（品牌主题=深绿/浅绿底，default=藏青/浅蓝底），去掉硬编码。
- 修复 `_catmull_rom` 解包 bug（把两个切片解成四个点名 → ValueError）。

验证：smoke 14/14（品牌主题与 default 各一遍，7 页，溢出 0/重叠 0/parity 100）、
validate 12/12、host-load 9/9、preflight PASS。

## 1.1.0 — 2026-10-28（设计系统重做：两风格 + 排版/动画/配色全部按调研落地）

用户质问「有动画吗、有审美吗、有字体字样设计吗」，如实回答：**动画=无、字体=单一微软雅黑、
审美=工程整洁**。随后按用户要求先做网络调研、再重做，并把 7 套主题收敛为 2 套。

### 调研依据（不再凭感觉配）

- 字号/字体/留白/6×6：[k-dense scientific-slides 设计原则](https://github.com/rubensliv/k-dense-ai/blob/main/scientific-skills/scientific-slides/references/slide_design_principles.md)、
  [slidesmate 排版指南](https://slidesmate.com/blog/presentation-fonts-and-typography)、
  [k-dense PowerPoint 设计指南](https://github.com/rubensliv/k-dense-ai/blob/main/scientific-skills/scientific-slides/assets/powerpoint_design_guide.md)、清华《PPT的设计与制作》
- 两字体系统 / 字号层级 / 行距 1.2-1.5x / 单行 50-65 字符：slidesmate
- 动画：[verdanabold（Fade 是主力、wipe 画线、rise up 收尾）](https://verdanabold.com/post/bring-your-next-presentation-to-life-with-powerpoint-animation)、
  [animation-principles/presentations（builds 0.3-0.5s、transitions 0.5-0.8s、错落 0.1-0.2s、禁线性）](https://github.com/dylantarre/animation-principles/blob/main/docs/01-by-domain/presentations.md)、
  Apple Keynote 0.8s cross-dissolve
- 配色：[60-30-10 与 WCAG 对比度实测](https://slidesbrain.com/the-60-30-10-color-rule-for-slides-how-to-choose-presentation-palettes)、
  Okabe-Ito 色盲安全板（[Nature 图规范](https://research-figure-guide.nature.com/figures/building-and-exporting-figure-panels)、k-dense color_palettes）
- 版式：[咨询五版式 / action title / chart-dominant / 章节分隔页](https://deckary.com/blog/powerpoint-layout-ideas)
- 中文场景字体惯例（学术宋体+Times、不确定会场用雅黑）：chengqingli 学术 PPT 教程

### 新增 / 重做

- **两风格**：`default`（近白+藏青30%+橙10%，干净商务）与 `brand`（品牌单色体系：
  六边形开口标识 + 品牌字标 + 深绿封面/分隔页 + 绿色细线页脚）。删除其余 5 套（过度设计）。
- **排版系统落地**：themes 的 fonts/scale 真正接进渲染器（此前 fontSizes 是死数据）。
  角色：display/title/body/caption/numeric/mono/accent；字号按调研（display 44-48、h1 30-32、
  body 18-20、caption 10-12、页脚 10-12）；字体按本机实测解析候选列表。
- **动画**：整页切换默认 Fade 0.6s（`SlideShowTransition`，读回校验防错常量）；
  元素入场 Fade 错落 0.3s（`TimeLine.MainSequence`）；可 `entrance="none"` 关闭。
- **版式细节**：自动页脚（deck 名+页码）、章节分隔页 divider、品牌行、主题纹理背景
  （alpha≈12 的淡网格，规避「纯白底黑字=instant boredom」又不抢内容）。
- **质量审查加 WCAG 对比度实测**（正文 4.5:1 / 大字 3:1）与「强调色不得用于低对比正文」告警。
- **图表分类色默认 Okabe-Ito**（品牌绿打头 for brand）；修复「主题单色 palette 画多分类饼图同 hue」翻车。

### 修复

- 标题 32pt 居中锚定导致字形 bbox 顶到 y=-2 被判画布溢出 → 改顶对齐 + 44pt 框高，溢出归零。
- 实时预览侧车偶发 404：①快照改原子写（tmp+os.replace）；②`ThreadingHTTPServer` 在 Windows
  上 `allow_reuse_address=1` 会与遗留进程双监听同端口 → 新增 `_NoReuseServer` 禁用之，端口干净顺延。

### 验证

smoke 14/14（default 与品牌主题 各一遍，溢出 0/重叠 0/parity 100）、validate 12/12、
host-load 9/9、preflight PASS。

## 1.0.1 — 2026-10-28（事故修复 + 防复发）

### 事故回顾（本机真实发生）

登记插件后首次重启，DSH 进入恢复模式。诊断包
（`diagnostics-1793173682876-…zip`）定位到唯一硬失败：

```
PackageOverlayNotFoundError: cannot resolve package "dsh-ppt"
  at resolveOverlayPackage (package-overlay-*.js:105)
  at loadRecoveryFilteredProfile (profile-*.js:697)   ← profile-composition 阶段
```

因果链（**顺序错误，非引擎 bug**）：

1. 我把 `dsh-ppt` 写进 profile 的 `dependencies` + `bundles` + `cordis.patch.yml`；
2. 随即在后台跑 `pnpm install`，但它因 **dsh-purge 的 github tarball 校验和与
   lockfile 漂移**（`ERR_PNPM_TARBALL_INTEGRITY`，作者推了新提交，与 dsh-ppt 无关）
   而整体中止 → `node_modules/dsh-ppt` 从未生成；
3. 用户在此窗口期重启 DSH → profile-composition 解析不到 `dsh-ppt` → 硬失败 → 恢复模式。

注：诊断包里的 renderer watchdog 崩溃在 10-22/23/26 亦有记录，属历史既有抖动，
与本次无关（当次已自动恢复 healthy）。

### 修复

- **手动 junction 代替 pnpm**：`mklink /J node_modules\dsh-ppt dsh-ppt`，
  使包解析不再依赖那条已损坏的 pnpm 链路。
- **`scripts/preflight.mjs` 启动前一致性守卫**：按 DSH 的真实解析根
  （profile/node_modules + 安装目录/resources/app/node_modules）校验 bundles 与
  cordis.patch.yml 的每个包引用；对「新出现的不可解析项」报 FAIL（退出码 1），
  宿主内建/`disabled` 条目进基线白名单。已做反向验证：撤掉 junction → FAIL，恢复 → PASS。
  **规矩：preflight 不绿，不重启 DSH。**
- **`lib/index.js` 防御外壳**：`apply()` 整体 try/catch，插件自身异常只记日志并降级为
  「不挂载」，绝不冒到 profile-composition。
- **`lib/client.js` 防御外壳**：client 半 `apply()` try/catch，异常只降级为「不显示芯片」，
  不冒到渲染进程（避免触发 renderer watchdog）。
- **`scripts/test-host-load.mjs`**：临时目录 + 桩 `@deepseek-ai/*`，隔离跑 `apply()`，
  9 项断言（含「ctx.plugin 抛异常时 apply 仍不冒异常」），不碰运行中的 DSH。

### 已知注意

- 若日后跑通 `pnpm install`（先处理 dsh-purge 校验和：`pnpm install --update-checksums`，
  属 supply-chain 决策，自行判断），pnpm 可能对手动 junction 报错；
  先 `rmdir node_modules\dsh-ppt` 再 install，link: 依赖会被正常重建。

## 1.0.0 — 2026-10-28

首个版本。从 [yb2460/harness-anything](https://github.com/yb2460/harness-anything)（MIT）
抽取 PPT 与 Zotero 精华，针对 Microsoft Office 实测重写，做成 DSH 插件。

### 新增

- **引擎**（`engine/pptx_engine/`，Python，零 `mcp` 包依赖）
  - `office.py`：PowerPoint COM 连接层，MS Office 优先 / WPS 兜底自动探测；
    修正上游「PowerPoint 不能设 Visible=False」与「无条件 taskkill」两个坑
  - `deck.py`：Deck 会话 + 实时可见三件套（GotoSlide / 当前页 PNG 快照 / 操作流水）
  - `elements.py`：36 种绘图原语与组合组件，内置上游 7 个 COM 坑规避
  - `renderer.py`：声明式 JSON spec → 绘图调用的类型路由表（`@route` 注册）
  - `assets.py`：PIL 程序化素材 + 主题色感知 + Windows 中文字体探测
  - `charts.py`：matplotlib 8 类图表，中文字体已配好；小扇区饼图自动改图例防标签重叠
  - `verify.py`：PyMuPDF 落盘校验（溢出/重叠/空页）+ PPTX/PDF parity + 逐页渲染
  - `themes.py`：7 套主题（早期 7 套，后收敛）
    + 机器可读布局硬约束（y+h≤518、标题间距≥24pt 等）
  - `quality.py`：pre-flight 静态审查（9 条规则）+ 上游 5 维度终评骨架
  - `zotero.py`：Zotero 只读 sqlite 检索 + GB/T 7714 / APA / IEEE 本地引用渲染
    （中文姓名连写全名）+ 参考文献页生成；未安装时优雅降级
  - `live.py`：实时预览 HTTP 侧车（只读线程，绝不触碰 COM）
  - `mcp_stdio.py`：手写 MCP stdio 协议层；stdout 重定向防协议污染
- **MCP 工具面**（25 个）：`ppt_status/new/open/slide/element(s)/build/chart/table/
  image/assets/goto/remove_slide/snapshot/quality/save/verify/close/live/themes`
  + `zotero_status/search/cite/refs_slide/collections`
- **DSH 插件壳**
  - `lib/index.js`：stdio 挂 MCP 面、代理 `/ppt/*` 同源路由（仅回环）、`/ppt` 命令、
    运行时 Skill `dsh-ppt`
  - `lib/client.js`：输入框右侧实时预览芯片 + 弹层（缩略图 + 操作字幕，1.2s 轮询）
- **测试**：`scripts/smoke.py`（14 项引擎冒烟）、`scripts/validate.mjs`
  （12 项 MCP 端到端，用官方 `@modelcontextprotocol/sdk` 当客户端）

### 相对上游的行为变更（有意为之）

- 不再 `taskkill` Office 进程；新建前未保存旧文稿自动存 `<workdir>/recovery/`
- 默认不退出 PowerPoint（`quitOnExit=false`），保留窗口给用户查看/续编
- 弃用上游 Illustrator/Photoshop harness（setup.py 指向不存在的包，无法安装）

### 验证

- smoke 14/14、validate 12/12（Microsoft Office 16.0，Python 3.12.8）
- PDF 实测：溢出 0 / 重叠 0 / 空页 0 / parity 100
- 目检：封面页与数据页中文渲染正常、无方块、无标签重叠
