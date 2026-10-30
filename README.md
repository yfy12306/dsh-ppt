# dsh-ppt

DSH 桌面端的 PPT 制作插件：**用本机真实 Microsoft PowerPoint / WPS 演示生成演示文稿**，
并且**全程实时可见**——PowerPoint 窗口里看着页面一页页长出来，DSH 输入框旁的预览面板
同步显示当前页缩略图和「AI 正在做什么」的操作字幕。

精华抽取自 [yb2460/harness-anything](https://github.com/yb2460/harness-anything)（MIT），
针对 Microsoft Office 实测重写，并补齐上游没有的实时预览能力。

---

## 一、它解决什么问题

AI 做 PPT 的三条老路都走不通：

| 老路 | 问题 |
|---|---|
| python-pptx 直接写文件 | 用户看不到过程，改一处要重开文件 |
| 截图 + 鼠标点 GUI（RPA） | 脆弱、慢、坐标一偏全错 |
| 生成 HTML 再转 | 丢格式、不可编辑 |

dsh-ppt 走第四条：**COM 自动化真实 Office**。生成的 `.pptx` 是原生可编辑的，
表格是真表格、图表是真图片、文本框是真文本框；同时因为走的是活着的
PowerPoint 进程，用户能实时观看。

## 二、实时可见的三层实现

1. **真实 PowerPoint 窗口**（最强）
   `Presentations.Add(WithWindow=msoTrue)` 建文稿，每加一页/一个元素后
   `ActiveWindow.View.GotoSlide(idx)`。你在 PowerPoint 窗口里看着它画。
2. **DSH 内预览面板**
   引擎进程内起一个**只读** HTTP 侧车（默认 `127.0.0.1:8790`），每次变更后把
   当前页 `Slide.Export` 成 PNG 落盘。Node 宿主把它代理成同源路由 `/ppt/*`，
   客户端芯片弹层每 1.2s 轮询，缩略图跟着变。
3. **操作流水字幕**
   每个动作记一条 `{seq, 时间, 类型, 页码, 摘要}`，面板按时间倒序展示。

线程安全：COM 对象只能在 `CoInitialize` 过的线程用。HTTP 侧车跑在独立 daemon
线程，**只读磁盘 PNG 和 state.json，绝不触碰 COM**，因此不会跨线程崩溃。

## 三、从 harness-anything 抽了什么、改了什么

### 抽取（保留价值）

| 上游资产 | 落到本插件 |
|---|---|
| `pptengine/wps.py` 的 Deck + 7 个 COM 坑规避 | `pptx_engine/office.py` + `deck.py` + `elements.py` |
| `pptengine/elements.py` 绘图原语与组合组件 | `elements.py`（36 种元素类型） |
| `pptengine/assets.py` PIL 程序化素材 | `assets.py`（+ 主题色感知 + 中文字体探测） |
| `pptengine/verify.py` PyMuPDF 溢出校验 | `verify.py`（+ 重叠检测 + 空页检测 + parity） |
| `styles/design_presets.py` 4 套预设 | `themes.py`（收敛为 default 风格 + 机器可读布局硬约束 + 排版/动画/配色规范） |
| `styles/quality_checks.py` 5 维度清单 | `quality.py`（把 SKILL.md 的口传规则变成可执行检查） |
| `cli_anything/zotero` 的 sqlite/paths/http | `zotero.py`（收窄为检索+引用+参考文献页） |
| SKILL.md「JSON 数据驱动 6 步流程」 | `ppt_build` 一个工具搞定 |

### 修正（上游的坑，本机实测）

1. **不无条件 `taskkill`**。上游每次生成前 `taskkill /F /IM wps.exe`，会毁掉用户
   未保存的文档。本插件只关自己建的文稿；新建前若有未保存旧文稿，先自动存到
   `<workdir>/recovery/` 再关。
2. **PowerPoint 的 `Application.Visible` 不能设 False**（抛 `com_error`）。
   无窗口模式必须走 `Presentations.Add(WithWindow=msoFalse)`。上游写法在 WPS 成立、
   在 MS Office 直接报错。
3. **ProgID 自动探测**：`PowerPoint.Application` → `KWPP.Application`，同一套元素代码通吃。
4. **中文姓名引用**：上游没有引用渲染；本插件本地实现 GB/T 7714 / APA / IEEE，
   且中日韩姓名连写全名（「于飞洋」而非「于, 飞.」），西文姓名才走首字母缩写。
5. **饼图小扇区标签重叠**：任一扇区 <6% 时自动改用图例承载名称，百分比只标大扇区。
6. **MCP 协议零依赖**：不引 `mcp` 包（本机唯一带它的解释器是 hermes-agent 的 venv，
   随时可能重建）。协议层手写于 `mcp_stdio.py`，且启动时把 `sys.stdout` 重定向到
   stderr，任何库的意外 `print()` 都不会撕坏协议流。
7. **上游 Illustrator harness 的 setup.py 指向不存在的包**（`cli_anything.powerpoint`）
   —— 直接弃用，不搬。

## 三·补、设计系统（调研驱动，2026-10-28 重做）

开源版内置 default 一套风格（内部品牌主题不开源），按用户要求「不要过度设计、干净、细节、有省有略、丰富」：

| 风格 | 取向 | 关键元素 |
|---|---|---|
| `default` | 干净商务 | 近白底 #F7F9FC(60%) + 藏青 #1F3A5F(30%) + 橙 #F28C28(10%)；雅黑标题+等线正文；页脚/细线/统一边距 |

每条规则都有出处（详见 CHANGELOG 1.1.0）：

- **字号**：封面 44-48 / 页标题 30-32 / 正文 18-20（最低 16）/ 注释 10-12 / 页脚 10-12；行距 1.35x
  （k-dense、slidesmate、清华讲义）
- **字体**：全篇 ≤2 正文字体族 + 1 数字族 + 1 等宽族；角色化（display/title/body/caption/numeric/mono/accent），
  按本机字体实测解析候选列表，避免 PowerPoint 静默回退（slidesmate 两字体系统）
- **配色**：60-30-10；强调色只给大数字/单条图表系列/高亮箭头；图表分类色默认 Okabe-Ito 色盲安全板；
  WCAG 对比度正文 ≥4.5:1、大字 ≥3:1，`ppt_quality` 实测拦截（slidesbrain、Nature）
- **动画**：整页 Fade 0.6s（无切换像「仓促拼凑的 PDF」）；入场 Fade 错落 0.3s；wipe 只给线/箭头；
  可 `entrance="none"` 全关（verdanabold、animation-principles、Apple 0.8s cross-dissolve）
- **版式**：一页一观点；正文左对齐；留白 40-50%；数据页图表占 70-80%；标题写结论（action title）；
  章节分隔页 divider（大编号+≤5词）；自动页脚（deck 名+页码）（deckary 咨询版式）
- **背景**：alpha≈12 的淡网格纹理，规避「纯白底黑字 = instant boredom」又不抢内容（k-dense 反模式警告）

## 四、目录结构

```
dsh-ppt/
├── package.json              # DSH 插件清单（host 半 + client 半）
├── cordis.patch.yml          # bundle patch：注入插件 + 默认配置
├── lib/
│   ├── index.js              # host 半：挂 MCP 面 / 代理 /ppt/* / /ppt 命令 / Skill
│   └── client.js             # client 半：实时预览芯片 + 弹层面板
├── engine/
│   ├── server.py             # MCP 工具定义 + 全局会话（25 个工具）
│   ├── requirements.txt
│   └── pptx_engine/
│       ├── office.py         # PowerPoint COM 连接（MS Office 优先，WPS 兜底）
│       ├── deck.py           # Deck 会话 + 实时可见三件套
│       ├── elements.py       # 36 种绘图原语/组合组件（含 7 个 COM 坑规避）
│       ├── renderer.py       # 声明式 JSON spec → 绘图调用的路由表
│       ├── assets.py         # PIL 素材（渐变横幅/竖条/圆环/箭头/背景/封面）
│       ├── charts.py         # matplotlib 图表（中文字体已配好）
│       ├── verify.py         # PyMuPDF 落盘校验 + 逐页渲染
│       ├── themes.py         # default 风格 + 排版系统 + 布局硬约束
│       ├── quality.py        # pre-flight 静态审查 + 5 维度终评
│       ├── zotero.py         # Zotero 只读检索 + 引用渲染
│       ├── live.py           # 实时预览 HTTP 侧车（只读）
│       └── mcp_stdio.py      # 零依赖 MCP stdio 服务端
└── scripts/
    ├── smoke.py              # 引擎独立冒烟（14 项，不经过 DSH）
    └── validate.mjs          # MCP 端到端（12 项，用官方 SDK 当客户端）
```

## 五、安装与启用

插件本体已就绪并通过全部测试。启用需要把它登记进 profile（**需重启 DSH 生效**）：

1. `profiles/desktop/package.json` 的 `dependencies` 加：
   `"dsh-ppt": "link:C:/Users/user/.dsh/profiles/desktop/dsh-ppt"`
2. `dependencies["dsh"].profile.bundles` 加 `"dsh-ppt"`
3. `profiles/desktop/cordis.patch.yml` 追加插件自带的 `cordis.patch.yml` 内容
4. 在 profile 目录 `pnpm install`，重启 DSH

启用后：
- 模型侧出现 `mcp__ppt__*` / `mcp__zotero__*` 工具与 `dsh-ppt` Skill
- 输入框右侧出现 PPT 芯片，点开即实时预览面板
- 斜杠命令 `/ppt` 做状态自检

## 六、使用示例（模型侧）

```
ppt_status                                  # 先看环境
ppt_new  theme=default                          # 建文稿（干净商务风格）
ppt_build slides=[                          # JSON 一次建整份
  {elements:[{type:cover, title:产线质量月报, ...}]},
  {elements:[{type:header, sec:1.1, title:合格率},
             {type:stat_cards, box:[16,84,928,96], items:[...]},
             {type:chart, chart:bar, box:[24,196,440,260], data:{...}},
             {type:table, box:[488,196,448,200], header:[...], rows:[...]}]}
]
ppt_quality                                 # pre-flight 审查
ppt_save                                    # 出 pptx + pdf + 落盘校验
ppt_verify                                  # PDF 实测溢出/重叠/parity
```

文献（Zotero 装了才可用）：

```
zotero_search query=外倾角
zotero_cite   refs=[...] style=gbt7714
zotero_refs_slide refs=[...] insert=true    # 直接生成参考文献页
```

## 七、测试

```powershell
cd C:\Users\user\.dsh\profiles\desktop\dsh-ppt
python scripts\smoke.py                 # 可见模式，能看着 PPT 长出来（14 项）
python scripts\smoke.py --headless      # 无窗口（CI 友好）
node   scripts\validate.mjs --headless  # MCP 端到端（12 项）
node   scripts\test-host-load.mjs       # 宿主半隔离加载（9 项，不碰运行中的 DSH）
node   scripts\preflight.mjs            # 启动前一致性守卫（重启 DSH 前必跑）
```

当前结果：smoke 14/14、validate 12/12、host-load 9/9、preflight PASS。
PDF 实测：溢出 0 / 重叠 0 / 空页 0 / parity 100。

## 七·补、安全与防崩（2026-10-28 事故后新增）

本插件曾把 DSH 拖进恢复模式一次，根因是**登记顺序错误**：配置先引用了包、
而 `pnpm install` 因无关的 dsh-purge 校验和漂移中止，包没装上。现已四重设防：

1. **junction 代替 pnpm**：`node_modules\dsh-ppt` 是手动目录链接，包解析不依赖
   那条已损坏的 pnpm 链路。
2. **preflight 守卫**：`node scripts/preflight.mjs` 按 DSH 真实解析根校验所有包引用，
   不绿就不重启。已做反向验证（撤链接→FAIL）。
3. **host 防御外壳**：`apply()` 整体 try/catch，插件异常只降级为「不挂载」，
   不冒到 profile-composition。
4. **client 防御外壳**：client 半异常只降级为「不显示芯片」，不冒到渲染进程。

回滚（任何时候，两步）：

```powershell
# 1) 摘掉配置引用
#    package.json 删 dsh-ppt 两行；cordis.patch.yml 删 dsh-ppt 那段
# 2) 摘掉链接
rmdir "C:\Users\user\.dsh\profiles\desktop\node_modules\dsh-ppt"
# 重启 DSH 即回到无插件状态。插件目录本身可留可删。
```
（事故当次的备份：`package.json.bak-dshppt-20261028-154018`、
`cordis.patch.yml.bak-dshppt-20261028-154018`）

## 八、配置项（cordis.patch.yml）

| 键 | 默认 | 说明 |
|---|---|---|
| `visible` | true | 弹出真实 PowerPoint 窗口实时观看 |
| `backend` | auto | auto / ms-office / wps |
| `theme` | default | 风格：default（干净商务）；支持自定义主题 |
| `livePort` | 8790 | 实时预览侧车端口（占用自动 +1 顺延） |
| `workdir` | ~/Desktop/dsh-ppt-out | 产物/素材/恢复副本目录 |
| `autoPdf` | true | 保存时同导 PDF |
| `snapshotScale` | 960 | 预览快照导出宽度 px |
| `quitOnExit` | false | 退出时不杀 PowerPoint，保留窗口给用户 |

## 九、已知边界

- **Windows only**（COM 依赖）。
- Zotero 未安装时 `zotero_*` 返回 `available=false`，不影响 PPT 主流程。
- 桌面锁屏时 PowerPoint 窗口不可见，但引擎与预览侧车照常工作（解锁后即见）。
- 实时预览面板依赖 DSH 回环路由；若 `/ppt/*` 404，先跑 `/ppt` 看 `live.up`。

## 许可

MIT。上游 harness-anything 亦为 MIT，抽取部分保留其设计意图并在代码注释中标明出处。
