# SKILL

本项目（`pptxtojson-pro`）的开发规范。先读一遍 `DESIGN.md` 了解层级与契约，再按本文件约束动手。

---

## 1. 改动前必做

1. **明确属于哪一层**。`parser` / `model` / `serializer` / `adapter` / `shapes` / `utils`，不要让一个 PR 同时跨越多层去改"同一件事"。如果发现要改两层以上，先在 commit message 里说清楚为什么。
2. **拿到一份能复现问题的 .pptx**，放进项目根目录（`*.pptx` 已 .gitignore，不会被提交）。
3. **先看源 XML**：
   ```bash
   node scripts/extract-pptx-structure.js ./xxx.pptx ./out
   ```
   把 `out/ppt/slides/slideN.xml` 拿来和 PowerPoint/WPS 的实际渲染对照，再决定 fix 的目标值。
4. **生成 JSON 对照**：
   ```bash
   npx tsx scripts/transvert.ts ./xxx.pptx ./slides.json
   ```
   修改前先跑一份"修改前 JSON"，修改后再跑一份"修改后 JSON"，diff 关键字段。

---

## 2. 代码风格

### 2.1 通用
- TypeScript `strict: true` 是底线，不要用 `// @ts-ignore` 或 `any` 绕过类型问题，确实要用 `any` 时局部最小化（如 `(c as any).path`）并附一行注释说明。
- 文件顶部用 JSDoc 注释（`/** … */`）一句话说清这个文件是干嘛的，与下面同层文件的关系是什么。
- 函数名、变量名优先英文；面向开发者的解释性注释、commit message、文档优先简体中文。
- 不写"narrating"注释（`// 给变量加 1`、`// 调用函数` 之类）。注释要解释**为什么**这样写、**有什么坑**，不要解释**做了什么**——代码本身就能说明。

### 2.2 注释要简洁有意义
**Good**：
```ts
// flipH+flipV 等价于绕中心旋转 180°；这里折算为 +180° 而非写 isFlipH/V，
// 因为部分渲染器会镜像 text 字形，与 PowerPoint/WPS 行为不符。
cRot += 180;
```
**Bad**：
```ts
// 把 cRot 增加 180
cRot += 180;
```

### 2.3 已有 helper 优先
- 单位换算：`parser/units.ts` 有 `emuToPx` / `emuToPt` / `angleToDeg` / `pctToDecimal` / `hundredthPtToPt` / `ptToPx` / `smartToPx`，**不要在业务代码里写裸的除以 914400 / 12700**。
- XML 访问：`parser/XmlParser.ts` 的 `SafeXmlNode` —— 不要直接 `.getElementsByTagName(…)[0].getAttribute(…)`，会因 PPTX 不规范而崩。
- 颜色变换：`utils/color.ts` 已经覆盖 OOXML 全套（lumMod / lumOff / satMod / tint / shade …），新代码不要重新实现。

---

## 3. 类型契约（最严格）

### 3.1 改 `src/adapter/types.ts` 要谨慎
这个文件是与下游 PPTist（及其它消费方）的**协议**。

- ❌ 不要改已存在字段的**名字**或**类型**。
- ❌ 不要把已有字段从可选改为必选。
- ✅ 新增字段一律 `?:` 可选，并附 JSDoc 说明（什么时候出现、单位、可能值）。
- ✅ 新增字段必须在 `README.md` 的"完整功能支持"里登记一行。
- ✅ 字段不要写得太通用（避免 `extra: Record<string, any>`），优先具名字段。

每次改完它都要问自己一遍：
1. PPTist 那边能消费这个新字段吗？不能就先别加。
2. 老数据没有这个字段时，下游会不会崩？最好用 `field?:` 表示。
3. 单位是不是 pt？长度类一律 pt，颜色一律 `#RRGGBB`，角度一律 deg。

### 3.2 内部类型可以更自由
`model/*` 里的类型定义（如 `ShapeNodeData`、`GroupNodeData`）只在 `src/` 内部流通，不是协议，重构相对自由——但仍然要保持"模型层不感知样式"这条原则。

---

## 4. 不同层的红线

| 层 | ✅ 该做 | ❌ 不该做 |
|---|---|---|
| `parser` | 字节解压、XML 解析、单位换算、关系映射 | 解析 OOXML 业务语义（颜色、字体、占位符…） |
| `model` | 解析"几何与结构"（位置、大小、层级、placeholder） | 解析"视觉样式"（颜色、字体、效果） |
| `serializer` | 把模型 + ctx 映射成对外 JSON 元素 | 直接读 zip / 直接操作 SafeXmlNode 之外的 DOM |
| `adapter` | 组装最终 Output、定义对外 JSON 类型 | 写业务逻辑（颜色解析、坐标转换） |
| `shapes` | 输出 SVG path 字符串 | 决定填充 / 边框 |
| `utils` | 通用、可独立测试的工具 | 引用 `model` / `serializer` 中的业务类型 |

层间依赖方向：`adapter → serializer → model → parser`，单向。`shapes` 与 `utils` 是底层工具，可被任何层引用。**禁止反向依赖**。

---

## 5. 几个高风险区，改之前看一眼

### 5.1 `serializer/groupSerializer.ts`
- `chOff/chExt` 缩放与 `bakeGroupTransform` 的 flip/rotation 烘焙。
- 改之前先把 `flipH+flipV → +180°` 这条等价规则想清楚（见 `DESIGN.md` §2.3）。
- 新增对 child 的特殊缩放规则（如 line 形状）时，**记得先做 fast-path 短路**（非 shape / 非目标 shapType 直接返回 `{ws, hs}`），避免误伤其它形状。

### 5.2 `serializer/shapeSerializer.ts`
- 800+ 行，承担了 Shape vs Text 类型判定、preset 路径生成、自适应字号、auto-fit 等多件事。改任意一段都要确认其它判定路径不受影响。
- "Shape vs Text 判定" 那段（`outputAsText` 的多分支判断）历史包袱较重，对齐了 `src1/pptxtojson.js` 行为，**调整前请用 `src1` 跑同样的 .pptx 比对**。

### 5.3 `shapes/presets.ts`
- 200+ preset，每个 preset 是一个 `(w, h, adjustments?) => string` 函数。
- 修一个 preset 时不要顺手改别的：很多 preset 之间有共享的辅助函数（如 `starShape`），改之前看一遍调用方。
- 验证修复时，最好做一份"修复前 vs 修复后"的 path diff，目视确认其它 preset 没动。

### 5.4 `parser/units.ts`
- 工具非常基础但被到处依赖，**不要改函数签名**。如果需要新单位，加新函数，不要修改老函数。

---

## 6. 脚本使用

> 完整命令清单在 `DESIGN.md` §5。这里只讲使用纪律。

- **`scripts/transvert.ts`**（推荐）：直接跑 `src/` 源码，**不需要先 `pnpm build`**。开发循环：改代码 → 跑这个 → diff JSON → 继续改。
  ```bash
  npx tsx scripts/transvert.ts ./xxx.pptx ./slides.json
  ```

- **`scripts/transvert.js`**：跑 `src1/` 原版，作回归基准。当新版输出形态可疑时，跑同一份 `.pptx` 与之对照。

- **`scripts/extract-pptx-structure.js`**：解压 .pptx，看源 XML。**先看 XML、再改代码**永远比"凭感觉改"快得多。
  ```bash
  node scripts/extract-pptx-structure.js ./xxx.pptx ./out
  ```

- **不要把脚本输出物纳入 git**：
  - `*.pptx`、`slides.json`、`out/`、`dist/`、`docs/` 都已在 `.gitignore` 中，不要 `git add -A` 然后把它们带进 commit。提交前用 `git status` 看一眼。

---

## 7. 调试与定位 bug 的工作流

定位"输出和 WPS/PowerPoint 不一致"类问题的标准流程：

1. **复现**：用 `演示文稿1.pptx` 之类的最小 case，跑 `transvert.ts` 拿到 JSON。
2. **看源 XML**：`extract-pptx-structure.js` 释放 zip，找到对应元素的 `<p:sp>` / `<p:grpSp>` / `<p:graphicFrame>`，确认 OOXML 原始数据。
3. **写一段 throwaway 脚本**做对比（直接 inline `python3 << 'EOF' …`）：用 `slides.json` 把对应元素的关键字段（left/top/width/height/path/rotate/flips）打出来，与 WPS 视觉效果对齐。
4. **沿着管线倒推**定位是哪一层出错：
   - JSON 数值离谱 → 看 `serializer` 或 `adapter`。
   - 节点类型识别错误 → 看 `model/Slide.ts` 的分派逻辑或 `model/nodes/*` 的 parser。
   - 颜色 / 填充错 → 看 `serializer/StyleResolver.ts` + `utils/color.ts`。
   - 模板继承错 → 看 `RenderContext.ts` 的链路。
5. **修最小切面**：能在 serializer 修就别改 model，能加 fast-path 就别改主路径。
6. **回归验证**：跑全套 `*.pptx` 看 diff（至少跑 `演示文稿1.pptx`）。统计性指标比"看起来对了"更可靠（例如修 line 时，统计全文档横线/竖线/对角线数量是否符合预期）。

---

## 8. Commit & PR

- Commit message 简短、单语言（中文优先）、动词起头、点出**修复了什么/为什么**：
  - ✅ `fix(line): 避免 group 内水平/垂直线 bbox 被 chOff/chExt 极端缩放放大`
  - ✅ `refactor(group): 简化 sizeScaleForChild 类型断言`
  - ❌ `fix bug` / `update` / `修复一些问题`
- 一个 commit 只做一件事。重构和 bug 修复分开提。
- 改了 `adapter/types.ts` 必须在 commit body 里写明**协议变更说明**。
- 提交前 `git status` 自检：不要带进 `*.pptx` / `slides.json` / `out/` / `dist/` / 临时调试改出的 `index.html` 等无关文件。

---

## 9. 何时该求助 `src1/`

`src1/pptxtojson.js` 是仓库原版 JS 实现，是新版的"行为基准"：

- 新版输出的字段值与你预期不符，但不确定是 bug 还是历史行为？→ 先跑 `transvert.js` 看 `src1` 的输出，对照决定到底改新版还是接受现状。
- 新增字段时不知道命名/形状？→ 先看 `src1` 是否已经有类似输出，对齐其命名习惯。
- 重构某段逻辑前？→ 看 `src1` 对应实现，确认 corner case 没漏。

`src1/` **只读**，不要改它，也不要让它参与构建。

---

## 10. 不要做的事 ❌

- 不要在 `model/` 里读 theme/master 颜色（那是 serializer 的事）。
- 不要在 `serializer/` 里直接读 zip 文件或 .rels。
- 不要在 `parser/` 里写 OOXML 业务逻辑（颜色解析、占位符回退等）。
- 不要修改 `parser/units.ts` 现有函数的签名。
- 不要给 `adapter/types.ts` 加非可选的新字段，除非协议方明确同意。
- 不要在不读 `src1/` 历史行为的情况下改 `shapeSerializer.ts` 的"Shape vs Text 判定"分支。
- 不要在 commit 中带入 `*.pptx`、`slides.json`、`out/`、`dist/`、临时调试用的 `index.html` 改动。
- 不要写 `// @ts-ignore`；不要随意 `as any`；确实需要时局部最小化并加注释。
- 不要在 PR 中既"重构"又"修 bug"。
