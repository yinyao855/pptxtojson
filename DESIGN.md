# DESIGN

本文档介绍 `pptxtojson-pro` 的实现思路与目录结构，方便后续开发与定位修复点。

阅读顺序建议：先读「整体管线」了解一份 `.pptx` 是如何变成 JSON 的，再按需翻阅「目录结构」对应章节查看每一层的职责与关键模块。

---

## 1. 整体管线

`parse(buffer)` 入口位于 `src/index.ts`，把一份 `.pptx`（即 zip）转成最终的 JSON。整体是 **5 层流水线**，每一层只负责一种抽象，层之间单向依赖：

```
ArrayBuffer
   │
   │  ① ZIP 解包
   ▼
PptxFiles                       (parser/ZipParser)
   │
   │  ② XML → 类型化模型
   ▼
PresentationData                (model/*)
   │
   │  ③ 准备渲染上下文
   ▼
RenderContext (per slide)       (serializer/RenderContext)
   │
   │  ④ 节点 → JSON 元素
   ▼
Element[] (Shape | Text | …)    (serializer/*)
   │
   │  ⑤ 组装最终输出
   ▼
Output { slides, themeColors, size }   (adapter/toPptxtojson)
```

### ① ZIP 解包：`parser/ZipParser.ts`
把 `.pptx`（OOXML 包）读入并按用途分类：`slides`、`slideLayouts`、`slideMasters`、`themes`、`media`、`charts`、`diagramDrawings`、`notesSlides` 等。同时执行 zip-bomb 防御（限制条目数 / 单文件解压大小 / 总解压大小 / 媒体大小 / 并发数）。

### ② XML → 模型：`model/`
- `Theme.ts` / `Master.ts` / `Layout.ts` 解析三层模板，建立 **theme → master → layout → slide** 的继承链与 placeholder 表。
- `Slide.ts` 把每页 `spTree` 下的 `p:sp` / `p:pic` / `p:grpSp` / `p:graphicFrame`（含 table / chart / OLE / SmartArt fallback）/ `p:cxnSp` / `mc:AlternateContent`（含公式）逐一识别后，分派给 `model/nodes/*` 的对应解析器。
- `nodes/BaseNode.ts` 抽出所有节点共有的属性（`position`、`size`、`rotation`、`flipH`、`flipV`、`placeholder`、`hlinkClick`、`source`、`xmlOrder`），其它 node 文件通过组合 `parseBaseProps` 复用。
- `Presentation.ts` 把上面所有零散结果组装成 `PresentationData`，并保留 slide → layout → master → theme 的索引。

> **重要约定**：模型层只解析 **几何与结构**，不解析视觉样式（填充、边框、文字格式）。视觉样式留到 serializer 层，因为它们才需要 theme/master 的级联解析。

### ③ 渲染上下文：`serializer/RenderContext.ts`
每张幻灯片在序列化前都会基于 `PresentationData` 创建一个 `RenderContext`，缓存：
- 当前 slide / layout / master / theme
- color map override 链
- 媒体输出模式（`base64` 或 `blob`）
- 字体、列表样式、占位符回退等所有"渲染时"需要查的东西

之后所有 `serializer/*ToElement(node, ctx, ...)` 都通过 `ctx` 拿样式来源，避免到处传参。

### ④ 节点 → JSON 元素：`serializer/`
每种节点都有专属 serializer：

| node                | serializer               | 输出类型                     |
| ------------------- | ------------------------ | ---------------------------- |
| `ShapeNodeData`     | `shapeSerializer.ts`     | `Shape` 或 `Text`            |
| `PicNodeData`       | `imageSerializer.ts`     | `Image` / `Video` / `Audio`  |
| `TableNodeData`     | `tableSerializer.ts`     | `Table`                      |
| `ChartNodeData`     | `chartSerializer.ts`     | `Chart`（柱/线/饼/散点 …）   |
| `GroupNodeData`     | `groupSerializer.ts`     | `Group`（递归处理子元素）   |
| `MathNodeData`      | `mathSerializer.ts`      | `Math`                       |
| TextBody（公共）    | `textSerializer.ts`      | HTML 富文本字符串            |

支撑模块：
- `slideSerializer.ts`：编排顺序 = **背景 → master 装饰 → layout 装饰 → slide 元素**；按需 fallback 到 layout/master 的占位符位置/尺寸。
- `StyleResolver.ts`：把 OOXML 颜色 / 填充节点解成 CSS 字符串。
- `borderMapper.ts`：把线型转成 `borderType` + `borderStrokeDasharray`。
- `backgroundSerializer.ts`：背景填充解析（颜色 / 渐变 / 图片 / 图案）。

### ⑤ 组装输出：`adapter/`
- `adapter/types.ts`：**对外 JSON 的 TypeScript 类型定义**，与下游 PPTist 一致。修改它会直接影响 JSON 形状。
- `adapter/toPptxtojson.ts`：把 `PresentationData` 中所有 slide 经 `slideToSlide` 转换后，再加上 `themeColors`、`size`，得到最终 `Output`。

---

## 2. 关键设计点

### 2.1 模型层不感知样式
`model/*` 只解析"它是什么、在哪里、多大"，不解析"它是什么颜色"。这条线如果模糊会导致 theme/master 的级联解析散落到各处。

### 2.2 Serializer 是"无副作用映射"
所有 `*ToElement(node, ctx, order, files?)` 都是纯函数式映射：输入 `model + ctx`，输出一个 `Element`。这样：
- 顺序由 `slideSerializer` 统一控制；
- 子元素递归（group）天然干净；
- 单元定位 bug 时只需怀疑一个文件。

### 2.3 Group 坐标空间烘焙
`groupSerializer.ts` 处理两个棘手问题，每个都集中在该文件内：

1. **`chOff/chExt` 坐标系缩放**：Group 内的子元素位置/尺寸是相对于 group 自定义坐标空间的（`a:chOff/a:chExt`），常见于 chart / 老 MS Graph 形状（chOff/chExt 不是 EMU 而是私有单位）。出口处统一用 `ws = ext.cx / chExt.cx`、`hs = ext.cy / chExt.cy` 把子元素映射到外层坐标。
2. **flip / rotation 烘焙**：通过 `bakeGroupTransform` 把 group 自身的 `flipH/flipV/rotate` 折算进每个子元素的 `left/top/rotate/isFlipH/isFlipV`，让最终输出的 group 始终是中性的（`rotate: 0`、`isFlipH/V: false`），下游渲染器无需再组合 group 级变换。
   - 特别注意：`flipH+flipV` 等价于绕中心旋转 180°，烘焙时直接转换为 `+180°`，不写 `isFlipH/V` —— 因为 PowerPoint/WPS 的 flip 不会镜像 text 字形，而部分渲染器会，按 `+180°` 写入是兼容性最好的做法。

### 2.4 Line 形状的特殊缩放
`shapeSerializer` 里 line-like 形状 (`prst="line"`、`straightConnector1`、connector 类) 的 sub-pixel 维度需要 bump 到 1px 才能给 SVG 一个非退化 viewBox。但若再被 group `ws/hs` 放大，stroke 厚度会变成几百 pt。

为此：
- `shapeSerializer` 仅 bump **垂直于线段方向**的轴（横线只 bump 高度、竖线只 bump 宽度），并通过 `pathH=0` / `pathW=0` 暗示 preset 走水平/垂直分支。
- `groupSerializer` 中的 `sizeScaleForChild` 对 line 形状同样跳过 stroke 厚度轴的缩放，path 也对应只缩放方向轴。

这样无论 group 的 `chOff/chExt` 多极端，横线/竖线的 bbox 始终 = `(线长, stroke 厚度)`，下游按 bbox 对角线绘制 line 的渲染器也不会画出 V 字。

### 2.5 单位约定
- `parser/units.ts` 定义所有单位换算。OOXML 的 EMU 在 `model` 层就转成 px，最终在 `adapter` / `groupSerializer` 中用 `pxToPt = px * 0.75` 转成 pt。
- **对外 JSON 一律 pt**。新增字段如带长度，必须先转 pt 再写入。

### 2.6 SafeXmlNode：null-safe XML 访问
`parser/XmlParser.ts` 提供 `SafeXmlNode`，所有 `.child(name)` / `.attr(name)` / `.numAttr(name)` 调用即使节点不存在也不会抛出 —— 这是应对 PPTX 文件结构高度不规范的关键。新代码应一律使用 `SafeXmlNode`，避免裸的 DOM 访问。

### 2.7 参考实现 `src1/` 与 `pptx-renderer-main/`
- `src1/` 是仓库原版 JavaScript 实现 (`pptxtojson.js`)。**不参与构建**，仅作为输出格式 / 行为对齐参考。当对某个字段输出形态有疑问时，先去 `src1` 看历史行为，确保新版兼容。
- `pptx-renderer-main/` 是用来真实渲染 OOXML 的参考 DOM 渲染器（被 .gitignore 忽略，本地参考用）。`shapeSerializer.ts` 等多个 serializer 的控制流刻意贴合它的对应 renderer，方便对照。

---

## 3. 目录结构

```
src/
├── index.ts                # 对外入口：parse() + 类型导出
│
├── adapter/                # 对外 JSON 形状层（PPTist 契约）
│   ├── types.ts            # ★ 输出 JSON 类型定义；改它即改协议
│   └── toPptxtojson.ts     # 组装 Output（slides + themeColors + size）
│
├── parser/                 # 字节/XML 层（不感知 OOXML 语义）
│   ├── ZipParser.ts        # .pptx → 分类后的 PptxFiles，含安全限额
│   ├── XmlParser.ts        # SafeXmlNode：null-safe DOM 包装
│   ├── RelParser.ts        # .rels → rId → target 映射
│   └── units.ts            # EMU / pt / px / 角度 / 百分比换算
│
├── model/                  # 几何与结构模型（不感知样式）
│   ├── Theme.ts            # 主题色与字体
│   ├── Master.ts           # slideMaster
│   ├── Layout.ts           # slideLayout（含 placeholder 表）
│   ├── Slide.ts            # slide spTree → SlideNode[]
│   ├── Presentation.ts     # 装配三层链 → PresentationData
│   └── nodes/
│       ├── BaseNode.ts     # 共用属性 parseBaseProps
│       ├── ShapeNode.ts    # p:sp / p:cxnSp（auto-shape / 文本框 / 连接器）
│       ├── PicNode.ts      # p:pic（图 / 视频 / 音频占位）
│       ├── TableNode.ts    # p:graphicFrame > a:tbl
│       ├── ChartNode.ts    # p:graphicFrame > c:chart
│       ├── GroupNode.ts    # p:grpSp（含 chOff/chExt）
│       └── MathNode.ts     # mc:AlternateContent > 公式
│
├── serializer/             # 模型 → JSON 元素
│   ├── index.ts            # 对外汇总导出
│   ├── RenderContext.ts    # 每页一份的解析上下文
│   ├── slideSerializer.ts  # 编排：bg → master → layout → slide
│   ├── shapeSerializer.ts  # ★ Shape vs Text 判定、preset 路径、自适应
│   ├── textSerializer.ts   # TextBody → HTML 富文本字符串
│   ├── tableSerializer.ts  # 表格样式级联（含 vAlign / 行高 / 列宽）
│   ├── chartSerializer.ts  # 各类图表的 series/colors/category 提取
│   ├── imageSerializer.ts  # 图 / 视频 / 音频 / OLE 预览
│   ├── mathSerializer.ts   # OMML → LaTeX
│   ├── groupSerializer.ts  # ★ chOff/chExt 缩放 + flip/rotation 烘焙
│   ├── backgroundSerializer.ts
│   ├── borderMapper.ts     # 线型 → CSS dasharray
│   └── StyleResolver.ts    # 颜色 / 填充 → CSS
│
├── shapes/                 # SVG path 生成
│   ├── presets.ts          # 200+ OOXML preset 几何（line / star / wedgeEllipseCallout …）
│   ├── customGeometry.ts   # a:custGeom → SVG path
│   └── shapeArc.ts         # OOXML arc → SVG arc
│
├── utils/                  # 通用工具
│   ├── color.ts            # OOXML 颜色变换（lumMod/Off/satMod/tint/shade …）
│   ├── media.ts            # MIME 推断 / 路径解析
│   ├── mediaWebConvert.ts  # TIFF/EMF/JXR → PNG（让 PPTist 能渲染）
│   ├── emfParser.ts        # EMF 内嵌 PDF / DIB 提取
│   ├── rgbaToPng.ts        # Canvas 编码 PNG
│   └── urlSafety.ts        # 外链协议白名单
│
├── export/
│   └── serializePresentation.ts  # 调试：去掉 SafeXmlNode 引用、扁平化 group
│
└── types/
    └── vendor-shims.d.ts
```

带 ★ 的是定位问题时最常进入的几个文件。

---

## 4. 数据结构关键约定

### 4.1 输出形状（节选）
完整定义见 `src/adapter/types.ts`。每个元素都有 `type` 区分：

| `type`    | 关键字段                                                         |
| --------- | ---------------------------------------------------------------- |
| `text`    | `content`(HTML) / `vAlign` / `isVertical` / `autoFit`            |
| `shape`   | `shapType` / `path` / `keypoints` / `content` / `vAlign`         |
| `image`   | `src` / `geom` / `rect` / `filters`                              |
| `table`   | `data[][]` / `borders` / `rowHeights` / `colWidths`              |
| `chart`   | `chartType` / `data` / `colors` / `barDir` / `marker` / `style` |
| `group`   | `elements[]`                                                     |
| `math`    | `picBase64` / `latex` / `text`                                   |
| `video`/`audio` | `src` / `blob`                                              |

**所有 `left/top/width/height` 单位都是 pt**。

### 4.2 元素顺序
每个元素都有 `order` 字段（来自 `BaseNode.xmlOrder`），用于下游恢复 z-order。

### 4.3 Group 的子元素坐标
`Group.elements[]` 中的 `left/top` **已经经过 `chOff/chExt` 转换**，是相对于"外层 group 起点"的坐标，可直接用 `group.left + child.left` 得到绝对坐标。同时 group 自身的 `rotate/isFlipH/isFlipV` 已被烘焙进子元素，外层 group 始终中性。

---

## 5. 脚本

| 命令 | 用途 |
| --- | --- |
| `npx tsx scripts/transvert.ts <a.pptx> [out.json]` | 用本库（`src/`）解析，**开发主力**：直接跑源码无需打包，方便边改边验。不传 `out.json` 时输出到 stdout。 |
| `npx tsx scripts/transvert.js <a.pptx> [out.json]` | 用 `src1/` 原版解析，**对照基准**：当怀疑新版输出有回归时，跑同一个 .pptx 比对两份 JSON 的差异。 |
| `node scripts/extract-pptx-structure.js <a.pptx> [outDir]` | 解压 .pptx 看内部 zip 结构，可选释放到目录。定位"源 XML 长什么样"时用。 |
| `pnpm build` | Rollup 打包 + 生成 `.d.ts` → `dist/`。 |
| `pnpm lint` | ESLint 检查 `src/`。 |

> 推荐工作流：拿到 bug 报告 → `extract-pptx-structure.js` 看源 XML → `transvert.ts` 生成 JSON 比对 → 改 `src/` → 再跑一次 `transvert.ts` 验证。

---

## 6. 与下游的契约

`adapter/types.ts` 是与 PPTist（及其它消费方）的**协议**。任何字段的增删改都会影响下游。新增字段时优先选择**可选字段**（`?:`）以保持向后兼容，并在 `README.md` 的"完整功能支持"里同步登记。
