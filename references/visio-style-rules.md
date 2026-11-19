# Visio 嵌入对象通用修改指南（给 LLM / Agent）

版本：v2.3（2026-11-19）
适用范围：任何 Word（.docx）文档中嵌入的 Visio 对象（.vsdx OLE 对象），以及独立的 .vsdx 图稿。与具体项目无关。
读者：负责检查、统一风格、重绘或修改这类图的 LLM / Agent，以及人工编辑者。
用法：先读第 1～3 节确定流程，改图时按第 4～9 节的视觉规则，动文件时按第 10 节的结构规则，工具见第 11～12 节，交付前过第 13 节清单。用户的明确要求优先于本指南；本指南的默认值（页宽、色板、字体）可按文档或用户要求整体替换，但一份文档内必须只用一套。

> 项目特有的信息（图号与嵌入文件的对应、每张图的布局选择、修改记录）不写在本文件里，写在项目自己的台账文件中（参考做法见第 14 节）。

---

## 1. 总原则

1. **内容不变，形式统一。** 保留原图的全部实体、关系和文字；只有修正明显错误（图文不符、悬空箭头、重复编号）时才改内容，并逐条记录。
2. **一图一意。** 每张图只回答一个问题（组成、流程、时序、对应关系），不混用多种画法。
3. **与正文一致。** 图中名称、编号、流程顺序与正文和图题一致；不一致时优先改图对齐正文，并记录。
4. **保持可编辑。** 图必须仍是双击可在 Visio 中编辑的嵌入对象，不得替换成纯图片。
5. **预览与源同步。** Word 显示的是对象的预览图，不是 vsdx 本身；改了 vsdx 必须同时更新预览，否则用户看到的仍是旧图。
6. **可复现。** 改图通过脚本从原稿生成修改稿，原稿只读；任何人重跑脚本得到同样结果。

## 2. 端到端工作流程

1. **备份与解包**：原稿另存为只读副本；docx 是 zip，解包到临时目录。
2. **盘点**：运行诊断脚本（第 11.3 节）列出每个对象的嵌入文件、预览文件、页面尺寸、Word 显示尺寸、字号、连线问题；按文档顺序建立“图号 → 嵌入文件”对照表，写入项目台账。
3. **目检**：把每张图的预览渲染成 PNG，逐张用看图工具检查重叠、溢出、悬空、压线。
4. **出诊断报告**：先给用户“全局问题 + 逐图问题 + 建议”，得到同意后再改。
5. **定规范**：确定页宽、字体、色板（默认用本指南第 4～8 节）。
6. **逐图重绘**：用布局 DSL 描述每张图（第 12 节），每张选一个布局模板（第 9 节）。
7. **回写**：替换 vsdx 页面 XML、页面尺寸、预览图、Word 中的对象尺寸（第 10 节）。
8. **自检**：XML 可解析、诊断脚本无问题、每张预览目检通过（第 13 节）。
9. **交付**：把修改稿发给用户，说明改了什么、仍有什么限制（例如未在真实 Visio 中做拖动测试、哪些连线按设计不粘附）。
10. **记录**：在项目台账追加修改记录（日期、版本、逐图改动）。

## 3. 诊断要点（常见问题与判定方法）

| 问题 | 怎么发现 | 处理 |
|---|---|---|
| 字号偏小且各图不一 | 页内字号 × Word 缩放比 = 有效字号；< 7.5 pt 或各图差异大 | 统一页宽 = 版心宽，100% 插入，统一字号 |
| 非等比拉伸 | Word 显示宽/页宽 与 显示高/页高 相差 > 2% | 重置为 100% 并锁定纵横比 |
| 被二次缩放 | Word 显示宽 ≠ 页宽 | 页宽改成版心宽度后 100% 插入 |
| 斜线段 | 连线几何相邻两点 x、y 同时变化 | 改正交走线 |
| 悬空箭头 / 拖动不随动 | 连线没有 `Connect` 记录；或有记录但 BeginX/EndX 没有粘附公式；或端点不在形状边上 | 端点落到形状边，按 7.2 写连接点、粘附公式和 Connect |
| 连线文字是独立文本框 | 连线附近有无边框的小文本形状，内容是“通过/发送”等 | 删掉文本框，文字写进连线自身的 `<Text>`（7.3） |
| 从线中间引出的分支线 | 分支起点落在另一条线上，不在任何形状边上 | 改为从源形状边上引出，与主线重叠走同一段（7.2） |
| 圆角滥用 | `Rounding` ≠ 0 且不是流程终结符 | 改为 0 |
| 图例、图上说明文字 | 图内有“注：”、图例短线 | 删除，说明放正文；关系用节点表达 |
| 字体公式错误 | 页面 XML 中出现 `#VALUE!` / `#REF!` / `#NAME?` | 所有 Character 行写明 Font 与 AsianFont |
| 连线文字压线、压边框 | 目检 | 文字放最长线段中点，白底（7.3） |
| 文字溢出、挤 | 目检；估算文字宽度 > 框宽 | 加宽框或按意义换行 |
| 孤立形状 | 无任何连线且非容器/说明 | 改为辅助样式或补关系 |
| 缺字符号 | ▾、☐ 等显示为方框 | 换成文字或缩进 |

注意：时间轴、刻度线、分隔线这类端点处本来就没有形状的线，诊断时“未粘附连线”可忽略；其余未粘附的都要处理。

## 4. 版面与尺寸

| 项目 | 规则 |
|---|---|
| 页面宽度 | 等于 Word 版心宽度（A4 常见 150～170 mm）；一份文档内所有图相同 |
| 页面高度 | 按内容定，取 5 mm 整数倍；常见 80～150 mm |
| 插入 Word | **100% 比例**，宽 = 页宽(mm)/25.4×72 pt，高 = 页高(mm)/25.4×72 pt，锁定纵横比 |
| 外边距 | 内容距页边 ≥ 4 mm |
| 网格 | 坐标、尺寸对齐到 0.5 mm；同类元素对齐到同一基线 |
| 元素间距 | 同排框间距 2.5～8 mm；层（容器）之间 8～10 mm |
| 禁止 | 非等比拉伸；在 Word 中二次缩放导致字号漂移 |

## 5. 字体与字号

| 用途 | 字体 | 字号 | 字重 | 颜色 |
|---|---|---|---|---|
| 普通框文字 | 中文文档：微软雅黑（Latin 名 Microsoft YaHei）；英文文档可用 Segoe UI / Arial | 9 pt | 常规 | #1A1A1A |
| 重点框文字 | 同上 | 9 pt | 加粗 | #1F3864 |
| 容器标题（标题带） | 同上 | 9 pt | 加粗 | #FFFFFF |
| 图内标题 | 同上 | 10 pt | 加粗 | #1F3864 |
| 连线文字（仅判定分支） | 同上 | 8 pt | 常规 | #404040 |
| 拥挤处下限 | — | 不低于 7.5 pt（以 Word 中有效字号计） | — | — |

- 每个形状的 Character 段写明 `Font` 和 `AsianFont`，不留公式错误。
- 多行文字按意义换行，每行完整，不溢出边框。
- 不用目标字体缺字的符号；树形层级用缩进表示。
- 渲染预览时若无微软雅黑，用度量接近的 Noto Sans SC 代替（见 11.2）。

## 6. 配色与形状样式（默认色板）

一份文档只用一套色板，不引入色板外的颜色。默认工程风格色板：

| 样式名 | 填充 | 边框 | 线型 | 文字 | 用途 |
|---|---|---|---|---|---|
| node 普通框 | #FFFFFF | #2E5C8A 0.75 pt | 实线 | #1A1A1A | 一般模块、步骤 |
| key 重点框 | #DCE8F5 | #2E5C8A 0.75 pt | 实线 | #1F3864 加粗 | 核心模块、主链路节点 |
| dark 关键框 | #1F3864 | #1F3864 | 实线 | #FFFFFF 加粗 | 每图最关键的 1～2 个节点 |
| ext 外部框 | #F2F2F2 | #7F7F7F 0.75 pt | 虚线 | #404040 | 外部系统、被测对象、外部输入 |
| soft 辅助框 | #F3F6FA | #A6B9D0 0.5 pt | 实线 | #1F3864 | 说明条、环境、次要信息 |
| ok 正向结果 | #E2F0D9 | #548235 | 实线 | #385723 | 一致、通过 |
| warn 异常结果 | #FBE5D6 | #C55A11 | 实线 | #843C0C | 偏差、缺失、拒绝 |
| container 容器 | #F7F9FC | #8EA9C1 0.75 pt | 实线 | — | 分层、分组 |
| container-ext 外部容器 | 无 | #8EA9C1 | 虚线 | — | 外部区域 |
| band 标题带 | #2F5597 | #2F5597 | — | #FFFFFF 加粗 | 容器左上角标题，高 6.5 mm |
| lane header 泳道头 | #1F3864 | — | — | #FFFFFF 加粗 | 泳道左侧标题，宽约 22 mm |

- 圆角：矩形一律直角（`Rounding=0`）；仅流程图的开始/结束终结符使用圆角（胶囊形）。
- 保持简洁：图上不放图例、注释和说明性文字；连线只在判定分支上标注条件（如“通过/不通过”）。原来靠连线标注表达的含义（通路名称、数据内容）改为节点表达；说明文字放正文。图号与图名只放在正文题注里。
- 无阴影、无渐变、无 3D。
- dark 每图 ≤ 2 个；ok / warn 只用于结果分类，不作装饰。

## 7. 连线

### 7.1 线型与走线
| 类型 | 颜色 | 线型 | 线宽 | 用途 |
|---|---|---|---|---|
| 主数据流 / 控制流 | #2E5C8A | 实线 | 1 pt | 正向流程 |
| 回传、留档、上报、离线 | #7F7F7F | 虚线 | 1 pt | 反向或非实时数据 |
| 闭环迭代 | #2F5597 | 虚线 | 1 pt | 回到起点的迭代 |
| 异常分支 | #C55A11 | 虚线 | 1 pt | 拒绝、失败分支 |

- 一律**正交（直角）走线**，无斜线段；折点 ≤ 3 个。
- 箭头：实心三角（Visio `EndArrow=13`，尺寸 1）。
- 不穿过无关形状、标题带或文字；绕行走层间空隙。
- 多线汇合/分叉用公共水平“总线”段，分叉点放层间空隙。
- 一对多支撑关系难画时，用“组到组”的一条关系（例：支撑模块组容器 → 处理链容器）代替多条悬空箭头。

### 7.2 端点粘附（拖动形状时端点随动）
目标：用户在 Visio 中拖动任一形状，连在它上面的连线端点跟着移动。只写 `<Connects>` 不够，Visio 靠端点单元格上的公式实现随动。

1. **每个有边框的形状（矩形、容器）都带连接点**：`<Section N="Connection">`，默认四个边中点，另把实际有连线落点的位置补成连接点。每行写公式，形状缩放时连接点跟着走：
   ```xml
   <Section N="Connection">
     <Row IX="0"><Cell N="X" V="0.59" F="Width*0.5"/><Cell N="Y" V="0.47" F="Height*1"/>
       <Cell N="DirX" V="0"/><Cell N="DirY" V="0"/><Cell N="Type" V="0"/><Cell N="AutoGen" V="0"/></Row>
     …
   </Section>
   ```
   第 IX 行在公式里叫 `Connections.X{IX+1}` / `Connections.Y{IX+1}`。Connection 段放在 Character、Paragraph 之后、Geometry 之前。
2. **连线端点写粘附公式**（点到点粘附，保留原来的出线位置）：
   ```xml
   <Cell N="BeginX" V="1.23" F="PAR(PNT(Sheet.3!Connections.X2,Sheet.3!Connections.Y2))"/>
   <Cell N="BeginY" V="2.10" F="PAR(PNT(Sheet.3!Connections.X2,Sheet.3!Connections.Y2))"/>
   ```
   EndX/EndY 同理。`Sheet.n` 是目标形状的 Shape ID，所以先输出所有二维形状、再输出连线。
3. **同时写 `<Connects>`**，与公式一一对应：
   ```xml
   <Connect FromSheet="20" FromCell="BeginX" FromPart="9" ToSheet="3" ToCell="Connections.X2" ToPart="101"/>
   ```
   `FromPart`：起点 9、终点 12；`ToPart` = 100 + 连接点行号 IX。
4. **连线自身的尺寸用公式跟随端点**（不旋转，Angle=0）：`PinX=(BeginX+EndX)/2`、`PinY=(BeginY+EndY)/2`、`Width=EndX-BeginX`、`Height=EndY-BeginY`、`LocPinX=Width*0.5`、`LocPinY=Height*0.5`。局部坐标原点即起点，终点在 (Width, Height)，宽高可为负或 0。
5. **折线几何每个点都写公式**：X 写 `Width*(x/W)`、Y 写 `Height*(y/H)`。W 或 H 为 0 的方向（纯竖线的 X、纯横线的 Y）不能只写数值，否则横向拖动目标后端点跟过去、线身却不动：首点写 `Width*0`、末点写 `Width*1`，中间点按序号就近跟随首端或末端并保留原偏移（如 `Width*0+5 mm`）；`TxtPinX/Y` 同理，零跨度时跟随中点。这样 L 形、Z 形折线在端点移动后仍保持直角，直线变斜线时线身与端点不脱节。
6. **设为可路由连线**：`ObjType=2`、`ShapeRouteStyle=1`（直角）、`ConLineRouteExt=1`、`ConFixedCode=0`。Visio 重新布线时可能调整折点位置，属正常现象。
7. **端点必须精确落在形状边框上**（容差 ≤ 0.3 mm），否则找不到粘附目标。自动识别时取边框包含该点的最小形状，不粘到纯文字框。
8. **不要从线中间引出分支**：线中点不是形状，无法粘附。分支线改为从源形状边上引出，前一段与主线重叠走同一路径，外观不变，两端都能粘附。
9. 只有端点处确实没有形状时才不粘附：时间轴、刻度线、分隔线，以及指向分隔线的箭头那一端。

### 7.3 连线文字
- 连线文字写在**连线自身的文本**里，不另建文本框；这样拖动连线或形状时文字随线移动，删除连线时一起删除。
- 写法：在连线 Shape 中加 `TxtPinX/TxtPinY`（按 Width/Height 比例写公式，零跨度方向见上条，放最长线段中点）、`TxtWidth/TxtHeight`、`TxtLocPinX=TxtWidth*0.5`、`TxtLocPinY=TxtHeight*0.5`、`TxtAngle=0`、白色文字背景 `TextBkgnd`（`F="RGB(255,255,255)"`）、`TextBkgndTrans=0`，文本边距 0；再加 Character（8 pt）、Paragraph（居中）段，最后 `<Text>`。
- 只在判定分支上用（7.1 之外的含义改为节点）。

## 8. 标题、编号与说明

- 图号与图名由 Word 题注提供，图内不重复图号。
- 不加图例和图上注释；线型含义在正文说明，或改用节点表达。
- 步骤编号用 ①②③…；同一图内只有一套编号。
- 图中术语与正文一致：正文口径变了（例如某功能从“存储并查找”改为“按模型推演”），相关图的节点名同步改，并记录。

## 9. 布局模板（按图类型）

| 图类型 | 布局 |
|---|---|
| 分层架构图 | 自上而下的容器层（如 客户端 → 服务 → 运行环境），层内模块等宽横排，层间箭头不加标签 |
| 系统组成 / 网络图 | 本系统放左/上容器，外部对象用 ext 放右/下；每台终端设备都要有连线接入网络，不留孤立设备 |
| 线性流程图 | 单列或单行等宽步骤框，主线居中；回路沿外侧走；只有开始/结束用圆角 |
| 判定 / 分支流程 | 输入 → 判定层 → 结果层（ok/warn）→ 汇总 → 输出，分叉用总线；分支条件写在连线文本里 |
| 双方交互图 | 左右两个等宽容器，同行的框一一对应，水平箭头穿过中缝 |
| 泳道 / 对应关系图 | 横向泳道 + 左侧深色泳道头；或多列表头 + 行对齐 |
| 时序图 | 分段容器；时间轴为深色实线箭头，事件用短竖箭头，刻度对齐 |
| 界面示意图 | 窗口外框直角；标题栏 dark，菜单栏/状态栏 soft，按钮左对齐，页签等宽网格，树形缩进 5 mm |

## 10. 文件结构与技术规则

### 10.1 docx 中的 Visio 对象
- 每个对象在 `word/document.xml` 中是一个 `<w:object>`，内含：
  - `<v:shape style="width:…pt;height:…pt">` 显示尺寸，`<v:imagedata r:id>` 指向预览图（原稿常为 EMF）；
  - `<o:OLEObject ProgID="Visio.Drawing.15" r:id>` 指向 `word/embeddings/Microsoft_Visio_DrawingN.vsdx`；
  - `w:dxaOrig / w:dyaOrig` 原始尺寸，单位 twips（mm/25.4×1440）。
- rId 到文件的映射在 `word/_rels/document.xml.rels`。按 `<w:object>` 在 document.xml 中出现的顺序即文档顺序；嵌入文件编号（Drawing、Drawing1…）**不一定**等于文档顺序，必须按 rId 实际对照。

### 10.2 修改一张图必须同时做
1. 改 vsdx 内 `visio/pages/pageN.xml`（形状、连线、Connects）。
2. 改 `visio/pages/pages.xml` 的 `PageWidth` / `PageHeight`（单位英寸）。
3. 生成新预览 PNG（≥ 250 dpi），放入 `word/media/`，把对应 imagedata 关系的 Target 改为它；删除不再被引用的旧预览文件。
4. 改 `<v:shape>` 的 `style` 宽高（pt），`w:dxaOrig/w:dyaOrig`（twips），把 `<v:shape>` 内的 `<o:lock aspectratio="f"/>` 改为 `"t"`。
5. `[Content_Types].xml` 确保有 `<Default Extension="png" ContentType="image/png"/>`。
6. 重新打包 vsdx 与 docx 时，`[Content_Types].xml` 放第一个条目，其余尽量保持原顺序。
7. 保留 vsdx 中原有的 document.xml、样式表、主控形状等部件，只替换页面内容，避免 Visio 打开报错。

### 10.3 Visio 页面 XML 要点
- 命名空间 `http://schemas.microsoft.com/office/visio/2012/main`；单位英寸；**Y 轴向上**（从上到下排版时 y_visio = 页高 − y）。
- 二维形状：`Type="Shape"`，样式引用 `LineStyle/FillStyle/TextStyle="0"`，格式全部写本地单元格（FillForegnd、LineColor、LinePattern、LineWeight、Rounding 等）。
- 顺序：Cell 在前，Section（Character、Paragraph、Geometry）在后，`<Text>` 最后。
- Character 行写 `Font`、`AsianFont`、`Size`（英寸，pt/72，`U="PT"`）、`Color`、`Style`（1=加粗）、`LangID`。
- 连线：一维形状，写 `BeginX/BeginY/EndX/EndY`、`ObjType=2`、`EndArrow=13`、折线 Geometry；端点粘附、尺寸公式、连线文字按第 7.2、7.3 节写。
- 只有 `<Connects>` 而端点单元格没有 `PAR(PNT(...))` 公式时，Visio 中拖动形状端点不会随动；这是最常见的“看着连上了其实没粘住”。
- 容器的标题带若是独立形状，拖动容器时不会跟随；需要整体移动时，把标题写进容器自身文本，或在 Visio 中组合。
- 无 Visio 的环境里无法做拖动测试；有 Windows + Visio 时用 `visio_finish.py --verify-only` 自动拖动验证，否则请使用者双击图、拖一个形状确认。

## 11. 工具链（本指南实际使用过的工具）

### 11.1 Agent 使用
任何能执行 shell、读写文件、以图片方式查看 PNG 的 LLM Agent 都可按本指南工作；每张图改完都要看一遍预览 PNG。

### 11.2 运行环境与库
- Python 3 标准库：`zipfile`（docx/vsdx 解包打包）、`re`（定位 `<w:object>`、改尺寸属性）、`xml.etree.ElementTree` / `lxml`（解析与校验 XML）、`xml.sax.saxutils.escape`（转义形状文字）、`math`（连线角度与局部坐标）。
- Pillow（PIL）：渲染预览 PNG（`ImageDraw.rounded_rectangle`、`text(anchor='mm')`、手工虚线）。
- 字体：Noto Sans SC Regular / Bold（`vlib.py` 自动查找微软雅黑 / 苹方 / Noto CJK，或用环境变量 `VISIO_KIT_FONT_R`、`VISIO_KIT_FONT_B` 指定），仅作 PNG 预览时微软雅黑的渲染替身；Visio XML 里仍写微软雅黑。
- 没有 Visio / LibreOffice 也能完成：预览由脚本自己画，不依赖 Office 渲染 vsdx。

### 11.3 XML 后端脚本（scripts/xml_backend/）
| 文件 | 作用 | 用法 |
|---|---|---|
| `scripts/xml_backend/extract_vsdx.py` | 从原稿 docx 解出全部嵌入 vsdx，作为该文档的**模板目录**（保留每张图原有样式表、窗口等部件），并打印正文顺序→嵌入后缀对照 | `python3 extract_vsdx.py 原稿.docx work/templates` |
| `assets/templates/blank_vsdx/` | 空白 vsdx 模板（单页、无形状、样式表含 vlib 用到的 0 号样式、已去作者信息和缩略图） | 新增图或原稿没有对应嵌入时自动使用 |
| `scripts/xml_backend/inspect_visio_docx.py` | **通用诊断**：逐个对象列出页面尺寸、Word 尺寸、缩放、非等比拉伸、纵横比锁定、公式错误、斜线、未粘附连线（按 Connects 统计）、最小有效字号 | `python3 inspect_visio_docx.py file.docx` |
| `scripts/xml_backend/vlib.py` | **通用绘图 DSL**：样式色板、容器、正交连线 → Visio 页面 XML + PNG 预览；自动给形状加连接点、把落在边框上的连线端点粘附上去、连线文字写入连线自身 | 见第 12 节 |
| `scripts/xml_backend/vsdxpkg.py` | vsdx 打包库：模板目录副本 + 新 page1.xml + 按图尺寸改页面大小 → vsdx | 被下面两个脚本调用 |
| `scripts/xml_backend/build_docx.py` | 回写 docx：逐图替换嵌入 vsdx、预览 PNG、Word 对象尺寸与纵横比锁定、Content_Types | `python3 build_docx.py 原稿.docx figspecs.py work/templates 输出.docx --preview work/png` |
| `scripts/xml_backend/build_vsdx.py` | 不经过 docx，直接生成独立 .vsdx + PNG 预览 | `python3 build_vsdx.py figspecs.py out/ [图号...]` |
| `scripts/xml_backend/visio_finish.py` | **Windows + 本机 Visio 的验证与收尾（可选）**：缓存值核对、粘附核对、连线几何首末点核对、自动拖动测试；强制重算、导出 EMF（纵横比不符则保留 PNG）、Visio 另存 vsdx、回写 docx。单图失败不影响其余，有问题时退出码 1 | `python visio_finish.py 生成.docx 终稿.docx --report r.json`；只验证加 `--verify-only` |
| `assets/example_xml_flowchart.py` | 最小示例：带判定分支的流程图 | `python scripts/xml_backend/build_vsdx.py assets/example_xml_flowchart.py out/` |

### 11.4 分工
- 生成一律用 XML 后端，任何平台都能跑，结果确定、可比对。
- 有 Windows + Visio 时，交付前跑 `visio_finish.py`：先 `--verify-only` 看报告，无问题再正式收尾。没有 Visio 时直接交付 build_docx 的输出（PNG 预览）。
- 只有带主控形状、主题、需局部修改的复杂原图，才直接用 COM 改原图。

## 12. 绘图 DSL 用法（`vlib.py`）

坐标单位 mm，原点左上，y 向下；导出时自动转为 Visio 英寸、Y 向上。

```python
from vlib import Fig, LINE, GRAY, BLUE
f = Fig(170, 110)                                   # 页宽 170 mm，页高 110 mm
f.container('sys', 4, 4, 162, 60, '本系统', tw=30)  # 容器 + 左上角标题带（dashed=True 为外部容器）
f.box('a', 10, 18, 30, 12, '配置导入')              # 样式默认 node
f.box('b', 50, 18, 40, 12, '处理核心', 'key')       # 样式：node/key/dark/ext/soft/ok/warn
f.conn('a', 'r', 'b', 'l')                          # 从 a 右边中点到 b 左边中点，自动正交
f.conn('b', 'r', 'c', 'l', label='通过')            # 判定分支：label 写进连线自身文本，自动放最长段中点
f.conn('c', 'l', 'a', 'b', via=[(20, 70)], dashed=True)  # via 手工给折点
f.line([(x1, y1), (x2, y2)], src='a', dst='b')      # 完全手工折线；src/dst 可省，端点落在边框上会自动粘附
f.render_png('fig.png', S=10)                       # S=像素/mm，10 ≈ 254 dpi
print(f.glue_report())                              # (已粘附端点数, 端点总数)，核对未粘附的是否都是轴线/刻度
xml = f.visio_xml()                                 # 写入 vsdx 的 visio/pages/page1.xml
```

- `conn` 的边：`t/b/l/r`；`fa/fb` 为边上比例位置（0～1），用于同一边引出多条线。
- 手工折线的端点坐标必须正好落在形状边框上；分支线从源形状边上引出，不从另一条线中间引出。
- `legend()`、`text(..., 'note')` 仍保留在库里，但按第 6、8 节默认不用。
- 辅助函数思路：`row(n, x0, x1, gap)` 计算等宽横排；`chain(f, ids)` 依次左右相连。
- 字体自动查找；找不到时设环境变量 `VISIO_KIT_FONT_R` / `VISIO_KIT_FONT_B`，无需改代码。

## 13. 交付前自检清单

- [ ] 页宽 = 版心宽，Word 中 100% 且纵横比锁定（诊断脚本无“缩放/拉伸/未锁定”）
- [ ] 字体统一，正文 9 pt、连线文字 8 pt，有效字号 ≥ 7.5 pt，无公式错误
- [ ] 只用一套色板；dark ≤ 2 个
- [ ] 矩形直角（终结符除外）；无图例、无图上注释
- [ ] 连线全部正交、不穿越无关形状
- [ ] 两端有形状的连线全部粘附：BeginX/EndX 有 `PAR(PNT(...))` 公式，`<Connects>` 一一对应；未粘附的只剩轴线、刻度、分隔线
- [ ] 没有从线中间引出的分支
- [ ] 连线文字只在判定分支上，且写在连线自身文本里，没有独立的标签文本框
- [ ] 文字不溢出，层级缩进规整
- [ ] 与正文、图题一致；编号不冲突
- [ ] 预览图已同步更新，旧预览已删除
- [ ] 所有 XML 可解析；docx 能被 zip 正常读取
- [ ] 每张预览已用看图工具逐张目检
- [ ] 交付说明里请用户在 Visio 中拖动一个形状验证随动
- [ ] 在项目台账写好修改记录

## 14. 项目台账（建议做法）

每份被修改的文档单独建一个台账文件（例如 `PROJECT_FIGURES.md`），记录：
- 图号 → 嵌入文件对照（文档顺序）；
- 每张图采用的布局模板；
- 对本指南默认值的偏离（例如页宽 150 mm、换用其他色板）及原因；
- 修改记录：日期、版本、逐图改动、内容性修改的理由。

本指南只放通用规则；发现新的通用经验（新问题类型、新模板、新工具用法）时，更新本指南并在下方记录。

## 15. 指南变更记录

- v2.3（2026-11-19）：随 XML 后端一并提交；规则来自一份 17 图中文设计方案的统一重绘与评审修改。
