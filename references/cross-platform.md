# XML 后端：不依赖 Visio 的跨平台出图

`scripts/xml_backend/` 直接生成 Visio 页面 XML 并打包为 vsdx，不需要 Windows、Visio 或 pywin32，Linux / macOS / Windows 都能运行。产物仍是原生可编辑的 Visio 对象：连接点、端点粘附（`PAR(PNT())` 公式 + `<Connects>`）、连线自带文本都按 Visio 规范写入，拖动形状时连线端点跟随。

## 与 COM 后端的分工
| | COM（visio_lib.py） | XML 后端 |
|---|---|---|
| 运行环境 | Windows + 桌面版 Visio | 任意平台，仅需 Python 3.8+ 和 Pillow |
| 速度与确定性 | 较慢，受 Visio 版本影响 | 秒级，同样输入同样输出，可做差异比对 |
| Word 预览 | 矢量 EMF | 高 DPI PNG（在 Visio 里双击编辑后自动重新生成） |
| 公式缓存值 | Visio 当场重算 | 脚本自己算好写入 |
| 适合 | 需要 EMF、或在原图上局部修改 | 批量生成、整页重绘、CI、AI Agent 沙箱 |

两者可以组合：用 XML 后端生成，在装有 Visio 的 Windows 上用 `visio_finish.py` 做验证与收尾（重算、导出 EMF、Visio 另存），这一步失败不影响交付。

## 用法
```bash
pip install pillow
# 1. 独立 vsdx + PNG 预览
python scripts/xml_backend/build_vsdx.py assets/example_xml_flowchart.py out/

# 2. 重绘 docx 里已有的嵌入图（原稿只读）
python scripts/xml_backend/inspect_visio_docx.py 方案.docx              # 诊断现状
python scripts/xml_backend/extract_vsdx.py 方案.docx work/templates     # 解出逐图模板，打印“顺序 → 嵌入后缀”
#    写 work/figspecs.py：FIGS = [(图号, 嵌入后缀, 生成函数), ...]
python scripts/xml_backend/build_docx.py 方案.docx work/figspecs.py work/templates out.docx --preview work/png
python scripts/xml_backend/inspect_visio_docx.py out.docx               # 复查

# 3. 测试（不需要 Visio）
python -m unittest discover -s tests -v

# 4. 可选：Windows + Visio 验证与收尾（需 pywin32）
python scripts/xml_backend/visio_finish.py out.docx --verify-only --report r.json
python scripts/xml_backend/visio_finish.py out.docx final.docx
```

## 主题
主题只换颜色，不换语义角色、版式、字号与连线规则；定义在 `scripts/theme.py` 的 `XML_PALETTES`，与 COM 主题共用注册表（`theme.py` 的 `THEMES`），两条路线同名同义。默认 `tech-cn`，不写主题时外观不变。

| 主题 | 适用 | 主色 |
|---|---|---|
| `tech-cn`（默认） | 中文工程文档 | 深蓝 #1F3864 |
| `mono` | 严格黑白印刷 | 全灰阶；ok/warn 只用填充深浅区分 |

```python
f = Fig(150, 70, theme="mono")   # 单图（黑白印刷）
```
```bash
python scripts/xml_backend/build_docx.py 方案.docx work/figspecs.py work/templates out.docx --theme mono
VISIO_THEME=mono python scripts/xml_backend/build_vsdx.py assets/example_xml_flowchart.py out/   # 环境变量也行
```
一份文档只用一套主题（见 `references/visio-style-rules.md` §6）。ok/warn 是语义色，除 mono 外各主题相同；每套主题的文字/填充对比度均 ≥ 4.5:1（见 `tests/test_xml_backend.py` 的对比度测试）。

## 文件
- `vlib.py`：绘图 DSL（box / text / container / line / conn / legend），输出 Visio 页面 XML 与 PNG 预览；单位 mm，左上原点，Y 向下。
- `vsdxpkg.py`：模板目录 + 新 page1.xml → vsdx，按图尺寸改页面大小。
- `extract_vsdx.py`：从 docx 解出全部嵌入 vsdx，保留每张图原有的样式表与窗口部件。
- `build_docx.py`：逐图替换嵌入 vsdx 与预览，设置对象尺寸并锁定纵横比。
- `build_vsdx.py`：不经 docx，直接出独立 vsdx。
- `inspect_visio_docx.py`：只读诊断（尺寸、缩放、拉伸、字号、未粘附连线等）。
- `visio_finish.py`：Windows + Visio 的缓存值核对、粘附核对、连线几何首末点核对、自动拖动测试（拖动后同时核对端点与几何）、重算、EMF 导出、另存与回写。
- `assets/templates/blank_vsdx/`：空白模板（单页、无形状、已去作者信息与缩略图）；原稿没有对应嵌入时自动使用。只支持单页模板，多页模板会直接报错。
- `tests/test_xml_backend.py`：冒烟与回归测试（粘附、零宽/零高连线几何公式、标题带不参与粘附、对齐、错误提示、合成 docx 往返、模板目录清理）。

## 字体
默认字体按操作系统选取（`scripts/font_support.py`，COM 与 XML 两条路线共用）：Windows＝Arial + Microsoft YaHei；macOS＝Helvetica Neue + PingFang SC；Linux＝fontconfig 检测的系统字体（如 Noto Sans CJK）。PNG 预览用同一字体；Visio XML 里 `Font`/`AsianFont` 分别写 Latin/CJK 字体族。可用环境变量覆盖：`VISIO_FONT_LATIN` / `VISIO_FONT_CJK`（字体族）与 `VISIO_FONT_{LATIN,CJK}{,_BOLD}_FILE`（预览度量字体文件）；旧变量 `VISIO_KIT_FONT_R/B`（.ttc 序号 `VISIO_KIT_FONT_INDEX`）仍兼容。找不到字体时报错，不静默换用无关字体。

## 已知限制
- 只做整页重绘，不在原图上局部修改，也不处理主控形状、主题继承和模具。
- 连线是带显式几何行的 1-D 形状，每个几何点都写公式：拖动时端点和线身一起跟随，中间折点按比例伸缩但不会自动重新走线（与 COM 后端的 connect() 一致）。
- 复杂公式的缓存值由脚本计算，建议有条件时用 `visio_finish.py --verify-only` 核对。
- `visio_finish.py` 已在真实 Visio 上回归（2026-11-19，Visio Pro + Win10）：粘附核对 / 拖动测试 / 缓存值 / EMF 导出 / 回写全链通过；EMF 导出按「白色不透明页框垫底」保证图幅=页面（透明框不扩大导出范围）。

完整的风格规则与审查清单见 `references/visio-style-rules.md`。
