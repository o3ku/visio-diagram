# -*- coding: utf-8 -*-
"""
配色主题 —— 供 visio_lib 使用。

所有颜色均为 **Visio 公式字符串**（形如 "RGB(31,78,121)"），
直接写入 CellsU 的 FillForegnd / LineColor / Char.Color。
"""

import os
import sys

from font_support import DEFAULT_CJK_FONT, DEFAULT_LATIN_FONT, font_file as _fs_font_file

# 常见中文字体 → 字体文件（theme.FONT_FILES 的超集，供 XML/COM 两条路线共用）
KNOWN_CJK_FILES = {
    "微软雅黑": ("msyh.ttc", "msyhbd.ttc"),
    "宋体": ("simsun.ttc", "simsun.ttc"),
    "黑体": ("simhei.ttf", "simhei.ttf"),
    "等线": ("Deng.ttf", "Dengb.ttf"),
    "仿宋": ("simfang.ttf", "simfang.ttf"),
    "楷体": ("simkai.ttf", "simkai.ttf"),
}

# Backward-compatible override map. Entries may be extended by callers.
FONT_FILES = {
    "微软雅黑": "C:/Windows/Fonts/msyh.ttc",
    "Microsoft YaHei": "C:/Windows/Fonts/msyh.ttc",
    "宋体": "C:/Windows/Fonts/simsun.ttc",
    "SimSun": "C:/Windows/Fonts/simsun.ttc",
    "黑体": "C:/Windows/Fonts/simhei.ttf",
    "SimHei": "C:/Windows/Fonts/simhei.ttf",
    "等线": "C:/Windows/Fonts/Deng.ttf",
    "Arial": "C:/Windows/Fonts/arial.ttf",
}

_FONT_WARNED = set()


def font_file(font_name, warn=True, bold=False):
    """Resolve the requested family; never silently measure with a random font."""
    if isinstance(font_name, str) and font_name.strip('"').isdigit():
        font_name = DEFAULT_CJK_FONT
    path = FONT_FILES.get(font_name)
    if path and not bold and os.path.isfile(path):
        return path
    if font_name in KNOWN_CJK_FILES and os.name == "nt":
        regular, bold_name = KNOWN_CJK_FILES[font_name]
        path = os.path.join(os.environ.get("WINDIR", "C:/Windows"), "Fonts",
                            bold_name if bold else regular)
        if os.path.isfile(path):
            return path
    path = _fs_font_file(font_name, bold=bold)
    if path:
        return path
    if warn and font_name not in _FONT_WARNED:
        _FONT_WARNED.add(font_name)
        print("  !! 找不到字体 %r。请安装该字体，或配置 VISIO_FONT_*_FILE。" % font_name,
              file=sys.stderr)
    return None


class Theme(object):
    """一套配色。字段用 self.__dict__ 展开，便于任意扩展。"""

    def __init__(self, name, font=None, latin_font=None, **colors):
        self.name = name
        self.latin_font = latin_font or DEFAULT_LATIN_FONT
        self.font = font or DEFAULT_CJK_FONT
        self.__dict__.update(colors)

    def c(self, key, default=None):
        return getattr(self, key, default)

    def __repr__(self):
        return "<Theme %s font=%s>" % (self.name, self.font)


# ---------------------------------------------------------------- 内置主题

# 正式技术文档（白底 + 柔和色块 + 深灰边框，2026-10-10 按用户截图风格化）
TECH_CN = Theme(
    "tech-cn", DEFAULT_CJK_FONT,
    black="RGB(62,63,63)",          # 图线主色（截图深灰 #3E3F3F）
    dark="RGB(38,38,38)",           # 正文说明文字
    gray="RGB(89,89,89)",           # 次要标注
    mgray="RGB(150,150,150)",       # 分隔框 / 虚线框
    lgray="RGB(242,242,242)",
    white="RGB(255,255,255)",       # 画布底色
    accent="RGB(56,71,103)",        # 标题 / 强调（深蓝灰）
    accent_fill="RGB(218,226,243)", # 主填充（截图 #DAE2F3 淡蓝紫）
    green_fill="RGB(226,239,217)",  # 次级填充（截图 #E2EFD9 淡绿）
    cream_fill="RGB(255,242,204)",  # 强调 / 外部块（截图 #FFF2CC 米黄）
    lane="RGB(242,246,252)",        # 泳道 / 分区底色（近白微蓝）
    muted_fill="RGB(226,239,217)",  # 外部对象 / 非本方模块 → 淡绿
)

# 纯黑白灰 —— 需要严格黑白印刷时用
MONO = Theme(
    "mono", DEFAULT_CJK_FONT,
    black="RGB(0,0,0)",
    dark="RGB(51,51,51)",
    gray="RGB(102,102,102)",
    mgray="RGB(178,178,178)",
    lgray="RGB(238,238,238)",
    white="RGB(255,255,255)",
    accent="RGB(0,0,0)",
    accent_fill="RGB(221,221,221)",
    lane="RGB(250,250,250)",
    muted_fill="RGB(230,230,230)",
)

THEMES = {t.name: t for t in (TECH_CN, MONO)}


# ---------------------------------------------------------------- XML 后端色板
# 两条路线共用同一份主题注册：COM 用上面的 RGB(...) 字段，XML 用下面的十六进制
# 色板。语义角色（node/key/dark/ext/ok/warn/soft…）在各主题里保持一致，只换颜色；
# ok/warn 是语义色，除 mono（黑白印刷只能用灰阶深浅区分）外各主题相同。
# tech-cn 的 XML 色板沿用 vlib.py 原有值（不改默认外观）。
_OK = ("#E2F0D9", "#548235", 0, "#385723", 0, 9.0, 0.75)
_WARN = ("#FBE5D6", "#C55A11", 0, "#843C0C", 0, 9.0, 0.75)

XML_PALETTES = {
    # 默认：中文工程文档（vlib.py 原有色板，原样保留）
    "tech-cn": {
        "styles": {
            #        fill      line     dash  text     bold  size  lw
            "node":  ("#FFFFFF", "#2E5C8A", 0, "#1A1A1A", 0, 9.0, 0.75),
            "key":   ("#DCE8F5", "#2E5C8A", 0, "#1F3864", 1, 9.0, 0.75),
            "dark":  ("#1F3864", "#1F3864", 0, "#FFFFFF", 1, 9.0, 0.75),
            "ext":   ("#F2F2F2", "#7F7F7F", 1, "#404040", 0, 9.0, 0.75),
            "ok":    _OK,
            "warn":  _WARN,
            "soft":  ("#F3F6FA", "#A6B9D0", 0, "#1F3864", 0, 8.5, 0.5),
            "text":  (None, None, 0, "#1A1A1A", 0, 9.0, 0),
            "title": (None, None, 0, "#1F3864", 1, 10.0, 0),
            "note":  (None, None, 0, "#595959", 0, 8.0, 0),
            "label": ("#FFFFFF", None, 0, "#404040", 0, 8.0, 0),
            "band":  ("#2F5597", "#2F5597", 0, "#FFFFFF", 1, 9.0, 0.75),
        },
        "container_fill": "#F7F9FC", "container_line": "#8EA9C1",
        "conn_line": "#2E5C8A", "conn_dashed": "#7F7F7F", "label_text": "#404040",
    },
    "mono": {
        # 黑白印刷：全灰阶；ok/warn 只用填充深浅区分（浅灰=通过，深灰=异常）
        "styles": {
            "node":  ("#FFFFFF", "#000000", 0, "#000000", 0, 9.0, 0.75),
            "key":   ("#DDDDDD", "#000000", 0, "#000000", 1, 9.0, 0.75),
            "dark":  ("#000000", "#000000", 0, "#FFFFFF", 1, 9.0, 0.75),
            "ext":   ("#F2F2F2", "#808080", 1, "#404040", 0, 9.0, 0.75),
            "ok":    ("#E6E6E6", "#404040", 0, "#000000", 0, 9.0, 0.75),
            "warn":  ("#BFBFBF", "#000000", 0, "#000000", 0, 9.0, 0.75),
            "soft":  ("#F5F5F5", "#A6A6A6", 0, "#000000", 0, 8.5, 0.5),
            "text":  (None, None, 0, "#000000", 0, 9.0, 0),
            "title": (None, None, 0, "#000000", 1, 10.0, 0),
            "note":  (None, None, 0, "#595959", 0, 8.0, 0),
            "label": ("#FFFFFF", None, 0, "#404040", 0, 8.0, 0),
            "band":  ("#000000", "#000000", 0, "#FFFFFF", 1, 9.0, 0.75),
        },
        "container_fill": "#FAFAFA", "container_line": "#808080",
        "conn_line": "#000000", "conn_dashed": "#808080", "label_text": "#404040",
    },
}


def xml_palette(name="tech-cn"):
    """XML 后端用的色板（含 styles 与容器/连线颜色）；名字与 COM 主题共用注册表。"""
    if isinstance(name, Theme):
        name = name.name
    if name not in XML_PALETTES:
        raise KeyError("未知主题 %r，可选：%s" % (name, sorted(XML_PALETTES)))
    pal = XML_PALETTES[name]
    return {"styles": dict(pal["styles"]),
            "container_fill": pal["container_fill"], "container_line": pal["container_line"],
            "conn_line": pal["conn_line"], "conn_dashed": pal["conn_dashed"],
            "label_text": pal["label_text"]}


def xml_styles(name="tech-cn"):
    """XML 后端用的样式表（vlib STYLES 的格式），默认 tech-cn。"""
    return xml_palette(name)["styles"]


def get(name="tech-cn"):
    """按名字取主题；也可直接传入 Theme 实例。"""
    if isinstance(name, Theme):
        return name
    if name not in THEMES:
        raise KeyError("未知主题 %r，可选：%s" % (name, sorted(THEMES)))
    return THEMES[name]
