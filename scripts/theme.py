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

# 蓝色系 —— 汇报材料 / 幻灯片
BLUE = Theme(
    "blue", DEFAULT_CJK_FONT,
    black="RGB(0,0,0)",
    dark="RGB(31,41,55)",
    gray="RGB(75,85,99)",
    mgray="RGB(156,163,175)",
    lgray="RGB(243,244,246)",
    white="RGB(255,255,255)",
    accent="RGB(21,94,169)",
    accent_fill="RGB(219,234,254)",
    lane="RGB(239,246,255)",
    muted_fill="RGB(229,231,235)",
)

THEMES = {t.name: t for t in (TECH_CN, MONO, BLUE)}


def get(name="tech-cn"):
    """按名字取主题；也可直接传入 Theme 实例。"""
    if isinstance(name, Theme):
        return name
    if name not in THEMES:
        raise KeyError("未知主题 %r，可选：%s" % (name, sorted(THEMES)))
    return THEMES[name]
