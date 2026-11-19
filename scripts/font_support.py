"""OS-specific font-family defaults and local font-file lookup.

Shared by scripts/theme.py (COM route) and scripts/xml_backend/vlib.py (XML route):
Latin/CJK family pairs per OS (Windows Arial + Microsoft YaHei, macOS Helvetica
Neue + PingFang SC, Linux via fontconfig), plus file resolution that never falls
back to an unrelated face. Override with VISIO_FONT_LATIN / VISIO_FONT_CJK and
VISIO_FONT_{LATIN,CJK}{,_BOLD}_FILE.
"""
import os
import platform
import shutil
import subprocess


def _fontconfig(family):
    exe = shutil.which("fc-match")
    if not exe:
        return None, None
    try:
        result = subprocess.run(
            [exe, "-f", "%{family}\\n%{file}", family],
            stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, text=True, timeout=2,
            check=False)
    except (OSError, subprocess.SubprocessError):
        return None, None
    lines = result.stdout.splitlines()
    if result.returncode or len(lines) < 2 or not os.path.isfile(lines[1]):
        return None, None
    return lines[0].split(",", 1)[0].strip(), lines[1]


def default_font_families(system=None):
    """Return (Latin family, CJK family), overridable by environment."""
    system = (system or platform.system()).lower()
    if system in ("windows", "win32"):
        latin, cjk = "Arial", "Microsoft YaHei"
    elif system in ("darwin", "mac", "macos"):
        latin, cjk = "Helvetica Neue", "PingFang SC"
    else:
        latin = _fontconfig("sans-serif:lang=en")[0] or "DejaVu Sans"
        cjk = _fontconfig("sans-serif:lang=zh-cn")[0] or "Noto Sans CJK SC"
    return (os.environ.get("VISIO_FONT_LATIN", latin),
            os.environ.get("VISIO_FONT_CJK", cjk))


DEFAULT_LATIN_FONT, DEFAULT_CJK_FONT = default_font_families()


def font_file(family, bold=False):
    """Resolve a family to an installed font file; never substitute a random face."""
    latin, cjk = default_font_families()
    role = "CJK" if family == cjk else "LATIN" if family == latin else None
    if role:
        suffix = "_BOLD_FILE" if bold else "_FILE"
        configured = os.environ.get("VISIO_FONT_" + role + suffix)
        if configured and os.path.isfile(configured):
            return configured

    name = family.casefold().replace(" ", "")
    windows = {
        "microsoftyahei": ("msyh.ttc", "msyhbd.ttc"), "微软雅黑": ("msyh.ttc", "msyhbd.ttc"),
        "arial": ("arial.ttf", "arialbd.ttf"),
    }
    mac = {
        "pingfangsc": ("/System/Library/Fonts/PingFang.ttc", "/System/Library/Fonts/PingFang.ttc"),
        "helveticaneue": ("/System/Library/Fonts/HelveticaNeue.ttc", "/System/Library/Fonts/HelveticaNeue.ttc"),
        "arial": ("/System/Library/Fonts/Supplemental/Arial.ttf", "/System/Library/Fonts/Supplemental/Arial Bold.ttf"),
    }
    linux = {
        "notosanscjksc": ("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc", "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc"),
        "notosanscjk": ("/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc", "/usr/share/fonts/truetype/noto/NotoSansCJK-Bold.ttc"),
        "dejavusans": ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
        "liberationsans": ("/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf", "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf"),
    }
    if os.name == "nt" and name in windows:
        path = os.path.join(os.environ.get("WINDIR", "C:/Windows"), "Fonts", windows[name][bool(bold)])
        if os.path.isfile(path):
            return path
    if platform.system() == "Darwin" and name in mac:
        if os.path.isfile(mac[name][bool(bold)]):
            return mac[name][bool(bold)]
    if platform.system() == "Linux" and name in linux:
        path = linux[name][bool(bold)]
        if os.path.isfile(path):
            return path

    query = family + (":style=Bold" if bold else ":style=Regular")
    if role == "CJK":
        query += ":lang=zh-cn"
    matched_family, path = _fontconfig(query)
    if matched_family and matched_family.casefold() == family.casefold() and path:
        return path
    return None
