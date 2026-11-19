#!/usr/bin/env python3
"""把 figspecs 里的图生成独立 .vsdx 文件和 PNG 预览（不经过 docx）。用于新图或单独交付。

用法：python3 build_vsdx.py <figspecs.py> <输出目录> [图号 ...]
模板：默认 assets/templates/blank_vsdx；可用 --template <解包目录> 指定。
"""
import argparse, importlib.util, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from vsdxpkg import BLANK, build_vsdx


def load(path):
    spec = importlib.util.spec_from_file_location('figspecs', path)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m.FIGS


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('figspecs'); ap.add_argument('outdir'); ap.add_argument('only', nargs='*')
    ap.add_argument('--template', default=BLANK)
    a = ap.parse_args()
    os.makedirs(a.outdir, exist_ok=True)
    for name, _suf, fn in load(a.figspecs):
        if a.only and name not in a.only:
            continue
        f = fn()
        build_vsdx(f, a.template, os.path.join(a.outdir, f'fig{name}.vsdx'), os.path.join(a.outdir, '.work', name))
        f.render_png(os.path.join(a.outdir, f'fig{name}.png'), S=10)
        print(name, f'{f.W}x{f.H}mm')
