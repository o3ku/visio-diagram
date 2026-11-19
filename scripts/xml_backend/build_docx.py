#!/usr/bin/env python3
"""用 figspecs 重绘 docx 里的全部嵌入 Visio 图，生成新 docx（原稿只读）。

用法：python3 build_docx.py <原稿.docx> <figspecs.py> <模板目录> <输出.docx> [--preview <PNG目录>]
  模板目录：extract_vsdx.py 从原稿解出的目录。
  figspecs.FIGS：[(图号, 嵌入后缀, 生成函数), ...]，后缀即 Microsoft_Visio_Drawing{后缀}.vsdx。
每张图：写入新 page1.xml 并按图尺寸改页面大小；docx 内预览图换成 PNG；对象尺寸按图尺寸设置并锁定纵横比。
"""
import argparse, importlib.util, os, re, shutil, sys, tempfile, zipfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from vsdxpkg import BLANK, build_vsdx


def load_figs(path):
    spec = importlib.util.spec_from_file_location('figspecs', path)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m.FIGS


def _read(p):
    with open(p, encoding='utf8') as fh:
        return fh.read()


def _write(p, s):
    with open(p, 'w', encoding='utf8') as fh:
        fh.write(s)


def build(src, figspecs, templates, out, preview=None):
    figs = load_figs(figspecs)
    tmp = tempfile.mkdtemp(prefix='visio_xml_')
    try:
        dx = os.path.join(tmp, 'docx')
        outdir = preview or os.path.join(tmp, 'preview')
        os.makedirs(outdir, exist_ok=True)
        with zipfile.ZipFile(src) as z:
            z.extractall(dx)
            order = z.namelist()

        rels_p = os.path.join(dx, 'word', '_rels', 'document.xml.rels')
        doc_p = os.path.join(dx, 'word', 'document.xml')
        rels, doc = _read(rels_p), _read(doc_p)
        rid_target = {}
        for r in re.findall(r'<Relationship [^>]*?/>', rels):
            rid_target[re.search(r'Id="([^"]+)"', r).group(1)] = re.search(r'Target="([^"]+)"', r).group(1)

        removed = []
        for name, suf, fn in figs:
            f = fn()
            png = f'fig{name}.png'
            f.render_png(os.path.join(outdir, png), S=10)
            # vsdx：优先用该图原有嵌入的解包目录（保留原样式表），没有则用空白模板
            emb = f'Microsoft_Visio_Drawing{suf}.vsdx'
            tpl = os.path.join(templates, f'Microsoft_Visio_Drawing{suf}')
            if not os.path.isdir(tpl):
                tpl = BLANK
            emb_path = os.path.join(dx, 'word', 'embeddings', emb)
            if not os.path.exists(emb_path):
                raise SystemExit(f'图 {name}：FIGS 里的后缀 {suf!r} 在 docx 中不存在（找不到 word/embeddings/{emb}）')
            build_vsdx(f, tpl, emb_path)
            # docx：找到嵌入这个 vsdx 的对象
            ole = [k for k, v in rid_target.items() if v.endswith('embeddings/' + emb)]
            if not ole:
                raise SystemExit(f'图 {name}：document.xml.rels 里没有指向 {emb} 的关系')
            m = re.search(r'<w:object [^>]*>(?:(?!</w:object>).)*?r:id="' + re.escape(ole[0]) +
                          r'"(?:(?!</w:object>).)*?</w:object>', doc, flags=re.S)
            if not m:
                raise SystemExit(f'图 {name}：document.xml 里没有引用 {ole[0]} 的 <w:object>')
            obj = m.group(0)
            mi = re.search(r'<v:imagedata [^>]*r:id="([^"]+)"', obj)
            if not mi:
                raise SystemExit(f'图 {name}：对象里没有 <v:imagedata> 预览')
            img_rid = mi.group(1)
            old_img = rid_target[img_rid]
            removed.append(old_img)
            rels = rels.replace(f'Target="{old_img}"', f'Target="media/{png}"')
            wpt, hpt = f.W / 25.4 * 72, f.H / 25.4 * 72
            new, n = re.subn(r'style="width:[^;"]*;height:[^;"]*', f'style="width:{wpt:.2f}pt;height:{hpt:.2f}pt', obj, count=1)
            if not n:
                raise SystemExit(f'图 {name}：v:shape 没有 width/height 样式')
            new = re.sub(r'w:dxaOrig="\d+"', f'w:dxaOrig="{round(f.W / 25.4 * 1440)}"', new)
            new = re.sub(r'w:dyaOrig="\d+"', f'w:dyaOrig="{round(f.H / 25.4 * 1440)}"', new)
            # 锁定纵横比：已有 o:lock 改为 t，没有则补一个
            if re.search(r'<o:lock\b[^>]*aspectratio=', new):
                new = re.sub(r'(<o:lock\b[^>]*aspectratio=)"[^"]*"', r'\1"t"', new)
            else:
                new, n = re.subn(r'(</v:shape>)', r'<o:lock v:ext="edit" aspectratio="t"/>\1', new, count=1)
                if not n:
                    raise SystemExit(f'图 {name}：找不到 </v:shape>，无法锁定纵横比')
            doc = doc.replace(obj, new)
            shutil.copy(os.path.join(outdir, png), os.path.join(dx, 'word', 'media', png))
            print(name, emb, img_rid, old_img, f'{f.W}x{f.H}mm')

        _write(rels_p, rels)
        _write(doc_p, doc)
        for o in removed:
            if f'Target="{o}"' not in rels:
                os.remove(os.path.join(dx, 'word', o))
        ct_p = os.path.join(dx, '[Content_Types].xml')
        ct = _read(ct_p)
        if 'Extension="png"' not in ct:
            ct = ct.replace('<Default Extension="rels"', '<Default Extension="png" ContentType="image/png"/><Default Extension="rels"')
        _write(ct_p, ct)

        files = []
        for root, _, fs in os.walk(dx):
            for fn_ in fs:
                files.append(os.path.relpath(os.path.join(root, fn_), dx).replace(os.sep, '/'))
        files.sort(key=lambda p: (p != '[Content_Types].xml', order.index(p) if p in order else 9999, p))
        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
            for p in files:
                z.write(os.path.join(dx, p), p)
        print('written', out, os.path.getsize(out))
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('src'); ap.add_argument('figspecs'); ap.add_argument('templates'); ap.add_argument('out')
    ap.add_argument('--preview', default=None, help='PNG 预览输出目录（默认临时目录，用完即删）')
    a = ap.parse_args(argv)
    build(a.src, a.figspecs, a.templates, a.out, a.preview)


if __name__ == '__main__':
    main()
