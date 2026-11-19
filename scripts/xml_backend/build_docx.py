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

ap = argparse.ArgumentParser()
ap.add_argument('src'); ap.add_argument('figspecs'); ap.add_argument('templates'); ap.add_argument('out')
ap.add_argument('--preview', default=None)
A = ap.parse_args()
SRC, VS, OUT = A.src, A.templates, A.out
OUTDIR = A.preview or os.path.join(tempfile.gettempdir(), 'visio_kit_preview')
os.makedirs(OUTDIR, exist_ok=True)
DX = os.path.join(tempfile.gettempdir(), 'visio_kit_dx')
_spec = importlib.util.spec_from_file_location('figspecs', A.figspecs)
_m = importlib.util.module_from_spec(_spec); _spec.loader.exec_module(_m)
FIGS = _m.FIGS

shutil.rmtree(DX, ignore_errors=True)
with zipfile.ZipFile(SRC) as z:
    z.extractall(DX)
    order = z.namelist()

rels_p = f'{DX}/word/_rels/document.xml.rels'
rels = open(rels_p, encoding='utf8').read()
doc_p = f'{DX}/word/document.xml'
doc = open(doc_p, encoding='utf8').read()
rid_target = {}
for m in re.finditer(r'<Relationship [^>]*?/>', rels):
    r = m.group(0)
    rid_target[re.search(r'Id="([^"]+)"', r).group(1)] = re.search(r'Target="([^"]+)"', r).group(1)




removed = []
for name, suf, fn in FIGS:
    f = fn()
    png = f'fig{name}.png'
    f.render_png(f'{OUTDIR}/{png}', S=10)
    # --- vsdx：优先用该图原有嵌入的解包目录（保留原样式表），没有则用空白模板
    emb = f'Microsoft_Visio_Drawing{suf}.vsdx'
    tpl = os.path.join(VS, f'Microsoft_Visio_Drawing{suf}')
    if not os.path.isdir(tpl):
        tpl = BLANK
    build_vsdx(f, tpl, f'{DX}/word/embeddings/{emb}', os.path.join(tempfile.gettempdir(), 'visio_kit_work', f'D{suf}'))
    # --- docx: find the object that embeds this vsdx
    ole_rid = [k for k, v in rid_target.items() if v.endswith('embeddings/' + emb)][0]
    m = re.search(r'<w:object [^>]*>(?:(?!</w:object>).)*?r:id="' + ole_rid + r'"(?:(?!</w:object>).)*?</w:object>', doc, flags=re.S)
    obj = m.group(0)
    img_rid = re.search(r'<v:imagedata r:id="(rId\d+)"', obj).group(1)
    old_img = rid_target[img_rid]
    removed.append(old_img)
    rels = rels.replace(f'Target="{old_img}"', f'Target="media/{png}"')
    wpt, hpt = f.W / 25.4 * 72, f.H / 25.4 * 72
    new = re.sub(r'style="width:[^;"]*;height:[^;"]*"', f'style="width:{wpt:.2f}pt;height:{hpt:.2f}pt"', obj, count=1)
    new = re.sub(r'w:dxaOrig="\d+"', f'w:dxaOrig="{round(f.W/25.4*1440)}"', new)
    new = re.sub(r'w:dyaOrig="\d+"', f'w:dyaOrig="{round(f.H/25.4*1440)}"', new)
    new = new.replace('<o:lock v:ext="edit" aspectratio="f"/></v:shape>', '<o:lock v:ext="edit" aspectratio="t"/></v:shape>')
    doc = doc.replace(obj, new)
    shutil.copy(f'{OUTDIR}/{png}', f'{DX}/word/media/{png}')
    print(name, emb, img_rid, old_img, f'{f.W}x{f.H}mm')

open(rels_p, 'w', encoding='utf8').write(rels)
open(doc_p, 'w', encoding='utf8').write(doc)
for o in removed:
    if o not in rels:
        os.remove(f'{DX}/word/{o}')
ct_p = f'{DX}/[Content_Types].xml'
ct = open(ct_p, encoding='utf8').read()
if 'Extension="png"' not in ct:
    ct = ct.replace('<Default Extension="rels"', '<Default Extension="png" ContentType="image/png"/><Default Extension="rels"')
open(ct_p, 'w', encoding='utf8').write(ct)

out = OUT
files = []
for root, _, fs in os.walk(DX):
    for fn_ in fs:
        files.append(os.path.relpath(os.path.join(root, fn_), DX))
files.sort(key=lambda p: (p != '[Content_Types].xml', order.index(p) if p in order else 9999, p))
with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
    for p in files:
        z.write(os.path.join(DX, p), p)
print('written', out, os.path.getsize(out))
